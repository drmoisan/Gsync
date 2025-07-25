using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using Gsync.Utilities.Extensions;
using Gsync.Utilities.HelperClasses.NewtonSoft;

namespace Gsync.Utilities.HelperClasses
{
    public class FilePathHelper: INotifyPropertyChanged, ICloneable
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region Constructors

        public FilePathHelper() { PropertyChanged += FilePathHelper_PropertyChanged; }

        public FilePathHelper(string fileName, string folderPath)
        {
            FileName = fileName;
            FolderPath = folderPath;
            FilePath = Path.Combine(_folderPath, _fileName);
            PropertyChanged += FilePathHelper_PropertyChanged;
        }

        private FilePathHelper(string fileNameSeed, string fileExtension, string fileNameSuffix, string folderPath)
        {
            FileStemSeed = fileNameSeed;
            FileStemSuffix = fileNameSuffix;
            FileExtension = fileExtension;
            FolderPath = folderPath;
            FilePath = Path.Combine(_folderPath, _fileName);
            PropertyChanged += FilePathHelper_PropertyChanged;
        }

        public static FilePathHelper FromSeed(string fileNameSeed, string fileExtension, string fileNameSuffix, string folderPath)
        {
            var fph = new FilePathHelper(fileNameSeed, fileExtension, fileNameSuffix, folderPath);
            
            return fph;
        }

        #endregion Constructors

        #region Public Properties

        private string _filePath = "";
        public string FilePath { get => _filePath; set { _filePath = value; NotifyPropertyChanged(); } }
        
        private string _folderPath = "";
        public string FolderPath { get=> _folderPath; set { _folderPath = value; NotifyPropertyChanged(); } }

        private string _fileName = "";
        public string FileName { get => _fileName; set { _fileName = value; NotifyPropertyChanged(); }}

        private string _fileStemSeed = null;
        public string FileStemSeed { get => _fileStemSeed; set { _fileStemSeed = value; NotifyPropertyChanged(); } }

        private string _fileStemSuffix = null;
        public string FileStemSuffix { get => _fileStemSuffix; set { _fileStemSuffix = value; NotifyPropertyChanged(); } }

        private string _fileStem = null;
        public string FileStem { get => _fileStem; protected set => _fileStem = value; }

        private string _fileExtension = null;
        public string FileExtension { get => _fileExtension; set { _fileExtension = value; NotifyPropertyChanged(); } }

        public const int MAX_PATH = 256;

        #endregion Public Properties

        #region Public Methods

        public virtual bool Exists()
        {
            if(FilePath.IsNullOrEmpty())
                return false;
            try
            {
                return File.Exists(FilePath);
            }
            catch (Exception e)
            {
                logger.Error($"{e.StackTrace}\n{e.Message}\nError checking if file exists for FilePath {FilePath}");
                return false;
            }
        }

        public virtual DateTime GetLastWriteTimeUtc()
        {
            if (Exists())
                try
                {
                    return File.GetLastWriteTimeUtc(FilePath);
                }
                catch (Exception e)
                {
                    logger.Error($"{e.StackTrace}\n{e.Message}\nError getting last write time for FilePath {FilePath}");
                    return default;                    
                }
            else
                return default;
        }

        internal bool StemInitialized()
        {
            if (FileStemSeed is null || FileStemSuffix is null || FileExtension is null)
            {
                if (FolderPath.IsNullOrEmpty() || !TryParseFileName(FileName))
                    return false;
            }
            return !FolderPath.IsNullOrEmpty();
        }

        public int CalcMaxSeedLength()
        {
            if (!StemInitialized())
                return MAX_PATH;
            
            return MAX_PATH - FolderPath.Length - FileExtension.Length - FileStemSuffix.Length;
        }

        public (string Stem, string Extension) ExtractStemAndExtension(string fileName)
        {
            var fileExtension = Path.GetExtension(fileName);
            var fileStem = Path.GetFileNameWithoutExtension(fileName);
            
            if (fileStem.IsNullOrEmpty())
            {
                fileStem = fileExtension ?? "";
                fileExtension = "";
            }

            return (fileStem, fileExtension ?? "");
        }
        
        public bool TryParseFileStem(string fileStem, out string fileStemSeed, out string fileStemSuffix)
        {
            fileStemSeed = FileStemSeed ?? "";
            fileStemSuffix = FileStemSuffix ?? "";
            string remainingChars = fileStem;

            // case 1: empty fileStem
            if (fileStem.IsNullOrEmpty())
                return false;

            // case 2: fileStemSeed is empty AND fileStemSuffix is empty
            if (fileStemSeed.IsNullOrEmpty() && fileStemSuffix.IsNullOrEmpty())
            {
                fileStemSeed = fileStem;
                remainingChars = null;
            }

            // case 3: Some part of seed or suffix remains
            // step 1 strip existing seed if it exists
            if (!fileStemSeed.IsNullOrEmpty() && (remainingChars?.StartsWith(fileStemSeed) ?? false))
                remainingChars = remainingChars.Replace(fileStemSeed, "");

            // step 2 strip existing suffix if it exists and append any remaining chars to seed
            if (!fileStemSuffix.IsNullOrEmpty() && (remainingChars?.EndsWith(fileStemSuffix) ?? false))
            {
                remainingChars = remainingChars.Replace(fileStemSuffix, "");
                fileStemSeed += remainingChars;
            }

            // step 3 append any remaining
            if (fileStemSeed.IsNullOrEmpty())
            {
                fileStemSeed = remainingChars;
                fileStemSuffix = "";
            }
            else if (fileStemSuffix.IsNullOrEmpty())
            {
                fileStemSuffix = remainingChars;
            }
            else 
            {
                fileStemSeed += remainingChars;
            }

            return true;
        }

        public bool TryParseFileName(string fileName)
        {
            if (fileName.IsNullOrEmpty())
                return false;
            
            var (fileStem, fileExtension) = ExtractStemAndExtension(fileName);
                        
            if (TryParseFileStem(fileStem, out string fileStemSeed, out string fileStemSuffix))
            {
                _fileStemSeed = fileStemSeed;
                _fileStemSuffix = fileStemSuffix;
                _fileStem = fileStem;
                _fileExtension = fileExtension;
                return true;
            }
            else
            {
                return false;
            }
        }
                
        public bool AdjustForMaxPath()
        {
            if (!StemInitialized())
                return false;

            var maxSeedLength = MAX_PATH - FolderPath.Length - FileExtension.Length - FileStemSuffix.Length;
            
            var fileName = $"{FileStemSeed}{FileStemSuffix}{FileExtension}";
            var filePath = Path.Combine(FolderPath, fileName);
            if (filePath.Length > MAX_PATH)
            {
                maxSeedLength = FileStemSeed.Length + MAX_PATH - filePath.Length;
                _fileStemSeed = FileStemSeed.Substring(0, maxSeedLength);
            }
            return true;
        }

        public static string AdjustForMaxPath(string folderPath, string filenameSeed, string fileExtension, string filenameSuffix = "")
        {
            var filename = $"{filenameSeed}{filenameSuffix}{fileExtension}";
            var filepath = Path.Combine(folderPath, filename);
            if (filepath.Length >= MAX_PATH)
            {
                var maxSeedLength = filenameSeed.Length + MAX_PATH - filepath.Length;
                filenameSeed = filenameSeed.Substring(0, maxSeedLength);
                filename = $"{filenameSeed.Substring(0, maxSeedLength)}{filenameSuffix}{fileExtension}";
                filepath = Path.Combine(folderPath, filename);
            }
            return filepath;
        }

        #endregion Public Methods

        #region INotifyPropertyChanged Implementation

        private void NotifyPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void FilePathHelper_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "FileName":
                    if (!_folderPath.IsNullOrEmpty() && !_fileName.IsNullOrEmpty())
                        _filePath = Path.Combine(_folderPath, _fileName);
                    else
                        _filePath = null;
                    break;
                case "FolderPath":
                    if (!_folderPath.IsNullOrEmpty() && !_fileName.IsNullOrEmpty())
                        _filePath = Path.Combine(_folderPath, _fileName);
                    else
                        _filePath = null;
                    break;
                
                case "FilePath":
                    if (_filePath.IsNullOrEmpty())
                    {
                        _folderPath = "";
                        _fileName = "";
                        return;
                    }                    
                    
                    try
                    {
                        _folderPath = Path.GetDirectoryName(_filePath);    
                    }
                    catch (System.Exception ex)
                    {
                        var st = string.Join("\n",TraceUtility.GetMyMethodNames(new StackTrace()));
                        string msg = $"FilePath: {_filePath} is invalid.\n{st}";
                        logger.Error(msg, ex);
                        throw;
                    }
                    
                    _fileName = Path.GetFileName(_filePath);
                    if (_fileName == "") 
                        throw new ArgumentException($"FilePath {_filePath} must include a FileName");
                    break;
                case "FileStemSeed":
                    if (AdjustForMaxPath())
                    {
                        _fileStem = $"{_fileStemSeed}{_fileStemSuffix}";
                        FileName = $"{_fileStem}{_fileExtension}";
                    }
                    break;
                case "FileStemSuffix":
                    if (AdjustForMaxPath())
                    {
                        _fileStem = $"{_fileStemSeed}{_fileStemSuffix}";
                        FileName = $"{_fileStem}{_fileExtension}";
                    }
                    break;
                case "FileExtension":
                    if (AdjustForMaxPath())
                    {
                        _fileStem = $"{_fileStemSeed}{_fileStemSuffix}";
                        FileName = $"{_fileStem}{_fileExtension}";
                    }
                    break;

                default:
                    break;
            }
        }

        #endregion INotifyPropertyChanged Implementation

        #region ICloneable Implementation

        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public FilePathHelper DeepCopy()
        {
            //var clone = new FilePathHelper();
            var clone = new FilePathHelper
            {
                _folderPath = _folderPath,
                _fileName = _fileName,
                _filePath = _filePath,
                _fileStemSeed = _fileStemSeed,
                _fileStemSuffix = _fileStemSuffix,
                _fileStem = _fileStem,
                _fileExtension = _fileExtension
            };
            return clone;
        }

        public void CopyFrom(FilePathHelper other)
        {
            _folderPath = other._folderPath;
            _fileName = other._fileName;
            _filePath = other._filePath;
            _fileStemSeed = other._fileStemSeed;
            _fileStemSuffix = other._fileStemSuffix;
            _fileStem = other._fileStem;
            _fileExtension = other._fileExtension;
        }

        internal IList<string> CopyChanged(FilePathHelper other)
        {
            var changed = new List<string>();
            if (_folderPath != other._folderPath)
            {
                _folderPath = other._folderPath;
                changed.Add(nameof(FolderPath));
            }

            if (_fileName != other._fileName)
            {
                _fileName = other._fileName;
                changed.Add(nameof(FileName));
            }
            
            if (_filePath != other._filePath)
            {
                _filePath = other._filePath;
                changed.Add(nameof(FilePath));
            }
            if (_fileStemSeed != other._fileStemSeed)
            {
                _fileStemSeed = other._fileStemSeed;
                changed.Add(nameof(FileStemSeed));
            }
            if (_fileStemSuffix != other._fileStemSuffix)
            {
                _fileStemSuffix = other._fileStemSuffix;
                changed.Add(nameof(FileStemSuffix));
            }
            if (_fileStem != other._fileStem)
            {
                _fileStem = other._fileStem;
                changed.Add(nameof(FileStem));
            }
            if (_fileExtension != other._fileExtension)
            {
                _fileExtension = other._fileExtension;
                changed.Add(nameof(FileExtension));
            }
            return changed;
        }

        #endregion ICloneable Implementation
    }
}
