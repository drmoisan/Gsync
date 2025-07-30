using log4net.Repository.Hierarchy;
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.Extensions;
using Gsync.Utilities.HelperClasses;
using Gsync.Utilities.Interfaces.IHelperClasses;

namespace Gsync
{

    public class AppFileSystemFolderPaths : IFileSystemFolderPaths
    {
        private readonly IEnvironment _environment;
        private readonly IDirectory _directory;

        public AppFileSystemFolderPaths() : this(new DefaultEnvironment(), new DefaultDirectory()) { }

        public AppFileSystemFolderPaths(IEnvironment environment) : this(environment, new DefaultDirectory()) { }

        public AppFileSystemFolderPaths(IEnvironment environment, IDirectory directory)
        {
            _environment = environment ?? throw new ArgumentNullException(nameof(environment));
            _directory = directory ?? throw new ArgumentNullException(nameof(directory));
            LoadFolders();
            _filenames = new AppStagingFilenames();
        }

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region ctor

        private AppFileSystemFolderPaths(bool async){}

        async public static Task<AppFileSystemFolderPaths> LoadAsync()
        {
            var fs = new AppFileSystemFolderPaths(true);
            await fs.LoadFoldersAsync();
            fs._filenames = new AppStagingFilenames();
            return fs;
        }

        #endregion ctor

        #region Methods
                
        private void CreateMissingPaths(string filepath)
        {
            if (!_directory.Exists(filepath))
            {
                _directory.CreateDirectory(filepath);
            }
        }

        async private Task CreateMissingPathsAsync(string filepath)
        {
            if (!_directory.Exists(filepath))
            {
                await Task.Run(() => _directory.CreateDirectory(filepath));
            }
        }

        public string MatchBestSpecialFolder(string path)
        {
            if (SpecialFolders.IsNullOrEmpty()) { return null; }
            var bestMatch = SpecialFolders.Where(x => path.Contains(x.Value)).OrderByDescending(x => x.Value.Length).FirstOrDefault();
            return bestMatch.Key;
        }

        private bool TryAddSpecialFolder(string name, string[] pathParts)
        {
            if (name.IsNullOrEmpty()) { return false; }
            
            else if (pathParts.IsNullOrEmpty())
            {
                logger.Debug($"Error in {nameof(TryAddSpecialFolder)} for key {nameof(name)} because {nameof(pathParts)} is null or empty. {TraceUtility.GetMyTraceString(new System.Diagnostics.StackTrace())}");
                return false;
            }
            
            else if (pathParts.Any(x => x is null || x.Trim().IsNullOrEmpty())) 
            {
                var locations = Enumerable.Range(0, pathParts.Length).Where(i => pathParts[i] is null).Select(i => i.ToString()).SentenceJoin();
                logger.Debug($"Error in {nameof(TryAddSpecialFolder)} for key {nameof(name)} because {nameof(pathParts)} has null elements at {locations}. {TraceUtility.GetMyTraceString(new System.Diagnostics.StackTrace())}");
                return false;
            }

            SpecialFolders ??= [];
            
            try
            {
                SpecialFolders[name] = Path.Combine(pathParts);
                CreateMissingPaths(SpecialFolders[name]);
                return true;
            }
            
            catch (Exception e)
            {
                logger.Error(e.Message, e);
                return false;
            }

        }

        private bool TryAddSpecialFolder(string name, Func<string[]> predicate)
        {
            try
            {
                var parts = predicate();
                return TryAddSpecialFolder(name, parts);
            }
            catch (Exception e)
            {
                logger.Error($"Error in {nameof(TryAddSpecialFolder)}. {nameof(predicate)} threw the following exception {e.Message}", e);
                return false;
            }
        }

        private void LoadFolders()
        {
            SpecialFolders = [];
            TryAddSpecialFolder("AppData", () => [_environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), nameof(Gsync)]);
            TryAddSpecialFolder("MyDocuments", () => [_environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)]);
            TryAddSpecialFolder("UserProfile", () => [_environment.GetFolderPath(Environment.SpecialFolder.UserProfile)]);
            TryAddSpecialFolder("MyComputer", () => [_environment.GetFolderPath(Environment.SpecialFolder.MyComputer)]);
            TryAddSpecialFolder("Favorites", () => [_environment.GetFolderPath(Environment.SpecialFolder.Favorites)]);
            TryAddSpecialFolder("Personal", () => [_environment.GetFolderPath(Environment.SpecialFolder.Personal)]);
            TryAddSpecialFolder("ApplicationData", () => [_environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)]);
            TryAddSpecialFolder("Desktop", () => [_environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)]);
            TryAddSpecialFolder("NetworkShortcuts", () => [_environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts)]);            
            if (!TryAddSpecialFolder("OneDrivePersonal", () => [_environment.GetEnvironmentVariable("OneDriveConsumer")])) 
            {
                TryAddSpecialFolder("OneDrivePersonal", () => [_environment.GetEnvironmentVariable("OneDrivePersonal")]);
            }
            if (!TryAddSpecialFolder("OneDrive", () => [_environment.GetEnvironmentVariable("OneDriveCommercial")])) 
            {
                if (!TryAddSpecialFolder("OneDrive", () => [_environment.GetEnvironmentVariable("OneDrive")])) 
                {
                    if (!TryAddSpecialFolder("OneDrive", () => [_environment.GetEnvironmentVariable("OneDrivePersonal")])) 
                    {
                        if(SpecialFolders.Count > 0) 
                        {
                            if(SpecialFolders.TryGetValue("AppData", out var appData))
                            {
                                TryAddSpecialFolder("OneDrive", [appData]);
                            }
                            else
                            {
                                TryAddSpecialFolder("OneDrive", [SpecialFolders.First().Value]);
                            }
                        }
                        else { throw new InvalidOperationException("No know network or local folders set in environment variables"); }
                    }
                }
            }
            SpecialFolders.TryGetValue("OneDrive", out var oneDrive);
            TryAddSpecialFolder("Flow", [oneDrive, "Email attachments from Flow"]);
            SpecialFolders.TryGetValue("Flow", out var flow);
            TryAddSpecialFolder("PreReads", [oneDrive, "_  Workflow", "_ Pre-Reads"]);
            TryAddSpecialFolder("System", () => [_environment.GetFolderPath(Environment.SpecialFolder.System)]);
            TryAddSpecialFolder("Root", () => [Path.GetPathRoot(_environment.GetFolderPath(Environment.SpecialFolder.System))]);
            
            if (SpecialFolders.TryGetValue("MyDocuments", out var myDocuments))
            {
                _remap = Path.Combine(myDocuments, "dictRemap.csv");
            }

            TryAddSpecialFolder("PythonStaging", [flow, "Combined", "data"]);            
        }

        //TODO: Cleanup Staging Files so that they are in one or two directories and not all over the place
        async private Task LoadFoldersAsync()
        {
            await Task.Run(LoadFolders);
        }

        public void Reload()
        {
            LoadFolders();
        }

        #endregion Methods

        #region Properties

        //private string _appData;
        //public string FldrAppData { get => _appData; protected set => _appData = value; }

        //private string _myDocuments;
        //public string FldrMyDocuments { get => _myDocuments; protected set => _myDocuments = value; }

        //private string _oneDrive;
        //public string FldrOneDrive { get => _oneDrive; protected set => _oneDrive = value; }

        //private string _flow;
        //public string FldrFlow { get => _flow; protected set => _flow = value; }

        //private string _prereads;
        //public string FldrPreReads { get => _prereads; protected set => _prereads = value; }

        //private string _fldrPythonStaging;
        //public string FldrPythonStaging { get => _fldrPythonStaging; protected set => _fldrPythonStaging = value; }

        private IAppStagingFilenames _filenames;
        public IAppStagingFilenames Filenames { get => _filenames; protected set => _filenames = value; }

        private ConcurrentDictionary<string, string> _specialFolders;
        public ConcurrentDictionary<string, string> SpecialFolders { get => _specialFolders; protected set => _specialFolders = value; }

        private string _remap;

        #endregion Properties
    }
}