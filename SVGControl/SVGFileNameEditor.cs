using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace SVGControl
{
    internal class SvgFileNameEditor : FileNameEditor
    {
        private string _currentValue = string.Empty;
        private string _absoluteFilepath = string.Empty;
        private string _fileName = string.Empty;
        private string _appPath;
        private OpenFileDialog _ofd;

        public SvgFileNameEditor()
        {
            SetDevLevelPath();
        }

        public override object EditValue(ITypeDescriptorContext context, 
                                         IServiceProvider provider, 
                                         object value)
        {
            if (value is string)
            {
                _currentValue = (string)value;
                _absoluteFilepath = (string)value;
                _fileName = (string)value;
                if (_currentValue[0].Equals('.')) 
                {
                    _absoluteFilepath = RelativePath.AbsoluteFromURI(
                        uriToMakeAbsolute: _currentValue, anchorPath: _appPath);
                    _fileName = Path.GetFileName(_absoluteFilepath);
                }
            }
            else
            {
                _absoluteFilepath = string.Empty;
            }
            if (_ofd != null)
            {
                _ofd.InitialDirectory = Path.GetDirectoryName(_absoluteFilepath);
            }
            return base.EditValue(context, provider, _fileName);
        }

        protected override void InitializeDialog(OpenFileDialog ofd)
        {
            base.InitializeDialog(ofd);
            if (!_absoluteFilepath.Equals(string.Empty))
                ofd.InitialDirectory = Path.GetDirectoryName(_absoluteFilepath);
            ofd.Filter = "Vector Graphics(*.svg) | *.svg";
            _ofd = ofd;
        }

        private void SetDevLevelPath()
        {
            string workingDirectory = Environment.CurrentDirectory;
            List<string> directories = new List<string>(workingDirectory.Split(Path.DirectorySeparatorChar));
            if ((directories.Count>2)&&(directories[directories.Count - 2] == "bin"))
            {
                // Backwards traverse 2 levels
                _appPath = Directory.GetParent(workingDirectory).Parent.FullName;
            }
            else { _appPath = workingDirectory; }
        }
    }
}
