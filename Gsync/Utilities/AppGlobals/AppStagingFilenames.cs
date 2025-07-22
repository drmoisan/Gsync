using Gsync.Properties;
using Gsync.Utilities.Interfaces;

namespace Gsync.Utilities
{
    public class AppStagingFilenames : IAppStagingFilenames
    {
                
        private string _emailSessionTemp;
        public string EmailSessionTemp
        {
            get => _emailSessionTemp ?? InitProp(ref _emailSessionTemp, Properties.Settings.Default.FileName_EmailSessionTmp);
            
            set
            {
                Properties.Settings.Default.FileName_EmailSessionTmp = value;
                Properties.Settings.Default.Save();
            }
        }

        private string _emailSession;
        public string EmailSession
        {
            get => _emailSession ?? InitProp(ref _emailSession, Properties.Settings.Default.FileName_EmailSession);
            set
            {
                _emailSession = value;
                Properties.Settings.Default.FileName_EmailSession = value;
                Properties.Settings.Default.Save();
            }
        }

        private string _emailInfoStagingFile;
        public string EmailInfoStagingFile 
        { 
            get => _emailInfoStagingFile ?? InitProp(ref _emailInfoStagingFile, Settings.Default.FileName_EmailInfoStaging); 
            set => _emailInfoStagingFile = value; 
        }


        internal string InitProp(ref string prop, string value)
        {
            prop = value;
            return value;
        }
    }
}