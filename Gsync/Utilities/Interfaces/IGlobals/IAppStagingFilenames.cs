
namespace Gsync.Utilities.Interfaces
{
    public interface IAppStagingFilenames
    {        
        string EmailSession { get; set; }
        string EmailSessionTemp { get; set; }        
        string EmailInfoStagingFile { get; set; }
    }
}