

namespace Gsync.Utilities.Interfaces
{
    public interface IIDList:IScoDictionaryNew<string, int>
    {
        System.Action CompressIDs { get; set; }
        string GetNextID();
        string GetNextID(string seed);
        System.Action SynchronizeIDs { get; set; }
        string MaxID { get; }
    }
}
