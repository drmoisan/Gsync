using System;
using System.Collections.ObjectModel;

namespace Gsync.Utilities.Interfaces
{
    public interface IFlagTranslator
    {
        ObservableCollection<string> AsListNoPrefix { get; set; }
        ObservableCollection<string> AsListWithPrefix { get; set; }
        string AsStringNoPrefix { get; set; }
        string AsStringWithPrefix { get; set; }
        Func<bool, ObservableCollection<string>> GetListFunc { get; set; }
        Func<bool, string> GetStrFunc { get; set; }
        string Identifier { get; set; }
        Action<bool, ObservableCollection<string>> SetListFunc { get; set; }
        Action<bool, string> SetStrFunc { get; set; }

        string AsString();
        string ToString();
    }
}
