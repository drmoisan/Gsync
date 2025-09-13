
namespace Gsync.Utilities.Interfaces
{
    public interface IGettingThingsDone
    {
        #region Core GTD

        IFlagTranslator Context { get; }
        IFlagTranslator People { get; }
        IFlagTranslator Projects { get; }
        IFlagTranslator Topics { get; }

        #endregion Core GTD

        #region Extended GTD

        IFlagTranslator Program { get; }
        IFlagTranslator KB { get; }

        #endregion Extended GTD
    }
}
