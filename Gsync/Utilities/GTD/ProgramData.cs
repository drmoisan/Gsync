using System;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.ReusableTypes;

namespace Gsync.Utilities.GTD
{
    // Implements the internal IProgramData interface, which inherits from IScoDictionaryNew<IProjectElement, int>
    internal class ProgramData : ScoDictionaryNew<IProjectElement, int>, IProgramData
    {
        // You can add additional constructors or members here if needed.
        // By default, this class inherits all functionality from ScoDictionaryNew.
    }
}