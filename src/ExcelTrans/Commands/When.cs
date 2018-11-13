using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum When : byte
    {
        Normal = 0,
        FirstSet = 1,
        First = 2,
        Before = 4,
        After = 8,
        Last = 16,
        LastSet = 32,
    }
}