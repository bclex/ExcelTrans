using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum When : byte
    {
        Normal = 0,
        First = 1,
        Before = 2,
        After = 4,
        Last = 8,
    }
}