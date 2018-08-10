using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum WhenRow : byte
    {
        FirstRow = 1,
        BeforeRow = 2,
        AfterRow = 4,
        LastRow = 8,
    }
}