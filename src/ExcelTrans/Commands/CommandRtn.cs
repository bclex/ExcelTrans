using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum CommandRtn
    {
        None = 0,
        Skip = 1,
        Formula = 2,
        Execute = 4,
    }
}