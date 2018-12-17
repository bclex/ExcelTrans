using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum Address : ushort
    {
        Cell = 1,
        Range = 2,
        ColOrRow = 3,
        // Flags
        IncX = 0x10,
        IncY = 0x20,
        //
        CellX1 = Cell | IncX,
        CellY1 = Cell | IncY,
    }
}