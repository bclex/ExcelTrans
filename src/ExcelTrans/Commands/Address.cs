using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum Address : ushort
    {
        Cell = 1,
        Range = 2,
        RowOrCol = 3,
        ColToCol = 4,
        RowToRow = 5,
        // Flags
        Rel = 0x10,
        //IncX = 0x20,
        //IncY = 0x40,
        // Mixture
        CellR = Cell | Rel,
        RangeR = Range | Rel,
    }
}