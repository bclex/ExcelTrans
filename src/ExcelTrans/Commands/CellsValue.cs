using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsValue : IExcelCommand
    {
        public string Cells { get; private set; }
        public string Value { get; private set; }
        public bool Formula { get; private set; }

        public CellsValue(int row, int col, object value, bool formula = false)
            : this(ExcelService.GetAddress(row, col), value, formula) { }
        public CellsValue(int fromRow, int fromCol, int toRow, int toCol, object value, bool formula = false)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formula) { }
        public CellsValue(Address r, object value, bool formula = false)
            : this(ExcelService.GetAddress(r, 0, 0), value, formula) { }
        public CellsValue(Address r, int row, int col, object value, bool formula = false)
            : this(ExcelService.GetAddress(r, row, col), value, formula) { }
        public CellsValue(Address r, int fromRow, int fromCol, int toRow, int toCol, object value, bool formula = false)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formula) { }
        public CellsValue(string cells, object value, bool formula = false)
        {
            Cells = cells;
            Value = value.ToString();
            Formula = formula;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Value = r.ReadString();
            Formula = r.ReadBoolean();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write(Value);
            w.Write(Formula);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.CellsValue(Cells, ExcelService.ParseValue(Value), Formula);
    }
}