using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsValue : IExcelCommand
    {
        public When When { get; private set; }
        public string Cells { get; private set; }
        public string Value { get; private set; }
        public CellValueKind ValueKind { get; private set; }

        public CellsValue(int row, int col, object value, CellValueKind valueKind = CellValueKind.Value)
            : this(ExcelService.GetAddress(row, col), value, valueKind) { }
        public CellsValue(int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, valueKind) { }
        public CellsValue(Address r, object value, CellValueKind valueKind = CellValueKind.Value)
            : this(ExcelService.GetAddress(r, 0, 0), value, valueKind) { }
        public CellsValue(Address r, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value)
            : this(ExcelService.GetAddress(r, row, col), value, valueKind) { }
        public CellsValue(Address r, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, valueKind) { }
        public CellsValue(string cells, object value, CellValueKind valueKind = CellValueKind.Value)
        {
            if (string.IsNullOrEmpty(cells))
                throw new ArgumentNullException(nameof(cells));
            When = When.Normal;
            Cells = cells;
            Value = value?.ToString();
            ValueKind = valueKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (CellValueKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.CellsValue(Cells, Value.ParseValue(), ValueKind);
    }
}