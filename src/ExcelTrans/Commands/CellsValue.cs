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
        public Type ValueType { get; set; }

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
            ValueType = value?.GetType();
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (CellValueKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellsValue(Cells, Value.CastValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellsValue[{ExcelService.DescribeAddress(Cells)}]: {Value}{(ValueKind == CellValueKind.Value ? null : $" - {ValueKind}")}"); }
    }
}