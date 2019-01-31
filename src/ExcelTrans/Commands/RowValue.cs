using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct RowValue : IExcelCommand
    {
        public When When { get; private set; }
        public int Row { get; private set; }
        public string Value { get; private set; }
        public RowValueKind ValueKind { get; private set; }
        public Type ValueType { get; set; }

        public RowValue(string row, object value, RowValueKind valueKind) : this(ExcelService.RowToInt(row), value, valueKind) { }
        public RowValue(int row, object value, RowValueKind valueKind)
        {
            When = When.Normal;
            Row = row;
            Value = value?.ToString();
            ValueKind = valueKind;
            ValueType = value?.GetType();
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Row = r.ReadInt32();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (RowValueKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Row);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.RowValue(Row, Value.CastValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}RowValue[{Row}]: {Value}{$" - {ValueKind}"}"); }

    }
}