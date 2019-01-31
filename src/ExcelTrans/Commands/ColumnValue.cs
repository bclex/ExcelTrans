using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct ColumnValue : IExcelCommand
    {
        public When When { get; private set; }
        public int Col { get; private set; }
        public string Value { get; private set; }
        public ColumnValueKind ValueKind { get; private set; }
        public Type ValueType { get; set; }

        public ColumnValue(string col, object value, ColumnValueKind valueKind) : this(ExcelService.ColToInt(col), value, valueKind) { }
        public ColumnValue(int col, object value, ColumnValueKind valueKind)
        {
            When = When.Normal;
            Col = col;
            Value = value?.ToString();
            ValueKind = valueKind;
            ValueType = value?.GetType();
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Col = r.ReadInt32();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (ColumnValueKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Col);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.ColumnValue(Col, Value.CastValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ColumnValue[{Col}]: {Value}{$" - {ValueKind}"}"); }
    }
}