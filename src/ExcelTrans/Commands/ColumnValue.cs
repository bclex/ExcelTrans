using System.IO;

namespace ExcelTrans.Commands
{
    public struct ColumnValue : IExcelCommand
    {
        public When When { get; private set; }
        public int Col { get; private set; }
        public string Value { get; private set; }
        public ColumnValueKind ValueKind { get; private set; }

        public ColumnValue(int col, object value, ColumnValueKind valueKind)
        {
            When = When.Normal;
            Col = col;
            Value = value?.ToString();
            ValueKind = valueKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Col = r.ReadInt32();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (ColumnValueKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Col);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.ColumnValue(Col, Value.ParseValue(), ValueKind);
    }
}