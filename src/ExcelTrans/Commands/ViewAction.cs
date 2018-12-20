using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct ViewAction : IExcelCommand
    {
        public When When { get; private set; }
        public string Value { get; private set; }
        public ViewActionKind ValueKind { get; private set; }
        public Type ValueType { get; set; }

        public ViewAction(int row, int col, ViewActionKind actionKind)
            : this(new Tuple<int, int>(row, col), actionKind) { }
        public ViewAction(object value, ViewActionKind actionKind)
        {
            When = When.Normal;
            Value = value?.ToString();
            ValueKind = actionKind;
            ValueType = value?.GetType();
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (ViewActionKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.ViewAction(Value.CastValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ViewAction: {Value} - {ValueKind}"); }
    }
}