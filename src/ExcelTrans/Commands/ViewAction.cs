using System.IO;

namespace ExcelTrans.Commands
{
    public struct ViewAction : IExcelCommand
    {
        public When When { get; private set; }
        public string Value { get; private set; }
        public ViewActionKind ActionKind { get; private set; }

        public ViewAction(int row, int col, ViewActionKind actionKind)
            : this(ExcelService.GetAddress(row, col), actionKind) { }
        public ViewAction(string value, ViewActionKind actionKind)
        {
            When = When.Normal;
            Value = value;
            ActionKind = actionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ActionKind = (ViewActionKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ActionKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.ViewAction(Value, ActionKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ViewAction: {Value} - {ActionKind}"); }
    }
}