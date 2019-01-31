using ExcelTrans.Utils;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public class Command : IExcelCommand
    {
        public When When { get; private set; }
        public Action<IExcelContext> Action { get; private set; }
        public Command(Action action)
            : this(When.Before, v => action()) { }
        public Command(Action<IExcelContext> action)
            : this(When.Before, action) { }
        public Command(When when, Action action)
            : this(when, v => action()) { }
        public Command(When when, Action<IExcelContext> action)
        {
            When = when;
            Action = action ?? throw new ArgumentNullException(nameof(action));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            When = (When)r.ReadByte();
            Action = ExcelSerDes.DecodeAction<IExcelContext>(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)When);
            ExcelSerDes.EncodeAction(w, Action);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => Action(ctx);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Command{(When == When.Before ? null : $"[{When}]")}: [action]"); }
    }
}