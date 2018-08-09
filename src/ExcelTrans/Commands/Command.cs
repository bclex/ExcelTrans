using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public class Command : IExcelCommand
    {
        public Action<ExcelContext> Action { get; private set; }

        public Command(Action action)
            : this(v => action()) { }
        public Command(Action<ExcelContext> action)
        {
            Action = action;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Action = ExcelContext.DecodeAction<ExcelContext>(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelContext.EncodeAction(w, Action);
        }

        void IExcelCommand.Execute(ExcelContext ctx) => Action(ctx);
    }
}