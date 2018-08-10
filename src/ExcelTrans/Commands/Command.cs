using ExcelTrans.Services;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public class Command : IExcelCommand
    {
        public Action<IExcelContext> Action { get; private set; }

        public Command(Action action)
            : this(v => action()) { }
        public Command(Action<IExcelContext> action)
        {
            Action = action;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Action = ExcelSerDes.DecodeAction<IExcelContext>(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelSerDes.EncodeAction(w, Action);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => Action(ctx);
    }
}