using System;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public struct PushCommand : IExcelCommand
    {
        public IExcelCommand[] Cmds { get; private set; }

        public PushCommand(params IExcelCommand[] cmds)
        {
            Cmds = cmds;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cmds = ExcelContext.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelContext.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(ExcelContext ctx)
        {
            var rows = Cmds.Select(x => x as CommandRow).Where(x => x != null).ToArray();
            var cols = Cmds.Select(x => x as CommandCol).Where(x => x != null).ToArray();
            ctx.cmds.Push(new Tuple<CommandRow[], CommandCol[]>(rows, cols));
        }
    }
}