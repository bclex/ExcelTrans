using ExcelTrans.Utils;
using System;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public struct PushCommand : IExcelCommand
    {
        public When When { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public PushCommand(params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Cmds = cmds ?? throw new ArgumentNullException(nameof(cmds));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cmds = ExcelSerDes.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelSerDes.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx)
        {
            var rows = Cmds.Select(x => x as CommandRow).Where(x => x != null).ToArray();
            var cols = Cmds.Select(x => x as CommandCol).Where(x => x != null).ToArray();
            ctx.Cmds.Push(new Tuple<CommandRow[], CommandCol[]>(rows, cols));
        }
    }
}