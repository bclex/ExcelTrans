using ExcelTrans.Utils;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct PushFrame : IExcelCommand
    {
        public When When { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public PushFrame(params IExcelCommand[] cmds)
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

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after)
        {
            ctx.Frames.Push(ctx.Frame);
            ctx.ExecuteCmd(Cmds, out after); //action?.Invoke();
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PushFrame:"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }
    }
}