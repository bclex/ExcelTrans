using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopCommand : IExcelCommand
    {
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(ExcelContext ctx) => ctx.cmds.Pop();

        internal static void Reset(ExcelContext ctx, int idx)
        {
            while (ctx.cmds.Count > idx)
                ctx.cmds.Pop();
        }
    }
}