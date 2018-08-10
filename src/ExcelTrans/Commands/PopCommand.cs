using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopCommand : IExcelCommand
    {
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx) => ctx.Cmds.Pop();

        internal static void Reset(IExcelContext ctx, int idx)
        {
            while (ctx.Cmds.Count > idx)
                ctx.Cmds.Pop();
        }
    }
}