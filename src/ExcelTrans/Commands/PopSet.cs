using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopSet : IExcelCommand
    {
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(ExcelContext ctx) => ctx.sets.Pop().Execute(ctx);

        internal static void Reset(ExcelContext ctx, int idx)
        {
            while (ctx.sets.Count > idx)
                ctx.sets.Pop().Execute(ctx);
        }
    }
}