using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopCommand : IExcelCommand
    {
        public When When { get; private set; }
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx) => ctx.Cmds.Pop();

        internal static void Reset(IExcelContext ctx, int idx)
        {
            while (ctx.Cmds.Count > idx)
                ctx.Cmds.Pop();
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopCommand"); }
    }
}