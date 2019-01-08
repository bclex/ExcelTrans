using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopFrame : IExcelCommand
    {
        public When When { get; private set; }
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx) => ctx.Frame = ctx.Frames.Pop();

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopFrame"); }
    }
}