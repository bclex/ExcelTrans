using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct Flush : IExcelCommand
    {
        public When When { get; private set; }
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Flush();

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Flush"); }
    }
}