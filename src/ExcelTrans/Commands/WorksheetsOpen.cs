using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorksheetsOpen : IExcelCommand
    {
        public When When { get; private set; }
        public string Name { get; private set; }

        public WorksheetsOpen(string name)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
        }

        void IExcelCommand.Execute(IExcelContext ctx)
        {
            var ctx2 = (ExcelContext)ctx;
            ctx2.WS = ctx2.WB.Worksheets[Name];
            ctx.XStart = ctx.Y = 1;
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetsOpen: {Name}"); }
    }
}