using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorksheetsDelete : IExcelCommand
    {
        public When When { get; private set; }
        public string Name { get; private set; }

        public WorksheetsDelete(string name)
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

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after)
        {
            var ctx2 = (ExcelContext)ctx;
            ctx2.WB.Worksheets.Delete(Name);
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetsDelete: {Name}"); }
    }
}