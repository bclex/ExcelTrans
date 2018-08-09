using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorksheetsAdd : IExcelCommand
    {
        public string Name { get; private set; }

        public WorksheetsAdd(string name)
        {
            Name = name;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
        }

        void IExcelCommand.Execute(ExcelContext ctx)
        {
            ctx.ws = ctx.wb.Worksheets.Add(Name);
            ctx.xstart = ctx.y = 1;
        }
    }
}