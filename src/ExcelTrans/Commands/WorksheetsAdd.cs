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

        void IExcelCommand.Execute(IExcelContext ctx)
        {
            var ctx2 = (ExcelContext)ctx;
            ctx2.WS = ctx2.WB.Worksheets.Add(Name);
            ctx.XStart = ctx.Y = 1;
        }
    }
}