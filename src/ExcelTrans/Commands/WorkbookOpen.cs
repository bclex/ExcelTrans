using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorkbookOpen : IExcelCommand
    {
        public When When { get; private set; }
        public string Path { get; private set; }
        public string Password { get; private set; }

        public WorkbookOpen(string path, string password = null)
        {
            When = When.Normal;
            Path = path ?? throw new ArgumentNullException(nameof(path));
            Password = password;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Path = r.ReadString();
            Password = r.ReadBoolean() ? r.ReadString() : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Path);
            w.Write(Password != null); if (Password != null) w.Write(Password);
        }

        void IExcelCommand.Execute(IExcelContext ctx)
        {
            var ctx2 = (ExcelContext)ctx;
            var pathFile = new FileInfo(Path);
            ctx2.OpenWorkbook(pathFile, Password);
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorkbookOpen: {Path}"); }
    }
}