using ExcelTrans.Commands;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelTrans
{
    public interface IExcelContext
    {
        int XStart { get; set; }
        int X { get; set; }
        int DeltaX { get; set; }
        int Y { get; set; }
        int DeltaY { get; set; }
        int CsvX { get; set; }
        int CsvY { get; set; }
        Stack<Tuple<CommandRow[], CommandCol[]>> Cmds { get; }
        Stack<IExcelCommandSet> Sets { get; }
        object GetCtx();
        void SetCtx(object ctx);
    }

    internal class ExcelContext : IDisposable, IExcelContext
    {
        public ExcelContext()
        {
            P = new ExcelPackage();
            WB = P.Workbook;
        }
        public void Dispose() => P.Dispose();

        public int XStart { get; set; } = 1;
        public int X { get; set; } = 1;
        public int DeltaX { get; set; } = 1;
        public int Y { get; set; } = 1;
        public int DeltaY { get; set; } = 1;
        public int CsvX { get; set; } = 1;
        public int CsvY { get; set; } = 1;
        public Stack<Tuple<CommandRow[], CommandCol[]>> Cmds { get; } = new Stack<Tuple<CommandRow[], CommandCol[]>>();
        public Stack<IExcelCommandSet> Sets { get; } = new Stack<IExcelCommandSet>();
        public ExcelPackage P;
        public ExcelWorkbook WB;
        public ExcelWorksheet WS;

        public ExcelWorksheet EnsureWorksheet() => WS ?? (WS = WB.Worksheets.Add($"Sheet {WB.Worksheets.Count + 1}"));

        public object GetCtx() => new Tuple<int, int>(Cmds.Count, Sets.Count);
        public void SetCtx(object ctx)
        {
            var v = (Tuple<int, int>)ctx;
            PopCommand.Reset(this, v.Item1);
            PopSet.Reset(this, v.Item2);
        }
    }
}