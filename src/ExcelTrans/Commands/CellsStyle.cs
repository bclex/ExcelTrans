using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsStyle : IExcelCommand
    {
        public When When { get; private set; }
        public string Cells { get; private set; }
        public string[] Styles { get; private set; }

        public CellsStyle(int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(row, col), styles) { }
        public CellsStyle(int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles) { }
        public CellsStyle(Address r, params string[] styles)
            : this(ExcelService.GetAddress(r, 0, 0), styles) { }
        public CellsStyle(Address r, int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(r, row, col), styles) { }
        public CellsStyle(Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles) { }
        public CellsStyle(string cells, params string[] styles)
        {
            if (string.IsNullOrEmpty(cells))
                throw new ArgumentNullException(nameof(cells));
            When = When.Normal;
            Cells = cells;
            Styles = styles ?? throw new ArgumentNullException(nameof(styles));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Styles = new string[r.ReadUInt16()];
            for (var i = 0; i < Styles.Length; i++)
                Styles[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((ushort)Styles.Length);
            for (var i = 0; i < Styles.Length; i++)
                w.Write(Styles[i]);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.CellsStyle(Cells, Styles);
    }
}