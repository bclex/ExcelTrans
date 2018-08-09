using OfficeOpenXml;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsValue : IExcelCommand
    {
        public string Cells { get; private set; }
        public string[] Values { get; private set; }

        public CellsValue(int row, int col, params string[] values)
            : this(ExcelCellBase.GetAddress(row, col), values) { }
        public CellsValue(int fromRow, int fromCol, int toRow, int toCol, params string[] values)
            : this(ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol), values) { }
        public CellsValue(ExcelContext r, int plusRow, int plusCol, params string[] values)
            : this(ExcelCellBase.GetAddress(r.y, r.x, r.y + plusRow, r.x + plusCol), values) { }
        public CellsValue(string cells, params string[] values)
        {
            Cells = cells;
            Values = values;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Values = new string[r.ReadUInt16()];
            for (var i = 0; i < Values.Length; i++)
                Values[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((ushort)Values.Length);
            for (var i = 0; i < Values.Length; i++)
                w.Write(Values[i]);
        }

        void IExcelCommand.Execute(ExcelContext ctx)
        {
            var range = ctx.ws.Cells[Cells];
            range.Value = Values;
            //foreach (var v in Values)
            //{
            //    var y = long.TryParse(v, out var vl) ? vl :
            //        float.TryParse(v, out var vf) ? (object)vf :
            //        v;
            //    range.Value = y;
            //}
        }
    }
}