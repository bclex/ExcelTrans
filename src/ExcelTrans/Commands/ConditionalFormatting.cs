using Newtonsoft.Json;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct ConditionalFormatting : IExcelCommand
    {
        public When When { get; private set; }
        public string Address { get; private set; }
        public string Value { get; private set; }
        public ConditionalFormattingKind FormattingKind { get; private set; }
        public int? Priority { get; private set; }
        public bool StopIfTrue { get; private set; }
        public string[] Styles { get; private set; }

        public ConditionalFormatting(int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
            : this(ExcelService.GetAddress(row, col), value, formattingKind, priority, stopIfTrue, styles) { }
        public ConditionalFormatting(int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue, styles) { }
        public ConditionalFormatting(Address r, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
            : this(ExcelService.GetAddress(r, 0, 0), value, formattingKind, priority, stopIfTrue, styles) { }
        public ConditionalFormatting(Address r, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
            : this(ExcelService.GetAddress(r, row, col), value, formattingKind, priority, stopIfTrue, styles) { }
        public ConditionalFormatting(Address r, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue, styles) { }
        public ConditionalFormatting(string address, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false, params string[] styles)
        {
            if (string.IsNullOrEmpty(address))
                throw new ArgumentNullException(nameof(address));
            When = When.Normal;
            Address = address;
            Value = value != null ? value is string ? (string)value : JsonConvert.SerializeObject(value) : null;
            FormattingKind = formattingKind;
            Priority = priority;
            StopIfTrue = stopIfTrue;
            Styles = styles;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Address = r.ReadString();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            FormattingKind = (ConditionalFormattingKind)r.ReadInt32();
            Priority = r.ReadBoolean() ? (int?)r.ReadInt32() : null;
            StopIfTrue = r.ReadBoolean();
            // Styles
            var styles = new string[r.ReadUInt16()];
            for (var i = 0; i < styles.Length; i++)
                styles[i] = r.ReadString();
            Styles = styles;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Address);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)FormattingKind);
            w.Write(Priority != null); if (Priority != null) w.Write(Priority.Value);
            w.Write(StopIfTrue);
            // Styles
            w.Write((ushort)(Styles?.Length ?? 0));
            if (Styles != null)
                foreach (var style in Styles)
                    w.Write(style);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.ConditionalFormatting(Address, Value, FormattingKind, Priority, StopIfTrue, Styles);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ConditionalFormatting[{Address}]: {Value} - {FormattingKind}"); }
    }
}