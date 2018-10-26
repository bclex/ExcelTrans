using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandRow : IExcelCommand
    {
        public When When { get; private set; }
        public Func<IExcelContext, Collection<string>, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public CommandRow(When when, Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        public CommandRow(When when, Func<CommandRtn> func, Func<IExcelCommand[]> cmds)
            : this(when, (a, b) => func(), a => cmds()) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, Func<IExcelContext, IExcelCommand[]> cmds)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = null;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            When = (When)r.ReadByte();
            Func = ExcelSerDes.DecodeFunc<IExcelContext, Collection<string>, CommandRtn>(r);
            Cmds = ExcelSerDes.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)When);
            ExcelSerDes.EncodeFunc(w, Func);
            ExcelSerDes.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx) { }
    }
}