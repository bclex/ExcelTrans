using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandRow : IExcelCommand
    {
        public Func<ExcelContext, Collection<string>, CommandRtn> Predicate { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        public CommandRow(Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b) => func(), cmds) { }
        public CommandRow(Func<ExcelContext, Collection<string>, CommandRtn> predicate, params IExcelCommand[] cmds)
        {
            Predicate = predicate;
            Cmds = cmds;
        }
        public CommandRow(Func<CommandRtn> func, Func<IExcelCommand[]> cmds)
            : this((a, b) => func(), a => cmds()) { }
        public CommandRow(Func<ExcelContext, Collection<string>, CommandRtn> predicate, Func<ExcelContext, IExcelCommand[]> cmds)
        {
            Predicate = predicate;
            Cmds = null;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Predicate = ExcelContext.DecodeFunc<ExcelContext, Collection<string>, CommandRtn>(r);
            Cmds = ExcelContext.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelContext.EncodeFunc(w, Predicate);
            ExcelContext.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(ExcelContext ctx) { }
    }
}