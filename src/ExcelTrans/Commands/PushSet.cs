using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public class PushSet : IExcelCommand, IExcelCommandSet
    {
        public int Headers { get; private set; }
        public Func<ExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>> Group { get; private set; }
        public Func<ExcelContext, object, IExcelCommand[]> Cmds { get; private set; }
        List<Collection<string>> _set;

        //public PushSet(Func<ExcelContext, IEnumerable<Collection<string>>, IEnumerable<TResult>> source, params IExcelCommand[] cmds)
        //{
        //    Source = source;
        //    Cmds = cmds;
        //}
        public PushSet(Func<ExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>> group, int headers = 1, Func<ExcelContext, IGrouping<string, Collection<string>>, IExcelCommand[]> cmds = null)
        {
            Headers = headers;
            Group = group;
            Cmds = (z, x) => cmds(z, (IGrouping<string, Collection<string>>)x);
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Headers = r.ReadByte();
            Group = ExcelContext.DecodeFunc<ExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>>(r);
            Cmds = ExcelContext.DecodeFunc<ExcelContext, object, IExcelCommand[]>(r);
            _set = new List<Collection<string>>();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)Headers);
            ExcelContext.EncodeFunc(w, Group);
            ExcelContext.EncodeFunc(w, Cmds);
        }

        void IExcelCommand.Execute(ExcelContext ctx) => ctx.sets.Push(this);

        void IExcelCommandSet.Add(Collection<string> s) => _set.Add(s);
        void IExcelCommandSet.Execute(ExcelContext ctx)
        {
            var headers = _set.Take(Headers).ToArray();
            if (Group != null)
                foreach (var g in Group(ctx, _set.Skip(Headers)))
                {
                    var si = ctx.Execute(Cmds(ctx, g));
                    foreach (var v in headers)
                        ExcelService.ProcessRow(ctx, v);
                    foreach (var v in g)
                        ExcelService.ProcessRow(ctx, v);
                    ctx.SetCtx(si);
                }
        }
    }
}