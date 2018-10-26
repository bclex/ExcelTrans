using ExcelTrans.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public class PushSet : IExcelCommand, IExcelCommandSet
    {
        public When When { get; private set; }
        public int Headers { get; private set; }
        public Func<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>> Group { get; private set; }
        public Func<IExcelContext, object, IExcelCommand[]> Cmds { get; private set; }
        List<Collection<string>> _set;

        public PushSet(Func<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>> group, int headers = 1, Func<IExcelContext, IGrouping<string, Collection<string>>, IExcelCommand[]> cmds = null)
        {
            if (cmds == null)
                throw new ArgumentNullException(nameof(cmds));
            When = When.Normal;
            Headers = headers;
            Group = group;
            Cmds = (z, x) => cmds(z, (IGrouping<string, Collection<string>>)x);
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Headers = r.ReadByte();
            Group = ExcelSerDes.DecodeFunc<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<string, Collection<string>>>>(r);
            Cmds = ExcelSerDes.DecodeFunc<IExcelContext, object, IExcelCommand[]>(r);
            _set = new List<Collection<string>>();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)Headers);
            ExcelSerDes.EncodeFunc(w, Group);
            ExcelSerDes.EncodeFunc(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx) => ctx.Sets.Push(this);

        void IExcelCommandSet.Add(Collection<string> s) => _set.Add(s);
        void IExcelCommandSet.Execute(IExcelContext ctx)
        {
            var headers = _set.Take(Headers).ToArray();
            if (Group != null)
                foreach (var g in Group(ctx, _set.Skip(Headers)))
                {
                    ctx.WriteFirst(null);
                    var si = ctx.Execute(Cmds(ctx, g), out var after);
                    ctx.CsvY = 0;
                    foreach (var v in headers)
                    {
                        ctx.CsvY--;
                        ctx.WriteRow(v);
                    }
                    ctx.CsvY = 0;
                    foreach (var v in g)
                    {
                        ctx.CsvY++;
                        ctx.WriteRow(v);
                    }
                    after?.Invoke();
                    ctx.WriteLast(null);
                    ctx.SetCtx(si);
                }
        }
    }
}