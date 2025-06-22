using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace md2visio.mermaid.cmn
{
    internal class SttIterator
    {
        SynContext ctx;
        int stateIndex = -1;

        public int Pos { get { return stateIndex; } }
        public SynContext Context { get { return ctx; } }

        public SttIterator(SynContext ctx)
        {
            this.ctx = ctx;
        }

        public SynState Current
        {
            get
            {
                return ctx.StateList[stateIndex];
            }
        }

        public bool HasNext()
        {
            return ctx.StateList.Count > stateIndex + 1;
        }

        public SynState Next()
        {
            return ctx.StateList[++stateIndex];
        }

        public SynState PeekNext()
        {
            return ctx.StateList[stateIndex + 1];
        }

        public bool HasPre()
        {
            return stateIndex - 1 >= 0 &&
                stateIndex - 1 < ctx.StateList.Count;
        }

        public SynState PeekPre()
        {
            return ctx.StateList[stateIndex - 1];
        }
    }
}
