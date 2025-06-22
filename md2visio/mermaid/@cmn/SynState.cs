using md2visio.struc.figure;
using System.Reflection;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.cmn
{
    internal abstract class SynState 
    {
        public static SynState EndOfFile = new EmptyState();

        SynContext context = Empty.Get<SynContext>();
        protected CompoDict compoList = new CompoDict();
        protected GroupCollection ExpectedGroups { get { return Ctx.ExpectedGroups; } }
        public SynContext Ctx { get { return context; } set { context = value; } }

        public string Fragment { get { return compoList.Entire; } set { compoList.Entire = value; } }
        public CompoDict CompoList { get { return compoList; } }
        public string Buffer { get { return Ctx.Cache.ToString(); } }

        public SynState() { }

        public SynState ClearBuffer()
        {
            Ctx.ClearCache();
            return this;
        }

        public string? Peek(int length = 1)
        {
            return Ctx.Peek(length);
        }

        public SynState SlideSpaces()
        {
            string? next = Ctx.Peek();
            while (next != null)
            {
                if (next == "\n") break;
                if (!string.IsNullOrWhiteSpace(next)) break;
                Ctx.Slide();
                next = Ctx.Peek();
            }

            return this;
        }

        public SynState Take(int length = 1)
        {
            Ctx.Take(length);
            return this;
        }

        public SynState Slide(int length = 1)
        {
            Ctx.Slide(length);
            return this;
        }

        protected bool Expect(string pattern, bool multiline = false)
        {
            return Ctx.Expect(pattern, multiline);
        }

        public bool Until(string pattern, bool multiline = true)
        {
            return Ctx.Until(pattern, multiline);
        }


        public SynState Restore(int length = 1)
        {
            Ctx.Restore(length);
            return this;
        }

        public string ExpectedGroup(string groupName)
        {
            if (ExpectedGroups.Count > 0) return ExpectedGroups[groupName].Value;
            return string.Empty;
        }

        public abstract SynState NextState();

        public SynState Forward<T>() where T : SynState, new()
        {
            return Create<T>().NextState();
        }

        public SynState Forward(Type sttType)
        {
            MethodInfo? forwardInfo = GetType().GetMethod("Forward", BindingFlags.Public | BindingFlags.Instance, 
                null, new Type[] { }, null);
            if (forwardInfo == null) throw new MissingMethodException("could not find the method 'Forward'");

            MethodInfo genericMethod = forwardInfo.MakeGenericMethod(sttType);
            object? rst = genericMethod.Invoke(this, null);

            if (rst == null) throw new NullReferenceException("target state can't be null");

            return (SynState) rst;
        }

        public SynState Create<T>() where T : SynState, new()
        {
            SynState clone = new T();
            clone.Ctx = Ctx;
            return clone;
        }

        public SynState Save(string frag)
        {
            compoList.Entire = frag;
            Ctx.AddState(this);
            return this;
        }

        public bool IsEndOfFile() { return this == EndOfFile; }

        public string GetPart(string name)
        {
            return compoList.Get(name);
        }

        public string GetPart(int index)
        {
            return compoList.Get(index);
        }

        public void AddCompo(string name, string value)
        {
            compoList.Add(name, value);
        }

        public void AddCompo(string value)
        {
            compoList.Add(value);
        }

        public override string ToString()
        {
            if (this is SttFinishFlag) return "■";
            return string.Format("{0} \t@ {1}", compoList.Entire, GetType().Name);
        }

    } // SynState

    internal class EmptyState : SynState
    {
        public static EmptyState Instance { get; } = Empty.Get<EmptyState>();

        public EmptyState() { }

        public override SynState NextState()
        {
            throw new NotImplementedException();
        }
    }
}
