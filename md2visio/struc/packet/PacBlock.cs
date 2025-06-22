using md2visio.struc.figure;
using md2visio.struc.graph;
using Microsoft.Office.Interop.Visio;

namespace md2visio.struc.packet
{
    internal class PacBlock : INode
    {
        public string ID { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
        public Shape? VisioShape { get; set; }
        public Container Container { get; set; } = Empty.Get<Container>();
        public int BitStart { get; set; } = 0;
        public int BitEnd { get; set; } = 0;
        public int BitNum { get => BitEnd - BitStart + 1; } 

        public PacBlock(string bits, string label) {
            string[] arr = bits.Split('-');
            (bool success, int val) = ParseInt(arr[0]);
            if(success && val >= 0) BitStart = BitEnd = val;

            if(arr.Length > 1) {
                (success, val) = ParseInt(arr[1]); 
                if(success && val >= BitStart) BitEnd = val;
            }

            ID = bits;
            Label = label;
        }

        public PacBlock Split(int headingBitNum)
        {
            if(headingBitNum >= BitNum) return this;

            int total = BitNum;
            BitEnd = BitStart + headingBitNum - 1;

            string bits = $"{BitEnd + 1}-{BitEnd + (total - headingBitNum)}";   
            return new PacBlock(bits, Label);
        }

        public List<GEdge> InputEdges => throw new NotImplementedException();

        public List<GEdge> OutputEdges => throw new NotImplementedException();

        (bool success, int val) ParseInt(string str)
        {
            int result = 0;
            bool success = int.TryParse(str.Trim(), out result);
            return (success, result);
        }

    }
}
