namespace md2visio.struc.figure
{
    internal class GrowthDirection
    {
        public int H { get; set; } = 1;
        public int V { get; set; } = 0;
        public bool Positive { get => V > 0 || H > 0; }

        public GrowthDirection(Container figure)
        {
            Decide(figure);
        }

        public void Decide(Container figure)
        {            
            if (figure.Direction == "RL") 
                { H = -1; V = 0; }
            else if (figure.Direction == "TD" || figure.Direction == "TB") 
                { H = 0; V = -1; }
            else if (figure.Direction == "BT")
                { H = 0; V = 1; }
            else // LR
                { H = 1; V = 0; }
        }

        public override string ToString()
        {
            return $"H: {H}, V: {V}";
        }
    }
}
