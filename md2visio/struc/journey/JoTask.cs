namespace md2visio.struc.journey
{
    internal class JoTask
    {
        string name;
        int score = 5;
        HashSet<string> joiners = new HashSet<string>();

        public JoSection? Section { get; set; } = null;
        public string Name { get => name; }
        public int Score { get => score; }
        public HashSet<string> Joiners { get => joiners; }

        public JoTask(string name, string score, string partners)
        {
            this.name = name;
            if (int.TryParse(score, out int number)) this.score = number;

            foreach (string joiner in partners.Split(","))
            {
                joiners.Add(joiner);
            }
        }
    }
}
