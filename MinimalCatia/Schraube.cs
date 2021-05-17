namespace MinimalCatia
{
    internal class Schraube
    {
        public double laenge { get; private set; }
        public double gewindeLaenge { get; private set; }
        public string gewindeart { get; private set; }
        public string gewinde { get; private set; }
        public double k { get; private set; }
        public double Ri { get; private set; }
        public double P { get; private set; }

        public Schraube(double laenge, double gewindeLaenge, string gewindeart, 
            string gewinde, double k, double Ri, double P)
        {
            this.laenge = laenge;
            this.gewindeLaenge = gewindeLaenge;
            this.gewindeart = gewindeart;
            this.gewinde = gewinde;
            this.k = k;
            this.Ri = Ri;
            this.P = P;
        }
    }
}