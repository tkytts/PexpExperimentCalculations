namespace ExperimentCalculations.Models
{
    internal class Result
    {
        public double Timestamp { get; set; }
        public int BlockID { get; set; }
        public int AttemptID { get; set; }
        public int Attempt { get; set; }
        public string? AttemptName { get; set; }
        public string? Event { get; set; }
    }
}
