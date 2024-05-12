using ExperimentCalculations.Enums;

namespace ExperimentCalculations.Models
{
    internal class Session
    {
        public required IEnumerable<Result> Results { get; set; }
        public required PhaseEnum Phase { get; set; }
        public required string Subject { get; set; }
    }
}
