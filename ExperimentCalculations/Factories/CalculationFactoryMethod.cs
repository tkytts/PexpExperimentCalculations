using ExperimentCalculations.Enums;
using ExperimentCalculations.Interfaces;

namespace ExperimentCalculations.Factories
{
    internal abstract class CalculationFactoryMethod
    {
        public abstract ICalculationService CreateCalculationService(PhaseEnum phase);
    }
}
