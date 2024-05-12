using ExperimentCalculations.Enums;
using ExperimentCalculations.Interfaces;
using ExperimentCalculations.Services;

namespace ExperimentCalculations.Factories
{
    internal class CalculationFactory : CalculationFactoryMethod
    {
        public override ICalculationService CreateCalculationService(PhaseEnum phase)
        {
            return phase switch
            {
                PhaseEnum.A => new Phase1CalculationService(),
                PhaseEnum.F1 or PhaseEnum.F2 or PhaseEnum.F3 => new Phase2CalculationService(),
                PhaseEnum.SC or PhaseEnum.SC1 or PhaseEnum.SC2 or PhaseEnum.SC3 or PhaseEnum.R => new Phase3CalculationService(),
                _ => throw new NotImplementedException(),
            };
        }
    }
}
