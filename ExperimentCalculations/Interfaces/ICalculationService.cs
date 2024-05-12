using ExperimentCalculations.Models;
using OfficeOpenXml;

namespace ExperimentCalculations.Interfaces
{
    internal interface ICalculationService
    {
        int CalculatePhase(List<Session> sessions, ExcelWorksheet worksheet, int previousTotal);
    }
}
