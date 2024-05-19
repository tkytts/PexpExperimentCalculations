using ExperimentCalculations.Interfaces;
using ExperimentCalculations.Models;
using OfficeOpenXml;

namespace ExperimentCalculations.Services
{
    internal class Phase1CalculationService : ICalculationService
    {
        public int CalculatePhase(List<Session> sessions, ExcelWorksheet worksheet, int previousTotal)
        {
            worksheet.Cells[1, 1].Value = "Total de respostas";
            worksheet.Cells[1, 1].Style.Font.Bold = true;

            var totalResponses = sessions.Sum(s => s.Results.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia"));
            worksheet.Cells[2, 1].Value = totalResponses;

            return totalResponses;
        }
    }
}
