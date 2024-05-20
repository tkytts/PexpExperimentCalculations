using ExperimentCalculations.Enums;
using ExperimentCalculations.Interfaces;
using ExperimentCalculations.Models;
using ExperimentCalculations.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExperimentCalculations.Services
{
    internal class Phase2CalculationService : ICalculationService
    {
        public int CalculatePhase(List<Session> sessions, ExcelWorksheet worksheet, int previousTotal)
        {
            int currentTotal = default;
            int currentColumn = 1;
            int previousResultCount = -1;


            XlsxUtils.FillCell(worksheet, 4, currentColumn, "Tempo", true);

            foreach (var session in sessions)
            {
                var sessionColumn = ++currentColumn;
                var sessionName = Enum.GetName(typeof(PhaseEnum), session.Phase);
                var previousSessionName = Enum.GetName(typeof(PhaseEnum), (PhaseEnum)((int)session.Phase - 1));

                XlsxUtils.FillCell(worksheet, 4, sessionColumn, $"Número total de respostas {Enum.GetName(typeof(PhaseEnum), session.Phase)}", true);

                var splitResults = SplitResults(session.Results);

                currentTotal = session.Results.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");
                var stability = ((currentTotal / (double)previousTotal) * 100) - 100;

                XlsxUtils.FillCell(worksheet, 1, currentColumn, $"Estabilidade {previousSessionName} comparado com {sessionName}", true);

                var stabilityValueCell = worksheet.Cells[2, currentColumn];

                stabilityValueCell.Value = $"{Math.Round(stability, 2)} %";
                stabilityValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                stabilityValueCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (stability >= -15 && stability <= 15)
                    stabilityValueCell.Style.Fill.BackgroundColor.SetColor(color: Color.LightGreen);
                else
                    stabilityValueCell.Style.Fill.BackgroundColor.SetColor(color: Color.Red);

                foreach (var row in splitResults)
                {
                    var rowIndex = splitResults.IndexOf(row);
                    var cellRow = 5 + rowIndex;
                    var currentCount = row.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");

                    XlsxUtils.FillCell(worksheet, cellRow, 1, $"T{1 + rowIndex}", true);
                    XlsxUtils.FillCell(worksheet, cellRow, sessionColumn, currentCount.ToString(), false);

                    if (previousResultCount >= 0)
                    {
                        var intraSessionStability = (currentCount / (double)previousResultCount) * 100;
                        var intraSessionCell = worksheet.Cells[cellRow, sessionColumn + 1];

                        XlsxUtils.FillCell(worksheet, cellRow, sessionColumn + 1, $"{Math.Round(intraSessionStability - 100, 2)} %", false);

                        previousResultCount = currentCount;
                    }
                    else
                    {
                        XlsxUtils.FillCell(worksheet, 4, sessionColumn + 1, "Estabilidade intra-tempos", true);
                        previousResultCount = currentCount;
                    }

                }
                var sumCellRow = splitResults.Count + 5;

                XlsxUtils.FillCell(worksheet, 4, sessionColumn + 2, $"Média da variação entre todos os tempos {sessionName}", true);
                XlsxUtils.FillCell(worksheet, 5, sessionColumn + 2, Math.Round((session.Results.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia")) / (double)splitResults.Count, 2).ToString(), false);
                XlsxUtils.FillCell(worksheet, sumCellRow, currentColumn - 1, $"Total de respostas {sessionName}", true);
                XlsxUtils.FillCell(worksheet, sumCellRow, sessionColumn, session.Results.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia").ToString(), false);

                var minSumRow = sumCellRow + 2;
                var maxSumRow = minSumRow + 1;

                XlsxUtils.FillCell(worksheet, minSumRow, currentColumn - 1, $"Limite mínimo do número de respostas (com base em {previousSessionName})", true);
                XlsxUtils.FillCell(worksheet, minSumRow, sessionColumn, Math.Round(previousTotal * 0.85, 0).ToString(), false);

                XlsxUtils.FillCell(worksheet, maxSumRow, currentColumn - 1, $"Limite máximo do número de respostas (com base em {previousSessionName})", true);
                XlsxUtils.FillCell(worksheet, maxSumRow, sessionColumn, Math.Round(previousTotal * 1.15, 0).ToString(), false);

                previousTotal = currentTotal;
                previousResultCount = -1;
                currentColumn += 3;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            return currentTotal;
        }
        private static List<List<Result>> SplitResults(IEnumerable<Result> results)
        {
            var sortedResults = results.OrderBy(r => r.Timestamp).ToList();
            var resultGroups = new List<List<Result>>();
            var minTimestamp = sortedResults.First().Timestamp;
            var maxTimestamp = sortedResults.Last().Timestamp;
            var totalSeconds = maxTimestamp - minTimestamp;
            var intervalSeconds = 8; // Desired interval between groups in seconds

            if (sortedResults.Count != 0)
            {
                var groupIndex = 0;
                var currentGroup = new List<Result>();
                var groupStartTime = minTimestamp;

                foreach (var result in sortedResults)
                {
                    var currentTimestamp = result.Timestamp;

                    if (currentTimestamp >= groupStartTime + intervalSeconds)
                    {
                        resultGroups.Add(currentGroup);
                        currentGroup = [];
                        groupIndex++;
                        groupStartTime = minTimestamp + groupIndex * intervalSeconds;
                    }

                    currentGroup.Add(result);
                }

                if (currentGroup.Count != 0)
                {
                    resultGroups.Add(currentGroup);
                }
            }

            return resultGroups;
        }
    }
}
