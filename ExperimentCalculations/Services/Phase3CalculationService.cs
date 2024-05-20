using ExperimentCalculations.Enums;
using ExperimentCalculations.Interfaces;
using ExperimentCalculations.Models;
using ExperimentCalculations.Utils;
using OfficeOpenXml;
using System.Runtime.InteropServices.Marshalling;

namespace ExperimentCalculations.Services
{
    internal class Phase3CalculationService : ICalculationService
    {
        public int CalculatePhase(List<Session> sessions, ExcelWorksheet worksheet, int previousTotal)
        {
            int currentTotal = default;
            int currentColumn = 1;

            XlsxUtils.FillCell(worksheet, 1, currentColumn, "Pareamento", true);

            foreach (var session in sessions)
            {
                var sessionColumn = ++currentColumn;
                var sessionName = Enum.GetName(typeof(PhaseEnum), session.Phase);
                var previousSessionName = Enum.GetName(typeof(PhaseEnum), (PhaseEnum)((int)session.Phase - 1));

                XlsxUtils.FillCell(worksheet, 1, sessionColumn, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Coeficiente B", true);

                currentTotal = session.Results.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");

                var splitAResults = SplitAResults(session.Results);
                var splitBResults = SplitBResults(session.Results);
                var firstRow = 2;

                var lastRow = FillCoefficient1(splitAResults, splitBResults, firstRow, sessionColumn, worksheet);

                lastRow += 2;

                XlsxUtils.FillCell(worksheet, lastRow, sessionColumn, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Nº de respostas durante A", true);
                XlsxUtils.FillCell(worksheet, lastRow++, sessionColumn + 1, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Nº de respostas durante B", true);

                lastRow = FillCoefficient1Totals(splitAResults, splitBResults, lastRow, sessionColumn, worksheet);

                var aLineSplitResults = SplitALineResults(session.Results);
                var bLineSplitResults = SplitBLineResults(session.Results);

                XlsxUtils.FillCell(worksheet, 1, ++currentColumn, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Coeficiente B'", true);

                FillCoefficient2(aLineSplitResults, bLineSplitResults, firstRow, currentColumn, worksheet);

                lastRow += aLineSplitResults.Count + 1;

                XlsxUtils.FillCell(worksheet, lastRow, sessionColumn, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Média de respostas por segundo durante A'", true);
                XlsxUtils.FillCell(worksheet, lastRow++, sessionColumn + 1, $"{Enum.GetName(typeof(PhaseEnum), session.Phase)} - Média de respostas por segundo durante B'", true);

                FillCoefficient2Totals(aLineSplitResults, bLineSplitResults, lastRow, sessionColumn, worksheet);

                previousTotal = currentTotal;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            return currentTotal;
        }


        private static List<List<Result>> SplitAResults(IEnumerable<Result> results)
        {
            var sortedResults = results.OrderBy(r => r.Timestamp).ToList();
            var timestampSplits = results.Where(r => r.Event == "TelaCinza.Inicio").Select(r => r.Timestamp - 8).ToList();
            var splitResults = new List<List<Result>>();

            foreach (var timestamp in timestampSplits)
            {
                var timestampIndex = timestampSplits.IndexOf(timestamp);
                splitResults.Add(sortedResults.Where(r => r.Timestamp < timestamp + 8 && r.Timestamp >= timestamp).ToList());
            }

            return splitResults;
        }

        private static List<List<Result>> SplitBResults(IEnumerable<Result> results)
        {
            var sortedResults = results.OrderBy(r => r.Timestamp).ToList();
            var timestampSplits = results.Where(r => r.Event == "TelaCinza.Inicio").Select(r => r.Timestamp).ToList();
            var splitResults = new List<List<Result>>();

            foreach (var timestamp in timestampSplits)
            {
                var timestampIndex = timestampSplits.IndexOf(timestamp);
                splitResults.Add(sortedResults.Where(r => r.Timestamp < timestamp + 8 && r.Timestamp >= timestamp).ToList());
            }

            return splitResults;
        }

        private static List<List<Result>> SplitALineResults(IEnumerable<Result> results)
        {
            var sortedResults = results.OrderBy(r => r.Timestamp).ToList();
            var startBlockSplits = results.Where(r => r.Event == "TelaCinza.Inicio").Select(r => r.Timestamp).ToList();
            var endBlockSplits = results.Where(r => r.Event == "TomAlto.Fim").Select(r => r.Timestamp).ToList();

            if (endBlockSplits.Count == 0)
                endBlockSplits = results.Where(r => r.Event == "TelaCinza.Fim").Select(r => r.Timestamp).ToList();

            var splitResults = new List<List<Result>>();

            foreach (var timestamp in startBlockSplits)
            {
                var timestampIndex = startBlockSplits.IndexOf(timestamp);

                if (timestampIndex > 0)
                    splitResults.Add(sortedResults.Where(r => r.Timestamp < timestamp && r.Timestamp > endBlockSplits[timestampIndex - 1]).ToList());
                else
                    splitResults.Add(sortedResults.Where(r => r.Timestamp <= timestamp).ToList());
            }

            return splitResults;
        }

        private static List<List<Result>> SplitBLineResults(IEnumerable<Result> results)
        {
            var sortedResults = results.OrderBy(r => r.Timestamp).ToList();
            var timestampSplits = results.Where(r => r.Event == "TelaCinza.Inicio").Select(r => r.Timestamp).ToList();
            var splitResults = new List<List<Result>>();

            foreach (var timestamp in timestampSplits)
            {
                var timestampIndex = timestampSplits.IndexOf(timestamp);
                splitResults.Add(sortedResults.Where(r => r.Timestamp < timestamp + 8 && r.Timestamp >= timestamp).ToList());
            }

            return splitResults;
        }

        private static int FillCoefficient1(List<List<Result>> splitAResults, List<List<Result>> splitBResults, int firstRow, int sessionColumn, ExcelWorksheet worksheet)
        {
            var lastRow = firstRow;

            foreach (var aResults in splitAResults)
            {
                // A and B are supposed to have the same indexes, if they don't, the timestamp is wrong.
                var resultIndex = splitAResults.IndexOf(aResults);
                var cellRow = firstRow + resultIndex;

                XlsxUtils.FillCell(worksheet, cellRow, 1, $"P{1 + resultIndex}", true);

                var coefficient1 = CalculateCoefficient1(aResults, splitBResults[resultIndex]);

                XlsxUtils.FillCell(worksheet, cellRow, sessionColumn, coefficient1.ToString(), false);

                lastRow = cellRow;
            }

            return lastRow;
        }

        private static double CalculateCoefficient1(List<Result> aResults, List<Result> bResults)
        {
            var totalAResponses = aResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");
            var totalBResponses = bResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");
            return Math.Round(totalBResponses / (totalAResponses + (double)totalBResponses), 2);
        }

        private static int FillCoefficient1Totals(List<List<Result>> splitAResults, List<List<Result>> splitBResults, int firstRow, int sessionColumn, ExcelWorksheet worksheet)
        {
            var lastRow = firstRow;

            foreach (var aResults in splitAResults)
            {
                // A and B are supposed to have the same indexes, if they don't, the timestamp is wrong.
                var resultIndex = splitAResults.IndexOf(aResults);
                var cellRow = lastRow + resultIndex;

                XlsxUtils.FillCell(worksheet, cellRow, 1, $"P{1 + resultIndex}", true);

                var aTotal = aResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");
                var bTotal = splitBResults[resultIndex].Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia");

                XlsxUtils.FillCell(worksheet, cellRow, sessionColumn, aTotal.ToString(), false);
                XlsxUtils.FillCell(worksheet, cellRow, sessionColumn + 1, bTotal.ToString(), false);
            }

            return lastRow;
        }

        private static int FillCoefficient2(List<List<Result>> aLineSplitResults, List<List<Result>> bLineSplitResults, int firstRow, int currentColumn, ExcelWorksheet worksheet)
        {
            var lastRow = firstRow;

            foreach (var aLineResults in aLineSplitResults)
            {
                // A and B are supposed to have the same indexes, if they don't, the timestamp is wrong.
                var resultIndex = aLineSplitResults.IndexOf(aLineResults);
                var cellRow = firstRow + resultIndex;

                XlsxUtils.FillCell(worksheet, cellRow, 1, $"P{1 + resultIndex}", true);

                var coefficient2 = CalculateCoefficient2(aLineResults, bLineSplitResults[resultIndex]);

                XlsxUtils.FillCell(worksheet, cellRow, currentColumn, coefficient2.ToString(), false);

                lastRow = cellRow;
            }

            return lastRow;
        }

        private static int FillCoefficient2Totals(List<List<Result>> aLineSplitResults, List<List<Result>> bLineSplitResults, int firstRow, int sessionColumn, ExcelWorksheet worksheet)
        {
            var lastRow = firstRow;

            foreach (var aLineResults in aLineSplitResults)
            {
                // A and B are supposed to have the same indexes, if they don't, the timestamp is wrong.
                var resultIndex = aLineSplitResults.IndexOf(aLineResults);
                var cellRow = lastRow + resultIndex;

                XlsxUtils.FillCell(worksheet, cellRow, 1, $"P{1 + resultIndex}", true);

                var aTotal = aLineResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia") / (aLineResults.Max(r => r.Timestamp) - aLineResults.Min(r => r.Timestamp));
                var bTotal = bLineSplitResults[resultIndex].Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia") / (bLineSplitResults[resultIndex].Max(r => r.Timestamp) - bLineSplitResults[resultIndex].Min(r => r.Timestamp));

                XlsxUtils.FillCell(worksheet, cellRow, sessionColumn, Math.Round(aTotal, 2).ToString(), false);
                XlsxUtils.FillCell(worksheet, cellRow, sessionColumn + 1, Math.Round(bTotal, 2).ToString(), false);
            }

            return lastRow;
        }

        private static double CalculateCoefficient2(List<Result> splitAResults, List<Result> splitBResults)
        {
            var averageA = splitAResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia") / (splitAResults.Max(r => r.Timestamp) - splitAResults.Min(r => r.Timestamp));
            var averageB = splitBResults.Count(r => r.Event == "Quadrado.Resposta" || r.Event == "Quadrado.Resposta.Latencia") / (splitBResults.Max(r => r.Timestamp) - splitBResults.Min(r => r.Timestamp));
            return Math.Round(averageB / (averageA + (double)averageB), 2);
        }
    }
}

