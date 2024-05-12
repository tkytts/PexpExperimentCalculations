using ExperimentCalculations.Enums;
using ExperimentCalculations.Factories;
using ExperimentCalculations.Models;
using ExperimentCalculations.Utils;
using OfficeOpenXml;

namespace ExperimentCalculations.Services
{
    internal class DataProcessingService
    {
        private static readonly string BASE_DIRECTORY = AppDomain.CurrentDomain.BaseDirectory + "/Participantes/";
        private static readonly PhaseEnum[] PHASE_2_ENUMS = [PhaseEnum.F1, PhaseEnum.F2, PhaseEnum.F3];
        private static readonly PhaseEnum[] PHASE_3_ENUMS = [PhaseEnum.SC, PhaseEnum.SC1, PhaseEnum.SC2];
        private static readonly PhaseEnum[] PHASE_4_ENUMS = [PhaseEnum.SC, PhaseEnum.SC1, PhaseEnum.SC2, PhaseEnum.R];

        private readonly CalculationFactory _calculationFactory;
        public DataProcessingService()
        {
            _calculationFactory = new CalculationFactory();
        }

        public void Process()
        {
            var directories = Directory.GetDirectories(BASE_DIRECTORY);

            foreach (var directory in directories)
            {
                using var excelPackage = XlsxUtils.CreateExcelPackage();
                var workbook = excelPackage.Workbook;
                var sessions = SessionParser.ParseSessions(directory).OrderBy(s => s.Phase).ToList();

                var phaseTotalResponses = this.CalculatePhase(workbook, sessions.Where(s => s.Phase == PhaseEnum.A).ToList(), PhaseEnum.A, 0);
                phaseTotalResponses = this.CalculatePhase(workbook, sessions.Where(s => PHASE_2_ENUMS.Contains(s.Phase)).ToList(), PhaseEnum.F1, phaseTotalResponses);
                phaseTotalResponses = this.CalculatePhase(workbook, sessions.Where(s => PHASE_3_ENUMS.Contains(s.Phase)).ToList(), PhaseEnum.SC, phaseTotalResponses);
                phaseTotalResponses = this.CalculatePhase(workbook, sessions.Where(s => PHASE_4_ENUMS.Contains(s.Phase)).ToList(), PhaseEnum.R, phaseTotalResponses);

                excelPackage.SaveAs(directory + "/" + sessions.First().Subject + ".xlsx");
            }
        }

        private int CalculatePhase(ExcelWorkbook workbook, List<Session> sessions, PhaseEnum phase, int phaseTotalResponses)
        {
            ExcelWorksheet worksheet = phase switch
            {
                PhaseEnum.A => workbook.Worksheets.Add("Aquisição"),
                PhaseEnum.F1 => workbook.Worksheets.Add("Fortalecimento"),
                PhaseEnum.SC => workbook.Worksheets.Add("Supressão Condicionada"),
                PhaseEnum.R => workbook.Worksheets.Add("Regra"),
                PhaseEnum.F2 or PhaseEnum.F3 => workbook.Worksheets["Fortalecimento"],
                PhaseEnum.SC1 or PhaseEnum.SC2 or PhaseEnum.SC3 => workbook.Worksheets["Supressão Condicionada"],
                _ => throw new NotImplementedException(),
            };

            var calculationService = _calculationFactory.CreateCalculationService(phase);

            return calculationService.CalculatePhase(sessions, worksheet, phaseTotalResponses);
        }
    }
}
