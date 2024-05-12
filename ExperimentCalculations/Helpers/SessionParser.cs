using ExperimentCalculations.Enums;
using ExperimentCalculations.Models;
using System.Globalization;

internal static class SessionParser
{
    public static IEnumerable<Session> ParseSessions(string subjectDirectory)
    {
        var sessions = new List<Session>();
        var directory = new DirectoryInfo(subjectDirectory);

        foreach (var sessionFile in directory.GetFiles())
        {
            if (sessionFile.Name.EndsWith(".timestamps"))
            {
                var sessionStream = sessionFile.OpenText();
                var firstLine = sessionStream.ReadLine() ?? throw new Exception($"Dados insuficientes no arquivo {sessionFile.Name}.");
                var splitName = firstLine.Split('.');
                var phaseName = splitName[2];
                var subject = splitName[1];
                var sessionPhase = (PhaseEnum)Enum.Parse(typeof(PhaseEnum), phaseName);

                var results = GetResult(sessionStream, sessionFile.Name, phaseName);

                if (sessions.Count != 0 && !sessions.Any(s => s.Subject == subject))
                    throw new Exception($"O arquivo de timestamp no caminho {sessionFile.FullName} tem um participante diferente de {sessions.First().Subject}. Nenhum cálculo foi feito, revise os arquivos e tente novamente.");

                sessions.Add(new Session
                {
                    Results = results,
                    Phase = sessionPhase,
                    Subject = subject
                });
            }
        }

        return sessions;
    }

    public static IEnumerable<Result> GetResult(StreamReader sessionStream, string fileName, string phaseName)
    {
        var resultText = sessionStream.ReadToEnd().Split(Environment.NewLine).Skip(5);
        var results = new List<Result>();
        if (resultText.Any())
        {
            Console.WriteLine($"Processando fase {phaseName} do arquivo {fileName}.");

            foreach (var line in resultText)
            {
                var result = new Result();

                if (!string.IsNullOrEmpty(line))
                {
                    var columns = line.Split("\t");

                    if (!float.TryParse(columns[0], CultureInfo.GetCultureInfoByIetfLanguageTag("pt"), out var timestamp))
                        break;

                    result.Timestamp = timestamp;
                    result.BlockID = int.Parse(columns[1]);
                    result.AttemptID = int.Parse(columns[2]);
                    result.Attempt = int.Parse(columns[3]);
                    result.AttemptName = columns[4];
                    result.Event = columns[5];

                    results.Add(result);
                }
            }

            return results;
        }
        else
            throw new Exception($"Dados insuficientes no arquivo {fileName}.");
    }
}