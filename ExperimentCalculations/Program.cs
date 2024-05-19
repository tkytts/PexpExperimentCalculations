using ExperimentCalculations.Services;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


try
{
    var dataProcessingService = new DataProcessingService();
    dataProcessingService.Process();

    Console.WriteLine("Processamento finalizado com sucesso. Aperte enter para sair.");
    Console.Read();
}
catch (Exception exception)
{
    Console.WriteLine(exception.Message);
    Console.Read();
}
