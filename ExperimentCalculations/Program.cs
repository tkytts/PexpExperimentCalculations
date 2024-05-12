using ExperimentCalculations.Services;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


try
{
    var dataProcessingService = new DataProcessingService();
    dataProcessingService.Process();

    Console.WriteLine("Processamento finalizado com sucesso. Aperte qualquer tecla para sair.");
    Console.ReadLine();
}
catch (Exception exception)
{
    Console.WriteLine(exception.Message);
    Console.ReadLine();
}
