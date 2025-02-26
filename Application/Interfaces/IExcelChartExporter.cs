using Domain;

namespace Application.Interfaces;

public interface IExcelChartExporter
{
    byte[] GenerateExcelWithChartsFirstTask(List<Vulnerability> vulnerabilities);
    byte[] GenerateExcelWithChartsSecondTask(List<Vulnerability> vulnerabilities);
    byte[] GenerateExcelWithChartsThirdTask(List<Vulnerability> vulnerabilities);
}
