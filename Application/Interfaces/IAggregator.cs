using Domain;

namespace Application.Interfaces;

public interface IAggregator
{
    Dictionary<int, Dictionary<string, int>> ByYearAndType(List<Vulnerability> vulnerabilities);
    Dictionary<string, int> TotalByType(Dictionary<int, Dictionary<string, int>> yearlyData);
}
