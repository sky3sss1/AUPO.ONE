using Application.Interfaces;
using Domain;

namespace Infrastructure.Services;

public class Aggregator : IAggregator
{
    public Dictionary<int, Dictionary<string, int>> ByYearAndType(List<Vulnerability> vulnerabilities)
    {
        var groupedData = new Dictionary<int, Dictionary<string, int>>();

        foreach (var vuln in vulnerabilities)
        {
            int year = vuln.Date.Year;
            if (!groupedData.ContainsKey(year))
                groupedData[year] = new Dictionary<string, int>();

            string[] types = vuln.Type.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var type in types)
            {
                if (!groupedData[year].ContainsKey(type))
                    groupedData[year][type] = 0;

                groupedData[year][type]++;
            }
        }

        return groupedData;
    }

    public Dictionary<string, int> TotalByType(Dictionary<int, Dictionary<string, int>> yearlyData)
    {
        var totalByType = new Dictionary<string, int>();

        foreach (var yearData in yearlyData.Values)
        {
            foreach (var kvp in yearData)
            {
                if (!totalByType.ContainsKey(kvp.Key))
                    totalByType[kvp.Key] = 0;

                totalByType[kvp.Key] += kvp.Value;
            }
        }

        return totalByType;
    }
}
