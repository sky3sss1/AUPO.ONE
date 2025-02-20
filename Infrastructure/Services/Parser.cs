using ClosedXML.Excel;
using Domain;
using Application.Interfaces;

namespace Infrastructure.Services;

public class Parser : IParser
{
    public List<Vulnerability> MapFromByte(byte[] byteArray)
    {
        using var stream = new MemoryStream(byteArray);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var vulnerabilities = new List<Vulnerability>(100000);
        var sixYearsAgo = DateTime.Now.AddYears(-6);
        var rows = worksheet.RowsUsed().Skip(1);

        foreach (var row in rows)
        {
            try
            {
                var cells = row.Cells(0, 24).ToList();
                DateTime parsedDate = DateTime.Now;
                if (!cells[9]?.TryGetValue<DateTime>(out parsedDate) ?? true)
                {
                    if (DateTime.TryParse(cells[9]?.GetString(), out var manualDate))
                    {
                        parsedDate = manualDate;
                    }
                    else
                    {
                        Console.WriteLine($"Ошибка даты в строке {row.RowNumber()} - {cells[9]?.GetString()}");
                        continue;
                    }
                }

                if (parsedDate < sixYearsAgo)
                {
                    Console.WriteLine($"Пропущена строка {row.RowNumber()} - старая дата ({parsedDate:yyyy-MM-dd})");
                    continue;
                }

                var vulnerability = new Vulnerability(
                    id: cells[0]?.GetString() ?? "",
                    name: cells[1]?.GetString() ?? "",
                    description: cells[2]?.GetString() ?? "",
                    vendor: cells[3]?.GetString() ?? "",
                    poName: cells[4]?.GetString() ?? "",
                    version: cells[5]?.GetString() ?? "",
                    type: cells[6]?.GetString() ?? "",
                    oSName: cells[7]?.GetString() ?? "",
                    @class: cells[8]?.GetString() ?? "",
                    date: parsedDate,
                    cVSS2: cells[10]?.GetString() ?? "",
                    cVSS3: cells[11]?.GetString() ?? "",
                    dangerousLevel: cells[12]?.GetString() ?? "",
                    measures: cells[13]?.GetString() ?? "",
                    status: cells[14]?.GetString() ?? "",
                    exploit: cells[15]?.GetString() ?? "",
                    information: cells[16]?.GetString() ?? "",
                    source: cells[17]?.GetString() ?? "",
                    anotherIdentities: cells[18]?.GetString() ?? "",
                    anotherInfo: cells[19]?.GetString() ?? "",
                    iBConnection: cells[20]?.TryGetValue<int>(out var ibConn) == true ? ibConn : 0,
                    useWay: cells[21]?.GetString() ?? "",
                    defenseWay: cells[22]?.GetString() ?? "",
                    errorDescription: cells[23]?.GetString() ?? "",
                    errorType: cells[24]?.GetString() ?? ""
                );

                vulnerabilities.Add(vulnerability);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка обработки строки {row.RowNumber()}: {ex.Message}");
            }
        }

        Console.WriteLine($"Загружено {vulnerabilities.Count} уязвимостей.");
        return vulnerabilities;
    }
}
