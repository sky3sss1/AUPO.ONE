using ClosedXML.Excel;
using Domain;
using Application.Interfaces;

namespace Infrastructure.Services;

public class Parser : IParser
{
    private int skipYears = 6;
    public List<Vulnerability> MapFromByte(byte[] byteArray)
    {
        using var stream = new MemoryStream(byteArray);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(1);

        var vulnerabilities = new List<Vulnerability>(100000);
        var sixYearsAgo = DateTime.Now.AddYears(-skipYears);
        var rows = worksheet.RowsUsed().Skip(1);

        foreach (var row in rows)
        {
            try
            {
                string GetCellValue(int index) => row.Cell(index + 1).GetValue<string>().Trim();

                DateTime parsedDate;
                if (!row.Cell(10).TryGetValue(out parsedDate))
                {
                    if (!DateTime.TryParse(GetCellValue(9), out parsedDate))
                    {
                        Console.WriteLine($"Ошибка даты в строке {row.RowNumber()} - {GetCellValue(9)}");
                        continue;
                    }
                }

                if (parsedDate < sixYearsAgo)
                {
                    Console.WriteLine($"Пропущена строка {row.RowNumber()} - старая дата ({parsedDate:yyyy-MM-dd})");
                    continue;
                }

                if (parsedDate.Year == DateTime.Now.Year)
                {
                    Console.WriteLine($"Пропущена строка {row.RowNumber()} - 2025 ({parsedDate:yyyy-MM-dd})");
                    continue;
                }

                var vulnerability = new Vulnerability(
                    id: GetCellValue(0),
                    name: GetCellValue(1),
                    description: GetCellValue(2),
                    vendor: GetCellValue(3),
                    poName: GetCellValue(4),
                    version: GetCellValue(5),
                    type: GetCellValue(6),
                    oSName: GetCellValue(7),
                    @class: GetCellValue(8),
                    date: parsedDate,
                    cVSS2: GetCellValue(10),
                    cVSS3: GetCellValue(11),
                    dangerousLevel: GetCellValue(12),
                    measures: GetCellValue(13),
                    status: GetCellValue(14),
                    exploit: GetCellValue(15),
                    information: GetCellValue(16),
                    source: GetCellValue(17),
                    anotherIdentities: GetCellValue(18),
                    anotherInfo: GetCellValue(19),
                    iBConnection: row.Cell(21).TryGetValue<int>(out var ibConn) ? ibConn : 0,
                    useWay: GetCellValue(21),
                    defenseWay: GetCellValue(22),
                    errorDescription: GetCellValue(23),
                    errorType: GetCellValue(24)
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
