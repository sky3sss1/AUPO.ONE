using Application.Interfaces;
using ClosedXML.Excel;
using Domain;

namespace Infrastructure.Services
{
    public class ExcelChartExporter : IExcelChartExporter
    {
        private readonly IAggregator _aggregator;

        public ExcelChartExporter(IAggregator aggregator)
        {
            _aggregator = aggregator ?? throw new ArgumentNullException(nameof(aggregator));
        }

        public byte[] GenerateExcelWithChartsFirstTask(List<Vulnerability> vulnerabilities)
        {
            var yearlyData = _aggregator.ByYearAndType(vulnerabilities);
            var totalData = _aggregator.TotalByType(yearlyData);

            using var workbook = new XLWorkbook();

            foreach (var (year, typeCounts) in yearlyData)
            {
                var yearlySheet = workbook.Worksheets.Add($"Vulnerabilities_{year}");
                yearlySheet.Cell(1, 1).Value = "Type";
                yearlySheet.Cell(1, 2).Value = "Count";

                int rowYear = 2;
                foreach (var (type, count) in typeCounts)
                {
                    yearlySheet.Cell(rowYear, 1).Value = type;
                    yearlySheet.Cell(rowYear, 2).Value = count;
                    rowYear++;
                }

                var yearlyTableRange = yearlySheet.RangeUsed();
                var yearlyTable = yearlyTableRange.CreateTable();
                yearlyTable.Theme = XLTableTheme.TableStyleMedium9;
                yearlySheet.Columns().AdjustToContents();
            }

            var totalSheet = workbook.Worksheets.Add("VulnerabilitiesTotal");
            totalSheet.Cell(1, 1).Value = "Type";
            totalSheet.Cell(1, 2).Value = "Count";

            int rowTotal = 2;
            foreach (var (type, count) in totalData)
            {
                totalSheet.Cell(rowTotal, 1).Value = type;
                totalSheet.Cell(rowTotal, 2).Value = count;
                rowTotal++;
            }

            var totalTableRange = totalSheet.RangeUsed();
            var totalTable = totalTableRange.CreateTable();
            totalTable.Theme = XLTableTheme.TableStyleMedium9;
            totalSheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            return stream.ToArray();
        }

        public byte[] GenerateExcelWithChartsSecondTask(List<Vulnerability> vulnerabilities)
        {

            var filteredData = vulnerabilities.Where(v => v.Type == "Прикладное ПО информационных систем").ToList();

            var yearlyData = _aggregator.ByYearAndType(filteredData);

            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("VulnerabilitiesByYear");

            sheet.Cell(1, 1).Value = "Year";
            sheet.Cell(1, 2).Value = "Count";

            int row = 2;
            foreach (var (year, typeCounts) in yearlyData)
            {
                var countForYear = typeCounts.ContainsKey("Прикладное ПО информационных систем") ? typeCounts["Прикладное ПО информационных систем"] : 0;

                sheet.Cell(row, 1).Value = year;
                sheet.Cell(row, 2).Value = countForYear;
                row++;
            }

            var tableRange = sheet.RangeUsed();
            var table = tableRange.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium9;
            sheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            return stream.ToArray();
        }


        public byte[] GenerateExcelWithChartsThirdTask(List<Vulnerability> vulnerabilities)
        {
            var filteredData = vulnerabilities.Where(v => v.Type == "Прикладное ПО информационных систем").ToList();

            var vendorData = filteredData
                .SelectMany(v => v.Vendor.Split(',').Select(vendor => vendor.Trim())) 
                .GroupBy(vendor => vendor)
                .Select(g => new { Vendor = g.Key, Count = g.Count() })
                .OrderByDescending(v => v.Count)
                .ToList();

            var classData = filteredData
                .GroupBy(v => v.Class)
                .Select(g => new { Class = g.Key, Count = g.Count() })
                .OrderByDescending(v => v.Count)
                .ToList();

            var dangerLevelData = filteredData
                .GroupBy(v => v.DangerousLevel)
                .Select(g => new { Level = g.Key, Count = g.Count() })
                .OrderByDescending(v => v.Count)
                .ToList();

            var cweErrorData = filteredData
                .GroupBy(v => v.ErrorType)
                .Select(g => new { ErrorType = g.Key, Count = g.Count() })
                .OrderByDescending(v => v.Count)
                .Take(25) 
                .ToList();

            var topSoftwareData = filteredData
                .GroupBy(v => v.PoName)
                .Select(g => new { Software = g.Key, Count = g.Count() })
                .OrderByDescending(v => v.Count)
                .Take(5)
                .ToList();

            var totalCount = filteredData.Count();
            var unconfirmedCount = filteredData.Count(v => v.Status.Trim().Equals("Потенциальная уязвимость", StringComparison.OrdinalIgnoreCase)); 
            var unconfirmedPercentage = totalCount > 0 ? (double)unconfirmedCount / totalCount * 100 : 0;

            var exploitCount = filteredData.Count(v => v.Exploit.Trim().Equals("Существует", StringComparison.OrdinalIgnoreCase)); 
            var exploitPercentage = totalCount > 0 ? (double)exploitCount / totalCount * 100 : 0;

            using var workbook = new XLWorkbook();

            var vendorSheet = workbook.Worksheets.Add("Vendors");
            vendorSheet.Cell(1, 1).Value = "Vendor";
            vendorSheet.Cell(1, 2).Value = "Count";
            int rowVendor = 2;
            foreach (var vendor in vendorData)
            {
                vendorSheet.Cell(rowVendor, 1).Value = vendor.Vendor;
                vendorSheet.Cell(rowVendor, 2).Value = vendor.Count;
                rowVendor++;
            }
            var vendorTableRange = vendorSheet.RangeUsed();
            var vendorTable = vendorTableRange.CreateTable();
            vendorTable.Theme = XLTableTheme.TableStyleMedium9;
            vendorSheet.Columns().AdjustToContents();

            var classSheet = workbook.Worksheets.Add("Vulnerability Classes");
            classSheet.Cell(1, 1).Value = "Class";
            classSheet.Cell(1, 2).Value = "Count";
            int rowClass = 2;
            foreach (var classItem in classData)
            {
                classSheet.Cell(rowClass, 1).Value = classItem.Class;
                classSheet.Cell(rowClass, 2).Value = classItem.Count;
                rowClass++;
            }
            var classTableRange = classSheet.RangeUsed();
            var classTable = classTableRange.CreateTable();
            classTable.Theme = XLTableTheme.TableStyleMedium9;
            classSheet.Columns().AdjustToContents();

            var dangerSheet = workbook.Worksheets.Add("Danger Levels");
            dangerSheet.Cell(1, 1).Value = "Danger Level";
            dangerSheet.Cell(1, 2).Value = "Count";
            int rowDanger = 2;
            foreach (var danger in dangerLevelData)
            {
                dangerSheet.Cell(rowDanger, 1).Value = danger.Level;
                dangerSheet.Cell(rowDanger, 2).Value = danger.Count;
                rowDanger++;
            }
            var dangerTableRange = dangerSheet.RangeUsed();
            var dangerTable = dangerTableRange.CreateTable();
            dangerTable.Theme = XLTableTheme.TableStyleMedium9;
            dangerSheet.Columns().AdjustToContents();

            var cweSheet = workbook.Worksheets.Add("CWE Errors");
            cweSheet.Cell(1, 1).Value = "CWE Error Type";
            cweSheet.Cell(1, 2).Value = "Count";
            int rowCwe = 2;
            foreach (var error in cweErrorData)
            {
                cweSheet.Cell(rowCwe, 1).Value = error.ErrorType;
                cweSheet.Cell(rowCwe, 2).Value = error.Count;
                rowCwe++;
            }
            var cweTableRange = cweSheet.RangeUsed();
            var cweTable = cweTableRange.CreateTable();
            cweTable.Theme = XLTableTheme.TableStyleMedium9;
            cweSheet.Columns().AdjustToContents();

            var softwareSheet = workbook.Worksheets.Add("Top Software");
            softwareSheet.Cell(1, 1).Value = "Software";
            softwareSheet.Cell(1, 2).Value = "Count";
            int rowSoftware = 2;
            foreach (var software in topSoftwareData)
            {
                softwareSheet.Cell(rowSoftware, 1).Value = software.Software;
                softwareSheet.Cell(rowSoftware, 2).Value = software.Count;
                rowSoftware++;
            }
            var softwareTableRange = softwareSheet.RangeUsed();
            var softwareTable = softwareTableRange.CreateTable();
            softwareTable.Theme = XLTableTheme.TableStyleMedium9;
            softwareSheet.Columns().AdjustToContents();

            var percentageSheet = workbook.Worksheets.Add("Percentages");
            percentageSheet.Cell(1, 1).Value = "Description";
            percentageSheet.Cell(1, 2).Value = "Percentage (%)";
            percentageSheet.Cell(2, 1).Value = "Unconfirmed Vulnerabilities";
            percentageSheet.Cell(2, 2).Value = unconfirmedPercentage;
            percentageSheet.Cell(3, 1).Value = "Exploitable Vulnerabilities";
            percentageSheet.Cell(3, 2).Value = exploitPercentage;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);

            return stream.ToArray();
        }
    }
}
