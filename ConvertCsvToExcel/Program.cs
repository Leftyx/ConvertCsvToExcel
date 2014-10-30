using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertCsvToExcel
{
    using System.Diagnostics;
    using System.IO;

    using OfficeOpenXml;

    using ClosedXML.Excel;

    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    using ConvertCsvToExcel.ExtensionMethods;

    class Program
    {
        static void Main(string[] args)
        {
            string csvDocument = @"FL_insurance_sample.csv";
            string conversionFolder =  Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "Results");

            if (!Directory.Exists(conversionFolder))
            {
                Directory.CreateDirectory(conversionFolder);
            }
            else
            {
                Directory.GetFiles(conversionFolder).ForEach(f => File.Delete(f));
            }



            Stopwatch stopwatch = new Stopwatch();

            Console.WriteLine("Converting with EPPLus ...");
            stopwatch.Start();
            ConvertWithEPPlus(csvDocument, Path.Combine(conversionFolder, @"EPPlus.xlsx"), "EPPlus", ',');
            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            Console.WriteLine("Converting with ClosedXml ...");
            stopwatch = new Stopwatch();
            stopwatch.Start();
            var lines = ReadCsv(csvDocument, delimiter: ',');
            ConvertWithClosedXml(Path.Combine(conversionFolder, @"ClosedXml.xlsx"), "ClosedXml", lines);
            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            Console.WriteLine("Converting with NPOI ...");
            stopwatch = new Stopwatch();
            stopwatch.Start();
            lines = ReadCsv(csvDocument, delimiter: ',');
            ConvertWithNPOI(Path.Combine(conversionFolder, @"NPOI.xlsx"), "NPOI", lines);
            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            Console.WriteLine("Finished!");
            Console.ReadLine();
        }

        private static bool ConvertWithNPOI(string excelFileName, string worksheetName, IEnumerable<string[]> csvLines)
        {
            if (csvLines == null || csvLines.Count() == 0)
            {
                return (false);
            }

            int rowCount = 0;
            int colCount = 0;

            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet(worksheetName);

            foreach (var line in csvLines)
            {
                IRow row = worksheet.CreateRow(rowCount);

                colCount = 0;
                foreach (var col in line)
                {
                    row.CreateCell(colCount).SetCellValue(TypeConverter.TryConvert(col));
                    colCount++;
                }
                rowCount++;
            }

            using (FileStream fileWriter = File.Create(excelFileName))
            {
                workbook.Write(fileWriter);
                fileWriter.Close();
            }

            worksheet = null;
            workbook = null;

            return (true);
        }

        private static bool ConvertWithClosedXml(string excelFileName, string worksheetName, IEnumerable<string[]> csvLines)
        {
            if (csvLines == null || csvLines.Count() == 0)
            {
                return (false);
            }

            int rowCount = 0;
            int colCount = 0;

            using (var workbook = new XLWorkbook())
            {
                using (var worksheet = workbook.Worksheets.Add(worksheetName))
                {
                    rowCount = 1;
                    foreach (var line in csvLines)
                    {
                        colCount = 1;
                        foreach (var col in line)
                        {
                            worksheet.Cell(rowCount, colCount).Value = TypeConverter.TryConvert(col);
                            colCount++;
                        }
                        rowCount++;
                    }
                    
                }
                workbook.SaveAs(excelFileName);
            }

            return (true);
        }

        private static bool ConvertWithEPPlus(string csvFileName, string excelFileName, string worksheetName, char delimiter = ';')
        {
            bool firstRowIsHeader = false;

            var format = new ExcelTextFormat();
            format.Delimiter = delimiter;
            format.EOL = "\r";              // DEFAULT IS "\r\n";
            // format.TextQualifier = '"';

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                package.Save();
            }

            return (true);
        }

        private static IEnumerable<string[]> ReadCsv(string fileName, char delimiter = ';')
        {
            var lines = System.IO.File.ReadAllLines(fileName, Encoding.UTF8).Select(a => a.Split(delimiter));
            return (lines);
        }
    }
}
