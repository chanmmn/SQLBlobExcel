using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ConsoleAppExcelPivot
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GenerateSpreadsheetWithPivotTable();
            Console.WriteLine("Spreadsheet with pivot table generated successfully.");
        }

        static void GenerateSpreadsheetWithPivotTable()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("SampleData");
            // Sample data
            var data = new List<string[]>
                {
                    new string[] { "Category", "Product", "Amount" },
                    new string[] { "Fruit", "Apple", "50" },
                    new string[] { "Fruit", "Banana", "30" },
                    new string[] { "Vegetable", "Carrot", "20" },
                    new string[] { "Vegetable", "Broccoli", "40" }
                };
            // Fill sheet with sample data
            for (int i = 0; i < data.Count; i++)
            {
                IRow row = sheet.CreateRow(i);
                for (int j = 0; j < data[i].Length; j++)
                {
                    row.CreateCell(j).SetCellValue(data[i][j]);
                }
            }
            // Create pivot table
            ISheet pivotSheet = workbook.CreateSheet("PivotTable");
            IName name = workbook.CreateName();
            name.RefersToFormula = "SampleData!$A$1:$C$5";
            name.NameName = "DataRange";
            AreaReference source = new AreaReference("SampleData!$A$1:$C$5", SpreadsheetVersion.EXCEL2007);
            CellReference position = new CellReference("A1");
            XSSFPivotTable pivotTable = (XSSFPivotTable)((XSSFSheet)pivotSheet).CreatePivotTable(source, position); // Cast pivotSheet to XSSFSheet
            // Configure pivot table
            pivotTable.AddRowLabel(0); // Category
            pivotTable.AddColumnLabel(DataConsolidateFunction.SUM, 2); // Sum of Amount
            pivotTable.AddColumnLabel(DataConsolidateFunction.COUNT, 1); // Count of Product
            // Save the workbook to a file
            using (var fileData = new FileStream("SamplePivotTable.xlsx", FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }
    }
}
