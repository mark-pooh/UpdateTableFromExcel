using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using System;
using System.IO;

namespace UpdateTableFromExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filename = "all range.xlsx";   //args[0];

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filename)))
                {
                    var workSheet = package.Workbook.Worksheets[0];
                    var start = workSheet.Dimension.Start;
                    var end = workSheet.Dimension.End;

                    OracleConnection conn = new OracleConnection("Provider = OraOLEDB.Oracle; Data Source = NEPSDEVL; User Id = NEPSBI; Password = xs2nepsbi");
                    conn.Open();

                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        //Row by row
                        for (int col = start.Column; col <= end.Column; col++)
                        {
                            //Cell by cell
                            object cellValue = workSheet.Cells[row, col].Text;
                            Console.Write("column : {0}, value: {1}", col, cellValue);
                        }
                        Console.WriteLine();
                    }

                    conn.Dispose();
                    conn.Close();
                }
            }
            catch (IOException e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
        }
    }
}