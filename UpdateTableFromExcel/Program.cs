using OfficeOpenXml;
using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;

namespace UpdateTableFromExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filename = args[0];
            string connString = args[1];

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(filename)))
                {
                    int count = package.Workbook.Worksheets.Count;
                    Console.WriteLine("Number of sheets : " + count);
                    Console.Write("Please insert desired sheet (1 - n): ");
                    int selection = int.Parse(Console.ReadLine());

                    if(selection <= count)
                    {
                        var workSheet = package.Workbook.Worksheets[selection - 1];
                        var start = workSheet.Dimension.Start;
                        var end = workSheet.Dimension.End;

                        OleDbConnection conn = new OleDbConnection(connString);
                        conn.Open();

                        Console.WriteLine("Connection Opened\n");
                        Console.Write("Please insert table name : ");
                        string tableName = Console.ReadLine();

                        Console.WriteLine("\nYou've picked table {0}.\n", tableName);

                        OleDbCommand cmd_Segment = new OleDbCommand
                        {
                            Connection = conn,
                            CommandText = "SELECT COLUMN_NAME FROM USER_TAB_COLUMNS WHERE TABLE_NAME = UPPER('" + tableName + "')"
                        };
                        cmd_Segment.CommandType = CommandType.Text;
                        OleDbDataReader dr_Segment = cmd_Segment.ExecuteReader();

                        string colFilter = "", colValue = "";

                        if(dr_Segment != null)
                        {
                            int i = 0;

                            List<string> columns = new List<string>();

                            while(dr_Segment.Read())
                            {
                                columns.Add(dr_Segment.GetString(0));
                                Console.WriteLine("{0}) {1}", ++i, dr_Segment.GetString(0));
                            }

                            Console.Write("\nPlease insert column index to filter (1, 2, n) : ");
                            selection = int.Parse(Console.ReadLine());

                            Console.WriteLine("\nYou've picked column {0} as the filter.", columns[selection - 1]);
                            colFilter = columns[selection - 1];

                            Console.Write("\nPlease insert column index to change (1, 2, n) : ");
                            selection = int.Parse(Console.ReadLine());

                            Console.WriteLine("\nYou've picked column {0} to be changed.", columns[selection - 1]);
                            colValue = columns[selection - 1];
                        }

                        dr_Segment.Close();
                        cmd_Segment.Dispose();

                        int excelColFilter = 1, excelColValue = 2;

                        Console.WriteLine("\n------------EXCEL------------\n");
                        Console.WriteLine("Total number of columns = {0}", end.Column);

                        Console.Write("\nPlease insert column index to filter (1, 2, n) : ");
                        excelColFilter = int.Parse(Console.ReadLine());

                        Console.WriteLine("\nYou've picked column {0} as the filter.", excelColFilter);

                        Console.Write("\nPlease insert column index to replace (1, 2, n) : ");
                        excelColValue = int.Parse(Console.ReadLine());

                        Console.WriteLine("\nYou've picked column {0} as the replacement value.", excelColValue);

                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();

                        for (int row = start.Row; row <= end.Row; row++)
                        {
                            ////Row by row
                            //for (int col = start.Column; col <= end.Column; col++)
                            //{
                            //    //Cell by cell
                            //    object cellValue = workSheet.Cells[row, col].Text;
                            //    Console.Write("column : {0}, value: {1}", col, cellValue);
                            //}
                            //Console.WriteLine();

                            string filterVal = "", replaceVal = "";

                            filterVal = workSheet.Cells[row, excelColFilter].Text;
                            replaceVal = workSheet.Cells[row, excelColValue].Text;

                            Console.WriteLine("Processing row {0}", row);

                            OleDbCommand cmd_UpdateCoor = new OleDbCommand
                            {
                                Connection = conn,
                                CommandText = "UPDATE " + tableName + " SET " + colValue + " = :replace WHERE " + colFilter + " = :filter"
                            };
                            cmd_UpdateCoor.Parameters.AddWithValue(":replace", replaceVal);
                            cmd_UpdateCoor.Parameters.AddWithValue(":filter", filterVal);
                            cmd_UpdateCoor.ExecuteNonQuery();
                            cmd_UpdateCoor.Dispose();

                            Console.Clear();
                        }

                        conn.Dispose();
                        conn.Close();
                    }
                    else
                    {
                        Console.WriteLine("Invalid selection");
                    }
                }
            }
            catch (IOException e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }

            Console.ReadKey();
        }
    }
}