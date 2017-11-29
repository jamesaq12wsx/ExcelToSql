using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToSql
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("輸入檔案路徑");

                string strFilePath = Console.ReadLine();

                if (strFilePath == "" || strFilePath == null)
                {
                    break;
                }
                else
                {
                    TranExcelToSql(strFilePath);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fp">檔案路徑</param>
        static public void TranExcelToSql(string fp)
        {
            //Console.WriteLine("請輸入檔案");

            string filePath = fp;

            //if (filePath == "")
            //{
            //    break;
            //}

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            Console.WriteLine(xlWorksheet.Name.Split('-').First());
            Console.WriteLine(xlRange.Rows.Count);
            Console.WriteLine(xlRange.Columns.Count);
            //var dbName = xlWorksheet.Name.Split('-').First();
            var fileName = xlWorkbook.Name;
            var dbName = xlWorkbook.Name.Split('-').First() + "_en8846";
            int rows = xlRange.Rows.Count;
            //測試 row 可不可以省略空的行
            rows = xlRange.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            int columns = xlRange.Columns.Count;

            for (int r = 0; r < Math.Ceiling(rows / 1000.00); r++)
            {
                string strSqlDeclare = string.Format("--{0}\ndeclare @INITDATE datetime\nset @INITDATE = '2017/11/30'\ndeclare @INIT varchar(10)\nset @INIT = 'INIT'\n", fileName);

                string strSqlTryBegin = string.Format("begin try\n");

                string insertSql = String.Format("INSERT INTO {0}\nVALUES ", dbName);

                string strCountSql = "";

                insertSql = strSqlDeclare + strSqlTryBegin + insertSql;
                string values = "";
                for (int i = r * 1000 + 1; i <= rows; i++)
                {
                    if (i - (r * 1000 + 1) >= 1000)
                    {
                        break;
                    }
                    //cmn_rating_level 地一行有欄位名稱
                    if ((dbName == "cmn_rating_level_en8846" && i == 1) ||
                        (dbName == "ims_invt_unit_en8846" && i == 1) ||
                        (dbName == "skl_portfolio_en8846" && i == 1) ||
                        (dbName == "cmn_equity_en8846" && i == 1) ||
                        (dbName == "ims_bank_acc_en8846" && i == 1) ||
                        (dbName == "cmn_region_en8846" && i==1))
                    {
                        continue;
                    }
                    string rowValue = "";
                    for (int j = 1; j <= columns; j++)
                    {
                        //write the value to the console
                        var value = "";

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            Console.Write(xlRange.Cells[i, j].Value.ToString().Trim() + "\t");

                            string cell = xlRange.Cells[i, j].Value.ToString().Trim();
                            //判斷是否為日期
                            //if (DateTime.TryParse(xlRange.Cells[i, j].Value.ToString(), out DateTime dt))
                            //{
                            //    //value = "\'" + dt.Date.ToString("yyyy/MM/dd") + "\'";

                            //    //帶入變數
                            //    value = "@INITDATE";
                            //}
                            //else if (xlRange.Cells[i, j].Value.ToString() == "admin" || xlRange.Cells[i,j].Value.ToString() == "ETL" || xlRange.Cells[i,j].Value.ToString() == "INIT")
                            //{
                            //    value = "@INIT";
                            //}
                            //else if (xlRange.Cells[i,j].Value.ToString().Contains("'"))
                            //{
                            //    value = "\'" + ExcapeSingleQuote(xlRange.Cells[i, j].Value.ToString().Trim()) + "\'";
                            //}
                            //else
                            //{
                            //    value = "\'" + xlRange.Cells[i, j].Value.ToString().Trim() + "\'";
                            //}
                            if (DateTime.TryParse(cell, out DateTime dt))
                            {
                                //value = "\'" + dt.Date.ToString("yyyy/MM/dd") + "\'";

                                //帶入變數
                                value = "@INITDATE";
                            }
                            else if (cell == "admin" || cell == "ETL" || cell == "INIT")
                            {
                                value = "@INIT";
                            }
                            else if (cell.Contains("'"))
                            {
                                value = "\'" + ExcapeSingleQuote(cell) + "\'";
                            }
                            else
                            {
                                value = "\'" + cell + "\'";
                            }
                        }
                        else
                        {
                            value = string.Format("\'\'");
                        }

                        if (j == 1)
                        {
                            rowValue += value;
                        }
                        else
                        {
                            rowValue += "," + value;
                        }
                    }
                    if (i == 1)
                    {
                        values += string.Format("({0})", rowValue);
                    }
                    else
                    {
                        values += "," + string.Format("({0})", rowValue);
                    }

                    //輸出換行
                    Console.WriteLine();

                    strCountSql += string.Format("select count(*)\nfrom {0}\n", dbName);
                }
                if ((dbName == "cmn_rating_level_en8846" && r == 0) ||
                    (dbName == "ims_invt_unit_en8846" && r == 0) ||
                    (dbName == "skl_portfolio_en8846" && r == 0) ||
                    (dbName == "cmn_equity_en8846" && r == 0) ||
                    (dbName == "ims_bank_acc_en8846" && r == 0))
                {
                    values = values.Substring(1);
                }

                insertSql += string.Format("{0};\n", values);

                var strSqlEndTry = string.Format("print '{0} - {1} insert success'\nend try\n", dbName, r);

                var strSqlCatch = string.Format("begin catch\nprint '{0} - {1} insert failed'\nprint error_message()\nend catch\n", dbName, r);

                insertSql = insertSql + strSqlEndTry + strSqlCatch;

                //寫在bin裡面的
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(Environment.CurrentDirectory + @"\\AppData\INIT_test.sql", true))
                {
                    file.WriteLine(insertSql);
                }
            }

            //string strSqlTryBegin = string.Format("begin try\n");

            //string insertSql = String.Format("INSERT INTO {0}\nVALUES ", dbName);

            //string strCountSql = "";

            //insertSql = strSqlTryBegin + insertSql;
            //string values = "";
            //for (int i = 1; i <= rows; i++)
            //{
            //    //cmn_rating_level 地一行有欄位名稱
            //    if ((dbName == "cmn_rating_level_en8846" && i == 1) ||
            //        (dbName == "ims_invt_unit_en8846" && i == 1) ||
            //        (dbName == "skl_portfolio_en8846" && i == 1) ||
            //        (dbName == "cmn_equity_en8846" && i == 1) ||
            //        (dbName == "ims_bank_acc_en8846" && i == 1))
            //    {
            //        continue;
            //    }
            //    string rowValue = "";
            //    for (int j = 1; j <= columns; j++)
            //    {
            //        //write the value to the console
            //        var value = "";

            //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //        {
            //            Console.Write(xlRange.Cells[i, j].Value.ToString() + "\t");

            //            //判斷是否為日期
            //            if (DateTime.TryParse(xlRange.Cells[i, j].Value.ToString(), out DateTime dt))
            //            {
            //                value = "\'" + dt.Date.ToString("yyyy/MM/dd") + "\'";
            //            }
            //            else
            //            {
            //                value = "\'" + xlRange.Cells[i, j].Value.ToString() + "\'";
            //            }
            //        }
            //        else
            //        {
            //            value = string.Format("\'\'");
            //        }

            //        if (j == 1)
            //        {
            //            rowValue += value;
            //        }
            //        else
            //        {
            //            rowValue += "," + value;
            //        }
            //    }
            //    if (i == 1)
            //    {
            //        values += string.Format("({0})", rowValue);
            //    }
            //    else
            //    {
            //        values += "," + string.Format("({0})", rowValue);
            //    }

            //    //輸出換行
            //    Console.WriteLine();

            //    strCountSql += string.Format("select count(*)\nfrom {0}\n", dbName);
            //}
            //if ((dbName == "cmn_rating_level_en8846") ||
            //    (dbName == "ims_invt_unit_en8846") ||
            //    (dbName == "skl_portfolio_en8846") ||
            //    (dbName == "cmn_equity_en8846") ||
            //    (dbName == "ims_bank_acc_en8846"))
            //{
            //    values = values.Substring(1);
            //}

            //insertSql += string.Format("{0};\n", values);

            //var strSqlEndTry = string.Format("print '{0} insert success'\nend try\n", dbName);

            //var strSqlCatch = string.Format("begin catch\nprint '{0} insert failed'\nend catch\n", dbName);

            //insertSql = insertSql + strSqlEndTry + strSqlCatch;

            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\Document\國外帳務系統INIT\excel\INIT.sql", true))
            //{
            //    file.WriteLine(insertSql);
            //}

            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\Document\國外帳務系統INIT\excel\countInit.sql", true))
            //{
            //    file.WriteLine(strCountSql);
            //}

            xlWorkbook.Close(false);
        }

        public static string ExcapeSingleQuote(string str)
        {
            return str.Replace("'", "''");
        }
    }
}
