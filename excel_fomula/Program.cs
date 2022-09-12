using System;
using System.Collections.Generic;
using System.Linq;
using Excel=Microsoft.Office.Interop.Excel;

namespace excel_fomula
{
    class Program
    {
        internal class Part
        {
            internal int rowid { get; set; }
            internal int colid { get; set; }
            internal string fomula_value { get; set; }
        }
        static void Main(string[] args)
        {
            List<Part> parts = new List<Part>();

            read_fomula(parts);  //在處理前先掃整個檔案的公式欄位
            //產生工單，結束要關掉 excel Workbook 避免檔案被鎖住
            write_fomula(parts); //產生工單之後寫回公式欄位
        }
        static void read_fomula(List<Part> target_fomula)
        {

            Excel.Range dataRange = null;
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string strPath = path + @"shopdoc-條碼版.xlsx";
            string target_item;
            int totalColumns, totalRows;


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook myWorkBook = excelApp.Workbooks.Open(strPath);
            Excel.Worksheet mySheet = myWorkBook.ActiveSheet;

            totalColumns = mySheet.UsedRange.Columns.Count;
            totalRows = mySheet.UsedRange.Rows.Count;

            for (int row = 1; row < totalRows; row++) 
            {
                for (int col = 1; col < totalColumns; col++) 
                {
                    dataRange = (Excel.Range)mySheet.Cells[row, col];
                    if (dataRange.Value2 != null)
                    {
                        target_item = dataRange.FormulaArray.ToString();
   
                        if (target_item.FirstOrDefault() == '=')
                        {
                            //System.Console.WriteLine(target_item);
                            target_fomula.Add(new Part() { rowid = row, colid = col, fomula_value = target_item });
                        }
                           
                    }
                    
                }
                
            }
            myWorkBook.Close();
            //System.Console.ReadKey();  
            //Console.WriteLine();
        }

        static void write_fomula(List<Part> target_fomula)
        {

            //int i = 0, j = 0;
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string strPath = path + @"shopdoc-條碼版.xlsx";

            Excel.Application App = new Excel.Application();
            Excel.Workbook Wbook = App.Workbooks.Open(strPath);
            Excel.Worksheet Wsheet = Wbook.ActiveSheet;


            //System.Console.ReadKey(); 
            foreach (Part aPart in target_fomula)
            {
                //Console.WriteLine(aPart.rowid.ToString() + "," + aPart.colid.ToString() + "," + aPart.fomula_value);
                Wsheet.Cells[aPart.rowid, aPart.colid] = aPart.fomula_value;
            }
            //System.Console.ReadKey();
            //Console.WriteLine();
            //Wsheet.Cells[1, 1] = "final_test";

            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;

            Wbook.Save();
            Wbook.Close();
            App.Quit();   

        }


    }
}
