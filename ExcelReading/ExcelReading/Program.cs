using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelReading
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "C:\\Users\\raziy\\Downloads\\malzemeAgirlikHesaplama.xlsx";

            Application Excel = new Application();
            Workbook wbook = Excel.Workbooks.Open(path);
            Worksheet excelSheet = wbook.ActiveSheet();

            Excel.Range range = excelSheet.UsedRange; //Excel sayfasındaki tüm satır ve sütun alanlarını alır.
            int satirSayisi = range.Rows.Count; //Excel sayfasının satır sayısını alır.
            int sutunSayisi = range.Columns.Count; // Excel sayfasının sütun sayısını alır.
            int excel_eleman_sayisi = satirSayisi * sutunSayisi;
            int i = 1;
            int j = 1;

            while(i <= satirSayisi)
            {
                while(j <= sutunSayisi)
                {
                    Console.Write(Convert.ToInt32(excelSheet.Cells[i, j].Value));
                    Console.Write("\t");
                    j++;
                }
                Console.WriteLine( "\n");
                j = 1;
                i++;
            }
            Console.ReadLine();
        }
    }
}
