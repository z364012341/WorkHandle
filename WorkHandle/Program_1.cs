using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace workhandleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //Build();
            EditXls();
            Console.WriteLine("over");
            Console.ReadLine();
        }

        static void Build()
        {
            string fileName = @"C:/Program Files (x86)/Microsoft Visual Studio 14.0/Common7/IDE/devenv.exe";
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = fileName;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.Arguments = "D:\\Documents\\Cocos\\CocosProjects\\CocosProject\\proj.win32\\CocosProject.vcxproj /Build \"Release|Win32\"";//参数以空格分隔，如果某个参数为空，可以传入””
            p.Start();
            p.WaitForExit();
        }

        static void EditXls()
        {
            //设置程序运行语言
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //创建Application
            Excel.Application excelApp = new Excel.Application();
            //设置是否显示警告窗体
            excelApp.DisplayAlerts = false;
            //设置是否显示Excel
            excelApp.Visible = false;
            //禁止刷新屏幕
            excelApp.ScreenUpdating = false;
            //根据路径path打开
            Excel.Workbook xlsWorkBook = excelApp.Workbooks.Open(@"D:\Documents\Visual Studio 2015\Projects\workhandleTest\workhandleTest\w1.xlsx");
            Excel.Worksheet xlsWorkSheet = (Worksheet)xlsWorkBook.Worksheets["Sheet1"];
            int rowsCount = xlsWorkSheet.UsedRange.Rows.Count;
            int colsCount = xlsWorkSheet.UsedRange.Columns.Count;
            //rowsCount：最大行    colsCount：最大列
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[1, 1];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[rowsCount, colsCount];
            Range rng = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.get_Range(c1, c2);
            xlsWorkSheet.Cells[3, 1] = "!!!";
            object[,] exceldata = (object[,])rng.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);
            xlsWorkBook.Save();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkBook);
            excelApp.Quit();
            GC.Collect();
        }

    }
}
