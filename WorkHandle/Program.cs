using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace WorkHandle
{
    class Program
    {
        static void Main(string[] args)
        {
            //String workPath = System.Environment.CurrentDirectory;
            String simulatorPath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            //System.Console.WriteLine(simulatorPath);
            System.IO.Directory.SetCurrentDirectory(simulatorPath);
            String logStr = "(null)";
            String shutDownStr = "";
            if (args.Length != 0)
            {
                logStr = args.First();
                shutDownStr = args.Last();
            }
            //IniFile.instance.IniWriteValue("data", "1122211", "2");
            //BuildRelease();
            EditXls(logStr);
            String str = logStr.Replace(",", ",\r\n");
            //str = logStr.Replace(" ", "");
            //System.Console.WriteLine(str);
            String timeStr = DateTime.Now.ToString();
            WriteToDoc(timeStr + "\r\n" + str);
            CompileProject();
            CopySimulator();
            CommitSVN();
            CommitClassesAndResources(logStr);
            if (shutDownStr == @"-s")
            {
                ShutDownComputer();
            }
        }

        public static void WriteToDoc(String str)
        {
            FileStream fs = new FileStream(@"doc.txt", FileMode.Append);
            //获得字节数组
            byte[] data = System.Text.Encoding.Default.GetBytes("\r\n" + str+ "\r\n");
            //开始写入
            fs.Write(data, 0, data.Length);
            //清空缓冲区、关闭流
            fs.Flush();
            fs.Close();
        }

        public static void RunBat(String str)
        {
            Process pro = new Process();
            String batPath = Path.Combine(System.Environment.CurrentDirectory, str);
            FileInfo file = new FileInfo(batPath);
            pro.StartInfo.WorkingDirectory = file.Directory.FullName;
            pro.StartInfo.FileName = batPath;
            pro.StartInfo.CreateNoWindow = false;
            pro.Start();
            pro.WaitForExit();
        }

        public static void RunCMD(String str)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;    //是否使用操作系统shell启动
            p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            p.Start();//启动程序
            //向cmd窗口发送输入信息
            p.StandardInput.WriteLine(str);
            p.StandardInput.AutoFlush = true;
            p.StandardInput.WriteLine("exit");
            p.WaitForExit();//等待程序执行完退出进程
            p.Close();
        }

        public static void CommitClassesAndResources(String logStr)
        {
            String projectName = IniFile.instance.ProjectName();
            String classesPath = "..\\..\\..\\games\\"+ projectName + "\\src\\client\\frameworks\\runtime-src\\Classes";
            String resourcesPath = "..\\..\\..\\games\\" + projectName + "\\src\\client\\frameworks\\runtime-src\\res";
            /*RunCMD("svn update ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Classes");
            RunCMD("svn update  ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Resources");
            RunCMD("svn add --force ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Classes");
            RunCMD("svn add --force ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Resources");
            RunCMD("svn ci -m \"" + logStr + "\" ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Classes");
            RunCMD("svn ci -m \"资源提交\" ..\\..\\..\\games\\bubbleDragonSecond\\src\\client\\Resources");
            RunCMD("svn update  E:\\svn\\shared_doc\\工作日志\\journal_黄泽昊.xlsx");
            RunCMD("svn ci -m \"日志提交\" E:\\svn\\shared_doc\\工作日志\\journal_黄泽昊.xlsx"); */
            RunCMD("svn update " + classesPath);
            RunCMD("svn update " + resourcesPath);
            RunCMD("svn add --force " + classesPath);
            RunCMD("svn add --force " + resourcesPath);
            RunCMD("svn ci -m \"" + logStr + "\" " + classesPath);
            RunCMD("svn ci -m \"资源提交\" " + resourcesPath);
            String journalPath = IniFile.instance.JournalPath();
            RunCMD("svn update " + journalPath);
            RunCMD("svn ci -m \"日志提交\" " + journalPath);
        }

        //public static void BuildRelease()
        //{
        //    //string fileName = @"C:/Program Files (x86)/Microsoft Visual Studio 14.0/Common7/IDE/devenv.exe";
        //    Process p = new Process();
        //    p.StartInfo.UseShellExecute = false;
        //    p.StartInfo.RedirectStandardOutput = true;
        //    p.StartInfo.FileName = IniFile.instance.DevenvPath();
        //    p.StartInfo.CreateNoWindow = true;

        //    p.StartInfo.Arguments = "..\\..\\..\\games\\"+ IniFile.instance.ProjectName() + "\\src\\client\\frameworks\\runtime-src\\proj.win32\\" + 
        //        IniFile.instance.ProjectName() + ".vcxproj /Build \"Release|Win32\"";//参数以空格分隔，如果某个参数为空，可以传入””
        //    p.Start();
        //    p.WaitForExit();
        //}

        static void EditXls(String str)
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
            Excel.Workbook xlsWorkBook = excelApp.Workbooks.Open(IniFile.instance.JournalPath());
            Excel.Worksheet xlsWorkSheet = (Worksheet)xlsWorkBook.Worksheets["Sheet1"];
            int rowsCount = xlsWorkSheet.UsedRange.Rows.Count;
            int colsCount = xlsWorkSheet.UsedRange.Columns.Count;
            //rowsCount：最大行    colsCount：最大列
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[1, 1];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.Cells[rowsCount, colsCount];
            Range rng = (Microsoft.Office.Interop.Excel.Range)xlsWorkSheet.get_Range(c1, c2);
            xlsWorkSheet.Cells[rowsCount+1, 1] = "泡泡西游";
            xlsWorkSheet.Cells[rowsCount + 1, 2] = str;
            xlsWorkSheet.Cells[rowsCount + 1, 3] = DateTime.Now.ToString();
            xlsWorkSheet.Cells[rowsCount + 1, 4] = "脚本处理";
            object[,] exceldata = (object[,])rng.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);
            xlsWorkBook.Save();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkBook);
            excelApp.Quit();
            GC.Collect();
        }

        static void ShutDownComputer()
        {
            RunCMD("shutdown -s -t 10");
        }

        static void CopySimulator()
        {
            RunCMD("rd .\\simulator /s /q");
            RunCMD("xcopy ..\\..\\..\\games\\" + IniFile.instance.ProjectName() + "\\src\\client\\publish\\*  .\\simulator /s /i");
            //RunCMD("xcopy ..\\..\\..\\games\\" + IniFile.instance.ProjectName() + "\\src\\client\\simulator  .\\ /i");
            //RunCMD("rd .\\Resources /s /q");
            //RunCMD("xcopy ..\\..\\..\\games\\" + IniFile.instance.ProjectName() + "\\src\\client\\Resources /s .\\Resources /y /i");
            CommitSVN();
        }

        static void CommitSVN()
        {
            RunCMD("svn add --force *");
            RunCMD("svn ci -m \"bat提交\" *");
        }

        static void CompileProject()
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;    //是否使用操作系统shell启动
            p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            p.Start();//启动程序
            //向cmd窗口发送输入信息
            //String simulatorPath = "..\\..\\..\\games\\" + IniFile.instance.ProjectName() + "\\src\\client\\simulator";
            p.StandardInput.WriteLine("cd ..\\..\\..\\games\\" + IniFile.instance.ProjectName() + "\\src\\client\\frameworks\\runtime-src");
            p.StandardInput.WriteLine("cocos compile -m release -p android");
            p.StandardInput.WriteLine("cocos compile -m release -p win32");
            //p.StandardInput.WriteLine("xcopy ..\\bin\\release\\android\\" + IniFile.instance.ProjectName() + "-release-signed.apk /k ..\\..\\..\\..\\..\\模拟器\\在研\\" + IniFile.instance.SimulatorName() + " /y /i");
            //p.StandardInput.WriteLine("cd ..\\..\\..\\..\\..\\模拟器\\在研\\" + IniFile.instance.SimulatorName() + "\\");
            p.StandardInput.AutoFlush = true;
            p.StandardInput.WriteLine("exit");
            p.WaitForExit();//等待程序执行完退出进程
            p.Close();
        }
    }
}
