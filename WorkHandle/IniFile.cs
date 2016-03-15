using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
namespace WorkHandle
{
    class IniFile
    {
        public readonly String SECTION_STRING = "config";
        public readonly String KEY_DEVENV_PATH = "devenvPath";
        public readonly String KEY_PROJECT_NAME = "projectName";
        public readonly String KEY_JOURNAL_PATH = "journalPath";
        public readonly String KEY_SIMULATOR_NAME = "simulatorName";
        private string path;             //INI文件名  
        //声明读写INI文件的API函数  
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key,
                    string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def,
                    StringBuilder retVal, int size, string filePath);

        //类的构造函数，传递INI文件名 
        public static readonly IniFile instance = new IniFile();
        private IniFile()
        {
            path = System.IO.Path.Combine(System.Environment.CurrentDirectory, "Config.ini");
            if (!System.IO.File.Exists(this.path))
            {
                String sectionStr = SECTION_STRING;
                this.IniWriteValue(sectionStr, KEY_DEVENV_PATH, @"C:/Program Files (x86)/Microsoft Visual Studio 14.0/Common7/IDE/devenv.exe");
                this.IniWriteValue(sectionStr, KEY_PROJECT_NAME, "bubbleDragonSecondLua");
                this.IniWriteValue(sectionStr, KEY_JOURNAL_PATH, @"E:\svn\shared_doc\工作日志\journal_黄泽昊.xlsx");
                this.IniWriteValue(sectionStr, KEY_SIMULATOR_NAME, "泡泡西游");
            }
        }

        //写INI文件  
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }

        //读取INI文件指定  
        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
            return temp.ToString();
        }

        //获取client路径
        public String DevenvPath()
        {
            return this.IniReadValue(SECTION_STRING, KEY_DEVENV_PATH);
        }

        public String ProjectName()
        {
            return this.IniReadValue(SECTION_STRING, KEY_PROJECT_NAME);
        }

        public String JournalPath()
        {
            return this.IniReadValue(SECTION_STRING, KEY_JOURNAL_PATH);
        }

        public String SimulatorName()
        {
            return this.IniReadValue(SECTION_STRING, KEY_SIMULATOR_NAME);
        }
    }
}
