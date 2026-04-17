using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.IO;

namespace pamana_approval_addon
{
    class Logger 
    {
        public static void WriteToFile(string sStatus, string sSubject, string Message)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SAP ADDONS\\APPROVAL\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SAP ADDONS\\APPROVAL\\Logs\\APPROVAL_LOG_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(DateTime.Now + "\t" + sStatus + "\t" + sSubject + "\t" + Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(DateTime.Now + "\t" + sStatus + "\t" + sSubject + "\t" + Message);
                }
            }

            path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SAP ADDONS\\APPROVAL\\";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string configpath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SAP ADDONS\\APPROVAL\\config.ini";
            if (!File.Exists(configpath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(configpath))
                {
                    sw.WriteLine("PORT: -->");
                    sw.WriteLine("SERVER: -->");
                    sw.WriteLine("FOLDER: -->");
                    sw.WriteLine("MODE: --> R");

                }
            }
        }


    }
}
