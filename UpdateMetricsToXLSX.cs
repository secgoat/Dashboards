using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProviderDashboards
{
    class UpdateMetricsToXLSX
    {
        private string[] metrics_files;
        private string name;
       
        public void Convert(String metricsFolder)
        {
            metrics_files = Directory.GetFiles(metricsFolder);
            foreach (String file in metrics_files)
            {
                string[] nameArray = file.Split('\\'); //get an array of all elements split at \\
                name = nameArray[nameArray.Length - 1]; //return the last element of the array which shoudl always be the file name
                string path = metricsFolder + "\\backups\\";

                //check to make sure backup directory exists, otherwise create it
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Open(file);
                wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                wb.Close();
                app.Quit();
                path += name;
                File.Move(file, path);
            }

        }
    }
}
