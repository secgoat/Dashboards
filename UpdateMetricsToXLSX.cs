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

        public void Convert(String metricsFolder)
        {

            metrics_files = Directory.GetFiles(metricsFolder);
            foreach (String file in metrics_files)
            {
               /* if (File.Exists(file))
                {
                    File.Copy(file, "..\\backups\\" +file);
                } */
                
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Open(file);
                wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                wb.Close();
                app.Quit();
            }

        }
    }
}
