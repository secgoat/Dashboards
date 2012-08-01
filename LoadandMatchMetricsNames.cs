using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace ProviderDashboards
{
     class MetricName
        {
         string fileName = "";//use this to keep track of the file name
         int columnId; // this is the column it comes from or goes to depending on the report
         string metricName = ""; //use this to display the metric name from the file if needed

         public MetricName(String fileName, String metricName, int columnID)
         {
             this.fileName = fileName;
             this.metricName = metricName;
             this.columnId = columnID;
         }

        }

    class LoadandMatchMetricsNames
    {
        private Excel.Application app = null;
        private Excel.Workbook dashboard = null;
        private Excel.Workbook metrics = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range workSheet_range = null;
        /*
         * use these to store the matches between metrics name / location on the dashboard
         * to the metrics name / location on each metrics file
         * string = metrics name
         * Pair,int,int> = keeps the locaiton first of the dshboard metrics then the metric file location
         */
        private Dictionary<MetricName, MetricName> diabetesMetrics = new Dictionary<MetricName, MetricName>();
        private Dictionary<MetricName, MetricName> depressionMetrics = new Dictionary<MetricName, MetricName>();
        private Dictionary<MetricName, MetricName> asthmaMetrics = new Dictionary<MetricName, MetricName>();
        private Dictionary<MetricName, MetricName> cardiovascularMetrics = new Dictionary<MetricName, MetricName>();
        private Dictionary<MetricName, MetricName> preventiveMetrics = new Dictionary<MetricName, MetricName>();
        
        public LoadandMatchMetricsNames(String dashboardFile)
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                dashboard = app.Workbooks.Open(dashboardFile,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);


              /*  for (int i = 0; i < metricsFiles.Length; i++)
                {
                    metrics = app.Workbooks.Open(metricsFiles[i],
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                    //this is where we read all the data into some sort of array
                    //ExcelScanInternal(metrics);
                    metrics.Close(false, metricsFiles[i], null);
                } */


                //cleanup
                //dashboard.Close(false, dashboardFile, null);

            }
            catch (Exception ex) { }
        }
    }
}
