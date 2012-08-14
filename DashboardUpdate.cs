using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

using ProviderDashboards.metrics;


namespace ProviderDashboards
{
    class DashboardUpdate
    {
        XLWorkbook _dashboard;
        XLWorkbook _metrics;
        IXLWorksheet _worksheet;
        IXLRange _worksheet_range;

        private string[] metrics_files;
        //use these to store the data from all the sheets
       
        List<object[,]> metricsList = new List<object[,]>(); //kepp track of all metrics data
        Dictionary<String, List<int>> dashboardRowLocations = new Dictionary<string, List<int>>();  //string = metrics name list = row numbers of new rows, should be bale to match them off based on provider list
        
        private List<String> providers;
        
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="providerNames"></param>
        public DashboardUpdate(List<String> providers)
        {
            this.providers = providers;
            stripnames();

        }
        private void stripnames()
        {
            ///<summary> use this to cut down the credentials form the provider names and just use first name
            /// this allows for  us to see if a cell contains that providers names so we can extrapolate the needed data from
            /// there
            /// </summary>
            providers.Sort();
            for (int i = 0; i < providers.Count; i++ )
            {
                int index = providers[i].IndexOf(' ');
                //index = providers[i].IndexOf(' ', index + 1); this line can trim it down to two names but has osme issues, so first name works for now
                providers[i] = providers[i].Substring(0, index);
            }
            providers.Insert(0, "Agency"); //add the Agency to identify the agency metrics locations
        }

        /// <summary>
        /// open the excel files and parse them into arrays and lists for easier 
        /// data manipulation
        /// </summary>
        /// <param name="dashboardFile"></param>
        /// <param name="metricsFolder"></param>
        public void _openExcel(String dashboardFile, String metricsFolder)
        {
            //grab all the files in the metrics folder
            metrics_files = Directory.GetFiles(metricsFolder);
            if (File.Exists("Copy.xlsx"))
            {
                File.Delete("Copy.xlsx");
            }
            File.Copy(dashboardFile, "Copy.xlsx");

            //try openeing the excel sheets listed in metricsFile
            try
            {
                _dashboard = new XLWorkbook(dashboardFile);

                for (int i = 0; i < metrics_files.Length; i++)
                {
                    _metrics = new XLWorkbook(metrics_files[i]);
                    //this is where we read all the data into some sort of array
                    MetricsToDictArray(_metrics);
                    _metrics.Dispose();
                }
                ExcelScanInternal(_dashboard);
                //MetricsToDashboard();
                _dashboard.Save();
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void MetricsToDashboard(Dictionary<String, IXLRangeRow> providerRows, IXLWorksheet sheet)
        {
            /* open the dsahboard here and create a dicitoary for each metric, where key = provider name and value =
             * row number of new row for provider. this will be the X value for the calculatiosn to add data to the sheets
             */

            XLWorkbook preventiveFile;
            List<XLWorkbook> PreventiveFiles = new List<XLWorkbook>();
            int[] preventivFileLocations = new int[] { 0, 3, 3, 15, 16, 17, 18 };//use 3 twice so we can get the 2 metrics from BMI file
            //this allows the loop in the preventive metrics object to iterate over file 3 2 times with different results

            foreach (int fileNum in preventivFileLocations)
            {
                preventiveFile = new XLWorkbook(metrics_files[fileNum]);
                PreventiveFiles.Add(preventiveFile);
            }
            // the following removed for brevity
            //bring it back after i get the data for 1 provider ot work correctly
            /*PreventiveMetric preventive = new PreventiveMetric(providers[1], PreventiveFiles);
            List<object> metrics = preventive.Metrics;
            var row = providerRows.ElementAt(1).Value;
            for (int x = 1; x < preventive.Metrics.Count; x++)
            {
                var cell = row.Cell(x);
                //cell.Style.NumberFormat.NumberFormatId = 0;
                cell.SetValue(preventive.Metrics.ElementAt(x));
            } */
            for(int i = 0; i < providers.Count; i++)
            {
                String fileName = metrics_files[3].ToString();
                PreventiveMetric preventive = new PreventiveMetric(providers[i], PreventiveFiles);
                List<object> metrics = preventive.Metrics;
                var row = providerRows.ElementAt(i).Value;
                for (int x = 1; x < metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(preventive.Metrics.ElementAt(x));
                } 
            }
       
            /* this was for asthma having some issues with the asthma one pulling data other than null from worksheets
            List<object[,]> asthmaMetrics = new List<object[,]>() { metricsList[1], metricsList[2] };
            
            XLWorkbook asthmaReport = new XLWorkbook(metrics_files[1]);
            foreach (string provider in providers)
            {
                AsthmaMetric asthma = new AsthmaMetric(provider, asthmaReport);
                foreach (var pair in providerRows)
                {
                    if (pair.Key.Contains(provider))
                    {
                        var row = pair.Value;
                        for (int i = 1; i < asthma.Metrics.Count; i++)
                        {
                            var cell = row.Cell(i);
                            cell.SetValue(asthma.Metrics.ElementAt(i) );
                        }
                    }
                }
            }*/
        }

        private void MetricsToDictArray(XLWorkbook metricsFile)
        {
            
            var sheet = metricsFile.Worksheet(1);
           
            int usedRows = sheet.RowsUsed().Count();

            List<String> data = new List<string>();

            var firstRowUsed = sheet.FirstRowUsed();
            var currentRow = firstRowUsed.RowUsed();
            string name = _metrics.Properties.Title;

            int lastColumn = sheet.LastColumnUsed().ColumnNumber();
            int lastRow = sheet.LastRowUsed().RowNumber(); //BAM! find the last row used
           
            var lastCell = currentRow.LastCellUsed(); //last cell gets set to nnull sometimes?

            object[,] sheetValue = new object[lastRow, lastColumn];
            
            while(lastCell == null)
            {
                currentRow = currentRow.RowBelow();
                lastCell = currentRow.LastCellUsed();
            }
            for (int i = 0; i < usedRows; i++)
            {
                lastCell = currentRow.LastCellUsed();
                if (lastCell != null)
                {
                    for (int j = 1; j <= lastCell.Address.ColumnNumber - 1; j++)
                    {
                        //String contents = currentRow.Cell(j).GetString();
                        var firp = currentRow.Cell(j).Value;
                        sheetValue[i, j - 1] = firp;
                        //data.Add(contents);
                    }
                }
                currentRow = currentRow.RowBelow();
            }

            metricsList.Add(sheetValue);

        }

        private void ExcelScanInternal(XLWorkbook workBookIn)
        {
            int numSheets = workBookIn.Worksheets.Count;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            //for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            //{
            //    var sheet = workBookIn.Worksheet(sheetNum);
            //    prepDashboardSheets(sheet);
           // }
            var sheet = workBookIn.Worksheet(5); //preventive
            prepDashboardSheets(sheet);


        }

        private void prepDashboardSheets(IXLWorksheet sheet)
        {
            //testing for now, try to scan for length of excel file (x =1, y= ?) 
            //if 2 blank cells in a row then last filled cell is end of report and only have to 
            //loop through that many cells to make sure I have inserted data for all providers.
            _worksheet = sheet;
            List<int> newRowList = new List<int>(); //use this to keep track of rows that need a new row so
            Dictionary<String, int> providerRowList = new Dictionary<String, int>();// use this to keep track of where each provisders new row is so we can insert metrics data into it

            //we can add new rows after checkign each row and thus not ad 285 new rows.

            String lastValue = ""; //use this to keep track of the last string that way on Agancy Metrics we can just delete the current row

            Dictionary<String, IXLRangeRow> providerRows = new Dictionary<String, IXLRangeRow>();
            //use this one to keep track of provider name and row number. then send it to metrics to dashboard and do accordignly


            if (_worksheet != null)
            {
                var firstCell = _worksheet.FirstCellUsed();
                var lastCell = _worksheet.LastCellUsed();
                _worksheet_range = _worksheet.Range(firstCell.Address, lastCell.Address);
                
                if (_worksheet_range != null)
                {
                    int nRows = _worksheet_range.RowCount();
                    int nCols = _worksheet_range.ColumnCount();
                    for (int i = 1; i < nRows + 1; i++)
                    {
                        var row = _worksheet_range.Row(i);
                        var newRow = _worksheet_range.Row(i + 1);
                        string value = row.Cell(1).Value as string;
                        
                        if (value != null)
                        {
                            //have to do the following because the depression sheet on dashboard does not follow same conventionm as others
                            if (_worksheet == _dashboard.Worksheet(2) && value.Contains("Agency"))
                            {
                                newRow = _worksheet_range.Row(i + 2);
                                newRow.InsertRowsBelow(1);
                                var blankRow = _worksheet_range.Row(i + 3);
                                blankRow.Style.NumberFormat.NumberFormatId = 0;
                                providerRows.Add(value, blankRow);
                               // i++;
                                lastValue = value;
                                continue;
                            }
                            //the following should take care of every other sheet for the agency metrics row
                             if (value == "Month")
                             {

                                 newRow = _worksheet_range.Row(i + 1);//this gets us int he right area and then insert the row above
                                 
                                 newRow.InsertRowsBelow(1); //try to insert rows after we have metrics and tehn insert metrics into cells then insert row
                                 var blankRow = _worksheet_range.Row(i + 2);
                                 blankRow.Style.NumberFormat.NumberFormatId = 0;
                               //  blankRow.Style.Fill.BackgroundColor = XLColor.Blue;
                                 providerRows.Add(lastValue, blankRow);
                                 //i++;
                                 lastValue = value;
                                 continue;
                             }
                            lastValue = value;
                        }
                    }
                   
                   MetricsToDashboard(providerRows, _worksheet);//figure out what we need to send to this method , worksheet, provider / row dict etc.

                }
            }
        }
    }
}
