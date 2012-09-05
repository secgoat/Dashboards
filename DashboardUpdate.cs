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
        IXLWorksheet _worksheet;
        IXLRange _worksheet_range;

        private string[] metrics_files;
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
            ///<summary> 
            /// use this to cut down the credentials form the provider names and just use first name
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

                //iterate thorugh all the dashboard sheets, prep them for new daata by adding new rows, then  extrract data from
                //metrics reports and insert into those sheets
                int numSheets = _dashboard.Worksheets.Count;
                for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
                    {
                        var sheet = _dashboard.Worksheet(sheetNum);
                        prepDashboardSheets(sheet);
                     }
                _dashboard.Save();
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void prepDashboardSheets(IXLWorksheet sheet)
        {
            /*.
             * can use htis for both types of sheets. 
             *  Regular: List.Count -1 (end of list: should be Month) List.Count -2 (last value: Should be provider name) 
             *  Diabetes: List.Count -1 (end of list: Month) List.Count -3 (blank space between provider name an dmonth on this one)
             */

            _worksheet = sheet;
            //use this one to keep track of provider name and row number. then send it to metrics to dashboard and do accordignly
            Dictionary<String, IXLRangeRow> providerRows = new Dictionary<String, IXLRangeRow>();

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
                            foreach (string provider in providers)
                            {
                                if (value.Contains(provider))
                                {
                                    if (_worksheet == _dashboard.Worksheet(2))//add a new row for the depression sheets
                                    {
                                        newRow = _worksheet_range.Row(i + 3);
                                        newRow.InsertRowsBelow(1);
                                        var blankRow = _worksheet_range.Row(i + 4);
                                        blankRow.Style.NumberFormat.NumberFormatId = 0;
                                        providerRows.Add(value, blankRow);

                                    }
                                    else //add a new row for every other sheet in the dashboard: Asthma, Diabetes, Cardiovascular, Preventive
                                    {
                                        newRow = _worksheet_range.Row(i + 2);//this gets us int he right area and then insert the row above
                                        newRow.InsertRowsBelow(1); //try to insert rows after we have metrics and tehn insert metrics into cells then insert row
                                        var blankRow = _worksheet_range.Row(i + 3);
                                        blankRow.Style.NumberFormat.NumberFormatId = 0;
                                        providerRows.Add(value, blankRow);
                                    }
                                    break; //break out of the foreach provider loop, we already found one, we wont find another match, time to go to the next row instead
                                }
                            }
                        }
                    }
                   MetricsToDashboard(providerRows, _worksheet);//figure out what we need to send to this method , worksheet, provider / row dict etc.
                }
            }
        }

        private void MetricsToDashboard(Dictionary<String, IXLRangeRow> providerRows, IXLWorksheet sheet)
        {
            _worksheet = sheet;

            if (_worksheet == _dashboard.Worksheet(1))
            {
                DiabetesToDashboard(providerRows);
            }
            else if (_worksheet == _dashboard.Worksheet(2))
            {
                DepressionToDashboard(providerRows);
            }
            else if (_worksheet == _dashboard.Worksheet(3))
            {
                AsthmaToDashboard(providerRows);
            }
            else if (_worksheet == _dashboard.Worksheet(4))
            {
                CardiovascularToDashboard(providerRows);
            }
            else if (_worksheet == _dashboard.Worksheet(5))
            {
                PreventiveToDashboard(providerRows);
            }
        }

        private void DiabetesToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook diabetesFile;
            List<XLWorkbook> diabetesFiles = new List<XLWorkbook>();
            int[] diabetesFileLocations = new int[] { 8, 10, 11, 12, 9 }; //order is pretty important here, because of the hardcode int he actual Metricsa Class file and the way it is searchinf for values. Just keep the order on all metrics the same as they are no unless osmehitng chnages like file name or what have you that might mess this order up.

            foreach (int fileNum in diabetesFileLocations)
            {
                diabetesFile = new XLWorkbook(metrics_files[fileNum]);
                diabetesFiles.Add(diabetesFile);
            }
           
            foreach (String provider in providers)
            {
                DiabetesMetric diabetes = new DiabetesMetric(provider, diabetesFiles);
                List<object> metrics = diabetes.Metrics;
             
                var kvp = providerRows.SingleOrDefault(s => s.Key.Contains(provider)); // find matching KVP by using linq
                var row = kvp.Value; //grap the row out of the value field for the matching KVP entry
                if (row != null)
                {
                   
                    for (int x = 1; x <= metrics.Count; x++)
                    {
                        var cell = row.Cell(x);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                        else if (x == 2 || x == 3 ) { cell.Style.NumberFormat.Format = "@"; }
                        else if (x == 5) { cell.Style.NumberFormat.Format = "0.0"; }
                        else { cell.Style.NumberFormat.Format = "0.0%"; }
                        //cell.Style.NumberFormat.Format = "@";
                        cell.Style.Font.SetBold(false);
                        cell.SetValue(metrics.ElementAt(x - 1));
                    }
                }
               
            }
            
        }

        private void DepressionToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook depressionFile;
            List<XLWorkbook> depressionFiles = new List<XLWorkbook>();
            int[] depressionFileLocations = new int[] { 7, 5, 4, 6 };

            foreach (int fileNum in depressionFileLocations)
            {
                depressionFile = new XLWorkbook(metrics_files[fileNum]);
                depressionFiles.Add(depressionFile);
            }
            foreach (String provider in providers)
            {
                DepressionMetric depression = new DepressionMetric(provider, depressionFiles);
                List<object> metrics = depression.Metrics;
                var kvp = providerRows.SingleOrDefault(s => s.Key.Contains(provider)); // find matching KVP by using linq
                var row = kvp.Value; //grap the row out of the value field for the matching KVP entry
                //if Keys do not contain provider name, then row will be null , and no sense in trying to populate!
                if (row != null)
                {
                    for (int x = 1; x <= metrics.Count; x++)
                    {
                        var cell = row.Cell(x);
                        if (x > 7) //do this to up the cell number it x => 7 because in the dashboard cell 7 is a blank line down th middle to seprate the two sides and we want to skip this column
                            cell = row.Cell(x + 1);

                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                        else if (x == 2 || x == 3 || x == 8) { cell.Style.NumberFormat.Format = "@";}
                        else { cell.Style.NumberFormat.Format = "0.0%"; }
                        //cell.Style.NumberFormat.Format = "@";
                        cell.Style.Font.SetBold(false);
                        cell.SetValue(metrics.ElementAt(x - 1));
                    }
                }

            }
        }

        private void CardiovascularToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook cardiovascularFile;
            List<XLWorkbook> cardiovascularFiles = new List<XLWorkbook>();
            int[] cardiovascularFileLocations = new int[] { 13 };

            foreach (int fileNum in cardiovascularFileLocations)
            {
                cardiovascularFile = new XLWorkbook(metrics_files[fileNum]);
                cardiovascularFiles.Add(cardiovascularFile);
            }

            foreach (String provider in providers)
            {
                CardiovascularMetric cardio = new CardiovascularMetric(provider, cardiovascularFiles);
                List<object> metrics = cardio.Metrics;
                var kvp = providerRows.SingleOrDefault(s => s.Key.Contains(provider)); // find matching KVP by using linq
                var row = kvp.Value; //grap the row out of the value field for the matching KVP entry
                if (row != null)
                {
                    for (int x = 1; x <= metrics.Count; x++)
                    {
                        var cell = row.Cell(x);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                        else if (x == 2) { cell.Style.NumberFormat.Format = "@";  }
                        else { cell.Style.NumberFormat.Format = "0.0%"; }
                        //cell.Style.NumberFormat.Format = "@";
                        cell.Style.Font.SetBold(false);
                        cell.SetValue(metrics.ElementAt(x - 1));
                    }
                }

            }
        }

        private void AsthmaToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook asthmaFile;
            List<XLWorkbook> asthmaFiles = new List<XLWorkbook>();
            int[] asthmaFileLocations = new int[] {1, 2};

            foreach (int fileNum in asthmaFileLocations)
            {
                asthmaFile = new XLWorkbook(metrics_files[fileNum]);
                asthmaFiles.Add(asthmaFile);
            }
            foreach (String provider in providers)
            {
                AsthmaMetric asthma = new AsthmaMetric(provider, asthmaFiles);
                List<object> metrics = asthma.Metrics;
                var kvp = providerRows.SingleOrDefault(s => s.Key.Contains(provider)); // find matching KVP by using linq
                var row = kvp.Value; //grab the row out of the value field for the matching KVP entry. this ensures the providert names are matched and no missing providers get data for some one who does exist

                if (row != null)
                {
                    for (int x = 1; x <= metrics.Count; x++)
                    {
                        var cell = row.Cell(x);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                        if (x == 2 || x == 4) { cell.Style.NumberFormat.Format = "@"; }
                        else { cell.Style.NumberFormat.Format = "0.0%"; }
                        cell.Style.Font.SetBold(false);
                        cell.SetValue(metrics.ElementAt(x - 1));
                    }
                }

            }
           
        }

        private void PreventiveToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook preventiveFile;
            List<XLWorkbook> preventiveFiles = new List<XLWorkbook>();
            int[] preventivFileLocations = new int[] { 3, 3, 16, 15, 17, 18, 0 };//use 3 twice so we can get the 2 metrics from BMI file
            //this allows the loop in the preventive metrics object to iterate over file 3 2 times with different results

            foreach (int fileNum in preventivFileLocations)
            {
                preventiveFile = new XLWorkbook(metrics_files[fileNum]);
                string name = preventiveFile.Properties.Title;
                preventiveFiles.Add(preventiveFile);
            }

            foreach (String provider in providers)
            {
                PreventiveMetric preventive = new PreventiveMetric(provider, preventiveFiles);
                List<object> metrics = preventive.Metrics;
                var kvp = providerRows.SingleOrDefault(s => s.Key.Contains(provider)); // find matching KVP by using linq
                var row = kvp.Value; //grap the row out of the value field for the matching KVP entry
                if (row != null)
                {
                    for (int x = 1; x <= metrics.Count; x++)
                    {
                        var cell = row.Cell(x);
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                        else { cell.Style.NumberFormat.Format = "0.0%"; }
                        //cell.Style.NumberFormat.Format = "@";
                        cell.Style.Font.SetBold(false);
                        cell.SetValue(metrics.ElementAt(x - 1));
                    }
                }

            }
        }


    }
}
