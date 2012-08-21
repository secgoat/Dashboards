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
            int[] diabetesFileLocations = new int[] { 8, 9, 10, 11, 12 };

            foreach (int fileNum in diabetesFileLocations)
            {
                diabetesFile = new XLWorkbook(metrics_files[fileNum]);
                diabetesFiles.Add(diabetesFile);
            }
            /* foreach(provider in providers)
             *  then check to see if providerrows.Keys contains provider
             *  if true then row = providerRows.Element.Key.Contains(provider)
             */
            foreach (String provider in providers)
            {
                DiabetesMetric diabetes = new DiabetesMetric(provider, diabetesFiles);
                List<object> metrics = diabetes.Metrics;
             //TODO: maybe use lambdas to return the appropriate row here?
        
                var row = providerRows.Any(p => p.Key.Contains(provider));
                if (row != null)
                {

                }
               
            }
            for (int i = 0; i < providers.Count; i++)
            {
                DiabetesMetric diabetes = new DiabetesMetric(providers[i], diabetesFiles);
                List<object> metrics = diabetes.Metrics;
                //TODO: fix the problem where Phil and Sarah get skipped because we are counting by provider and not matching names
                //      also fix the diabetes rows, counting by numbe rand need to match names, Loren and Jessica are not on diabetes reports so the nbumber sin the loops below are off.
                var row = providerRows.ElementAt(i).Value; // this is why Phil and sarah are getting skipped, put their metrics in clair and linda stephens instead. need to match names  instea dof wokrign by number
                for (int x = 1; x <= metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(diabetes.Metrics.ElementAt(x - 1));
                }
            }
        }

        private void DepressionToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook depressionFile;
            List<XLWorkbook> depressionFiles = new List<XLWorkbook>();
            int[] depressionFileLocations = new int[] { 4, 5, 6, 7 };

            foreach (int fileNum in depressionFileLocations)
            {
                depressionFile = new XLWorkbook(metrics_files[fileNum]);
                depressionFiles.Add(depressionFile);
            }
            for (int i = 0; i < providers.Count; i++)
            {
                DepressionMetric depression = new DepressionMetric(providers[i], depressionFiles);
                List<object> metrics = depression.Metrics;
                //TODO: fix the problem where Phil and Sarah get skipped because we are counting by provider and not matching names
                var row = providerRows.ElementAt(i).Value; // this is why Phil and sarah are getting skipped, put their metrics in clair and linda stephens instead. need to match names  instea dof wokrign by number
                for (int x = 1; x <= metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(depression.Metrics.ElementAt(x - 1));
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
            for (int i = 0; i < providers.Count; i++)
            {
                CardiovascularMetric cardiovascular = new CardiovascularMetric(providers[i], cardiovascularFiles);
                List<object> metrics = cardiovascular.Metrics;
                //TODO: fix the problem where Phil and Sarah get skipped because we are counting by provider and not matching names
                var row = providerRows.ElementAt(i).Value; // this is why Phil and sarah are getting skipped, put their metrics in clair and linda stephens instead. need to match names  instea dof wokrign by number
                for (int x = 1; x <= metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(cardiovascular.Metrics.ElementAt(x - 1));
                }
            }
        }

        private void AsthmaToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook asthmaFile;
            List<XLWorkbook> AsthmaFiles = new List<XLWorkbook>();
            int[] asthmaFileLocations = new int[] {1, 2};

            foreach (int fileNum in asthmaFileLocations)
            {
                asthmaFile = new XLWorkbook(metrics_files[fileNum]);
                AsthmaFiles.Add(asthmaFile);
            }
            for (int i = 0; i < providers.Count; i++)
            {
                AsthmaMetric asthma = new AsthmaMetric(providers[i], AsthmaFiles);
                List<object> metrics = asthma.Metrics;
                //TODO: fix the problem where Phil and Sarah get skipped because we are counting by provider and not matching names
                var row = providerRows.ElementAt(i).Value; // this is why Phil and sarah are getting skipped, put their metrics in clair and linda stephens instead. need to match names  instea dof wokrign by number
                for (int x = 1; x <= metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(asthma.Metrics.ElementAt(x - 1));
                }
            }
           
        }

        private void PreventiveToDashboard(Dictionary<String, IXLRangeRow> providerRows)
        {
            XLWorkbook preventiveFile;
            List<XLWorkbook> PreventiveFiles = new List<XLWorkbook>();
            int[] preventivFileLocations = new int[] { 3, 3, 16, 15, 17, 18, 0 };//use 3 twice so we can get the 2 metrics from BMI file
            //this allows the loop in the preventive metrics object to iterate over file 3 2 times with different results

            foreach (int fileNum in preventivFileLocations)
            {
                preventiveFile = new XLWorkbook(metrics_files[fileNum]);
                string name = preventiveFile.Properties.Title;
                PreventiveFiles.Add(preventiveFile);
            }

            for (int i = 0; i < providers.Count; i++)
            {
                PreventiveMetric preventive = new PreventiveMetric(providers[i], PreventiveFiles);
                List<object> metrics = preventive.Metrics;
                //TODO: fix the problem where Phil and Sarah get skipped because we are counting by provider and not matching names
                var row = providerRows.ElementAt(i).Value; // this is why Phil and sarah are getting skipped, put their metrics in clair and linda stephens instead. need to match names  instea dof wokrign by number
                for (int x = 1; x <= metrics.Count; x++)
                {
                    var cell = row.Cell(x);
                    if (x == 1) { cell.Style.NumberFormat.NumberFormatId = 17; }//mmm-yy
                    else { cell.Style.NumberFormat.NumberFormatId = 10; }//0.00%
                    //cell.Style.NumberFormat.Format = "@";
                    cell.Style.Font.SetBold(false);
                    cell.SetValue(preventive.Metrics.ElementAt(x - 1));
                }
            }
        }


    }
}
