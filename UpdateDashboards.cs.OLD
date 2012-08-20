using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

using System.Diagnostics;


namespace ProviderDashboards
{
    class UpdateDashboards
    {
        private Excel.Application app = null;
        private Excel.Workbook dashboard = null;
        private Excel.Workbook metrics = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range workSheet_range = null;

        private string[] metricsFiles;
        //use these to store the data from all the sheets
        //i think the numbering goes liek this for the sheets:
        List<object[,]> dashboardArray = new List<object[,]>(); // should be 0 -4 for the following
        Dictionary<MetricsTitles, object[,]> metricsArrays = new Dictionary<MetricsTitles, object[,]>();
        List<object[,]> diabetesArray = new List<object[,]>();  //workbook.sheet[1]
        List<object[,]> depressionArray = new List<object[,]>();//workbook.sheet[2]
        List<object[,]> asthmaArray = new List<object[,]>();    //workbook.sheet[3]
        object[,] cardiovascularArray;                          //workbook.sheet[4]
        List<object[,]> preventiveArray = new List<object[,]>();//workbook.sheet[5]
        //------------------------------------------------//
        private List<String> providers;
        private enum MetricsTitles
        {
            Asthma,
            AsthmaLeukotriene,
            CardiovascularCare,
            CardiovascularMeasureBP,
            Depression5PointCSD,
            DepressionMajorCSD1Year,
            DepressionFU_PHQ_CSD,
            DepressionCareManagementMetrics,
            DiabetesCareManagement,
            DiabetesACEI_ARB,
            DiabetesHGBA1C,
            DiabetesDRE,
            DiabetesStatin,
            AdvancedDirecttives,
            CHDPerformanceMeasure3BMIPediatric,
            BreastCancer,
            CervicalCancer,
            ColonCancer,
            PneumococcalImmu,
        };
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="providerNames"></param>
        public UpdateDashboards()
        {
            //this.providers = providerNames; don't really need to use the provider names from the settings, that is more for makign a new one

        }
        
        /// <summary>
        /// open the excel files and parse them into arrays and lists for easier 
        /// data manipulation
        /// </summary>
        /// <param name="dashboardFile"></param>
        /// <param name="metricsFolder"></param>
        public void openExcel(String dashboardFile, String metricsFolder)
        {
            //grab all the files in the metrics folder
            metricsFiles = Directory.GetFiles(metricsFolder);
            if(File.Exists("Copy.xlsx"))
            {
                File.Delete("Copy.xlsx");
            }
            File.Copy(dashboardFile, "copy.xlsx");

            //try openeing the excel sheets listed in metricsFile
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                //dashboard = app.Workbooks.Add();
                dashboard = app.Workbooks.Open(dashboardFile,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing); 


                for (int i = 0; i < metricsFiles.Length; i++)
                {
                    metrics = app.Workbooks.Open(metricsFiles[i],
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                    
                    //this is where we read all the data into some sort of array
                    MetricsToDictArray(metrics);
                    metrics.Close(false, metricsFiles[i], null);
                }

                ExcelScanInternal(dashboard);
                
                MetricsToDashboard();
                //cleanup
                //dashboard.Close(false, dashboardFile, null);
               
            }
            catch (Exception ex) 
            {
                Debug.WriteLine(ex);
            }

        }

        private void MetricsToDashboard()
        {
            
        }

        private void MetricsToDictArray(Excel.Workbook metricsFile)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)metricsFile.Sheets[1];

                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                Excel.Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);

            if (metrics.Name.Contains("Asthma"))
            {
                if (metrics.Name.Contains("Leukotriene"))
                {
                    metricsArrays.Add(MetricsTitles.AsthmaLeukotriene, valueArray);
                    return;
                }
                //this should only get added if the title has asthma but not leukotriene
                    metricsArrays.Add(MetricsTitles.Asthma, valueArray);
                    return;

            }//end asthma checks

            if (metrics.Name.Contains("Cardiovascular"))
            {
                if(metrics.Name.Contains("BP"))
                {
                    metricsArrays.Add(MetricsTitles.CardiovascularMeasureBP, valueArray);
                    return;
                }
                //this should be added if title has Cardiovascular but not BP
                metricsArrays.Add(MetricsTitles.CardiovascularCare, valueArray);
                return;
            }//end cardiovadscular checks

            if (metrics.Name.Contains("Depression"))
            {
                if(metrics.Name.Contains("5 Point"))
                {
                    metricsArrays.Add(MetricsTitles.Depression5PointCSD, valueArray);
                    return;
                }
                if(metrics.Name.Contains("Major_CSD"))
                {
                    metricsArrays.Add(MetricsTitles.DepressionMajorCSD1Year, valueArray);
                    return;
                }
                if(metrics.Name.Contains("Major_FU"))
                {
                    metricsArrays.Add(MetricsTitles.DepressionFU_PHQ_CSD, valueArray);
                    return;
                }
                //last case means it is just depression care metrics and 
                metricsArrays.Add(MetricsTitles.DepressionCareManagementMetrics, valueArray);
                return;
            }//end depression checks

            if (metrics.Name.Contains("Diabetes"))
            {
                if (metrics.Name.Contains("ACEI_ARB"))
                {
                    metricsArrays.Add(MetricsTitles.DiabetesACEI_ARB, valueArray);
                    return;
                }
                if (metrics.Name.Contains("HgbA1c"))
                {
                    metricsArrays.Add(MetricsTitles.DiabetesHGBA1C, valueArray);
                    return;
                }
                if (metrics.Name.Contains("DRE"))
                {
                    metricsArrays.Add(MetricsTitles.DiabetesDRE, valueArray);
                    return;
                }
                if (metrics.Name.Contains("Statin"))
                {
                    metricsArrays.Add(MetricsTitles.DiabetesStatin, valueArray);
                    return;
                }
                metricsArrays.Add(MetricsTitles.DiabetesCareManagement, valueArray);
                return;
            }//end diabeetus checks

            //start checks for Preventive care, these are all 1 if statement
            if (metrics.Name.Contains("Directives"))
            {
                metricsArrays.Add(MetricsTitles.AdvancedDirecttives, valueArray);
                return;
            }
            if (metrics.Name.Contains("BMI"))
            {
                metricsArrays.Add(MetricsTitles.CHDPerformanceMeasure3BMIPediatric, valueArray);
                return;
            }
            if (metrics.Name.Contains("Breast"))
            {
                metricsArrays.Add(MetricsTitles.BreastCancer, valueArray);
                return;
            }
            if (metrics.Name.Contains("Cervical"))
            {
                metricsArrays.Add(MetricsTitles.CervicalCancer, valueArray);
                return;
            }
            if (metrics.Name.Contains("Colon"))
            {
                metricsArrays.Add(MetricsTitles.ColonCancer, valueArray);
                return;
            }
            if (metrics.Name.Contains("Pneumococcal"))
            {
                metricsArrays.Add(MetricsTitles.PneumococcalImmu, valueArray);
                return;
            }
            //aaaaaand thats a wrap!
           int num =  metricsArrays.Count;
        }

        private void ExcelScanInternal(Excel.Workbook workBookIn)
        {
            /// <summary>
            /// Scan the selected Excel workbook and store the information in the cells
            /// for this workbook in an object[,] array. Then, call another method
            /// to process the data.
            /// 
            /// Shoudl only send the dashboard file here, it will then iterate through each sheet, send
            /// data array to new method to pull the provider names and insert new rows.
            /// come back and pull corresponding metrics files within a new method and then populate the empty 
            /// rows in the dashboard with the metrics info.
            /// </summary>
            //
            // 
            //
            int numSheets = workBookIn.Sheets.Count;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)workBookIn.Sheets[sheetNum];

                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
               /* Excel.Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);
                dashboardArray.Add(valueArray);
                */
                //
                // need to take the data from each sheet and process it and transfer over to the dashboard excel sheet here
                //
                sheet.Activate();
                prepDashboardSheets(sheet);
            }

        }

        private void prepDashboardSheets(Excel.Worksheet sheet)
        {
            //finally what i need to do is scan each sheet, and keep track of eachnew row and who it is assigned to, then goto the metrics
            //arrays and import the appropriate data.

            // here is my third attempt and was really well done, except soem random instances where the
             // line was inserted in the wrong position or a provider was skipped because names did not match 1-00%
              
            //testing for now, try to scan for length of excel file (x =1, y= ?) 
            //if 2 blank cells in a row then last filled cell is end of report and only have to 
            //loop through that many cells to make sure I have inserted data for all providers.
            worksheet = sheet;
            List<int> newRowList = new List<int>(); //use this to keep track of rows that need a new row so
            //we can add new rows after checkign each row and thus not ad 285 new rows.
            
            String lastValue = ""; //use this to keep track of the last string that way on Agancy Metrics we can just delete the current row
            
            Dictionary<String, Excel.Range> providerRows = new Dictionary<string,Excel.Range>();
            //use this one to keep track of provider name and row number. then send it to metrics to dashboard and do accordignly

            
            if (worksheet != null)
            {
                workSheet_range = worksheet.UsedRange;
                if(workSheet_range != null)
                {
                    int nRows = workSheet_range.Rows.Count;
                    int nCols = workSheet_range.Columns.Count;
                    for(int i = 1; i < nRows + 1; i++)
                    {
                        Excel.Range row = worksheet.Cells[i, 1];
                        string value = row.Cells[1].Value as string;
                       /*
                        //have to do the following because the depression sheet on dashboard does not follow same conventionm as others
                        if (worksheet == dashboard.Worksheets[2] && value.Contains("Agency"))
                        {
                            newRowList.Add(i + 4);
                            Excel.Range clearRow = worksheet.Cells.EntireRow[i + 4];
                            clearRow.ClearContents();
                            continue;
                        } */
                        //the following should take care of every other sheet for the agency metrics row
                        if(value == "Month")
                        {
                            /*if(lastValue.Contains("Agency"))
                            {
                                newRowList.Add(i+2);
                                Excel.Range clearRow = worksheet.Cells.EntireRow[i+2];
                                clearRow.ClearContents();
                                lastValue = value;
                                continue;
                            }*/

                            newRowList.Add(i+2);
                            //Excel.Range newRow = worksheet.Cells.EntireRow[i + 2];
                            //Excel.Range entireRow = newRow.EntireRow[i + 2];
                            //providerRows.Add(lastValue, newRow);
                            
                        }
                        lastValue = value;

                    }
                    int x = 0;
                    foreach (int cell in newRowList)
                    {
                        
                        Excel.Range newRow = worksheet.Cells[cell + x];
                        Excel.Range entireRow = newRow.EntireRow[cell+ x];
                        entireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                        x++;
                    }
                   // MetricsToDashboard();//figure out what we need to send to this method , worksheet, provider / row dict etc.
                    
                }
            }
        }


        public void addData(int row, int col, string data, string format)
        {
            worksheet.Cells[row, col] = data;
            workSheet_range.WrapText = true;
            workSheet_range.NumberFormat = format;

        }
    }
}
