using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

//TODO column C(metric 1) is the same for all providers definitely wrong, appears right for Agency.
namespace ProviderDashboards.metrics
{
    class PreventiveMetric
    {
        List<XLWorkbook> workbooks;
        XLWorkbook workbook;
        List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        List<object> metrics = new List<object>(); // this is the actuall contnets of the metrics locations
        List<String> metricNames = new List<String>();
        String provider;

        public List<object> Metrics { get { return metrics; } set{ return;  } }

        public PreventiveMetric(String provider, List<XLWorkbook> workbooks)
        {
            this.provider = provider;
            this.workbooks = workbooks;
            setmetricNames();
            findProvderName();
            //use this so date is only inserted once, couldnt get it right in the find provider loops
            System.DateTime now = DateTime.Today;
            String monthYear = now.ToString("MMM-yy");
            metrics.Insert(0, monthYear);
        }

        /// <summary>
        /// make  list of metrics names so we can match provider
        /// and then metric to get the x,y location of metrics that need to be moved across
        /// </summary>
        private void setmetricNames()
        {
            /* the array of files goes like this:
             * 
             * 0: CHD BMI
             * 1: CHD BMI
               2: Cervical Cancer
             * 3: Breast Cancer
             * 4: Colon Cancer
             * 5: pneumococcal
               6: Advanced Directives
             */


            metricNames.Add("Total # of eligible patients with a documented BMI percentile >= 85 within the past year: ");//percent (y+14) BMI
            metricNames.Add(">= 85 who received 5-2-1-0");//percent (x+14) BMI
           
            metricNames.Add("Percent: "); //precent (x+1) CERVICAL       
            
            metricNames.Add("Percent:");//percent (x+1) BREAST
            
            metricNames.Add("Percent: "); //percent (x+1) COLON
            
            metricNames.Add("Percent:");//percent(x+2) PNEUMOCOCOAL

            metricNames.Add("Total # of patients with an Advance Directive on File: "); //percent (x+10) ADVANCED DIRECTIVES 
        }

        private void findProvderName()
        {
            Point providerLocation = new Point(0, 0);
            //int fileNumber = 0;
            for (int fileNumber = 0; fileNumber < workbooks.Count; fileNumber++)
            {
                workbook = workbooks[fileNumber];
                var sheet = workbook.Worksheet(1);
                var colRange = sheet.Range("A:A");
                foreach (var cell in colRange.CellsUsed())
                {
                    if (cell.Value != null)
                    {
                        String value = (String)cell.Value;
                        int cellRow = cell.Address.RowNumber;
                        if (value.Contains(provider))
                        {
                            providerLocation = new Point(1, cellRow);
                            setMetricDataLocations(providerLocation, fileNumber);
                            break;
                        }
                    }

                }
               
            }
            
            //metrics.Insert(0, now.Month +"-" + now.Year);
           
        }


        private void setMetricDataLocations(Point providerLocation, int fileNumber)
        {

            var sheet = workbook.Worksheet(1);
            String metricName = ""; //use this to match the cell from worksheet on which to grab the metrics data
            int xOffset = 0; //use this to tell the program how many cells on x to move over to find desired data
            //start the actual retrieval of data here
            switch (fileNumber)
            {
                case 0:
                    metricName = metricNames[0];
                    xOffset = 14;
                    break;
                case 1:
                    metricName = metricNames[1];
                    xOffset = 14;
                    break;
                case 2:
                    metricName = metricNames[2];
                    xOffset = 2;
                    break;
                case 3:
                    metricName = metricNames[3];
                    xOffset = 1;
                    break;
                case 4:
                    metricName = metricNames[4];
                    xOffset = 1;
                    break;
                case 5:
                    metricName = metricNames[5];
                    xOffset = 2;
                    break;
                case 6:
                    metricName = metricNames[6];
                    xOffset = 10;
                    break;
            }

            var providerRow = sheet.Row(providerLocation.Y); //get the location row of the matched provider name
            var curRow = providerRow; // use this as an iteratior to step trouhh the rows below provder row
            int lastRow = sheet.LastRowUsed().RowNumber(); //BAM! find the last row used
            var lastCell = curRow.LastCellUsed(); //last cell gets set to nnull sometimes?

            while (lastCell == null)
            {
                curRow = curRow.RowBelow();
                lastCell = curRow.LastCellUsed();
            }
            for (int r = curRow.RowNumber(); r < lastRow; r++)
            {
                lastCell = curRow.LastCellUsed();
                for (int c = 1; c < lastCell.Address.ColumnNumber; c++)//this does too many, maybe just search for the next 10 rows?
                {
                    var firp = curRow.Cell(c).Value.ToString();
                    if (firp != "")
                    {
                        if (Strings.Match(metricName, firp))
                        {
                            var value = curRow.Cell(c + xOffset).Value;
                            double percentValue = (double)value / 100;
                            metrics.Add(percentValue);
                            return;
                        }
                    }
                 }
                curRow = curRow.RowBelow();
            }
        }
    }
}
