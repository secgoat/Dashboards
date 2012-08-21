using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ProviderDashboards.metrics
{
    class DepressionMetric
    {
        List<XLWorkbook> workbooks;
        XLWorkbook workbook;
        List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        List<object> metrics = new List<object>(); // this is the actuall contnets of the metrics locations
        List<String> metricNames = new List<String>();
        String provider;

        public List<object> Metrics { get { return metrics; } set { return; } }

        public DepressionMetric(String provider, List<XLWorkbook> workbooks)
        {
            this.provider = provider;
            this.workbooks = workbooks;
            setmetricNames();
            findProvderName();
            //use this so date is only inserted once, couldnt get it right in the find provider loops
            System.DateTime now = DateTime.Today;
            metrics.Insert(0, now);
        }

        /// <summary>
        /// make  list of metrics names so we can match provider
        /// and then metric to get the x,y location of metrics that need to be moved across
        /// </summary>
        private void setmetricNames()
        {
            //x == provider name
            //Depression care management metrics
            metricNames.Add("Total # of patients with a depression dx managed at ACHS:"); //# X+7
            metricNames.Add("Total # of CSD patients managed at ACHS:");//# X+6
            metricNames.Add("% of CSD patients managed at ACHS with 50% reduction in score:"); //% X+8
            metricNames.Add("% of patients managed at ACHS with a follow-up PHQ-9 done w/in last 6 months:"); //% X+9
            metricNames.Add("% of patients managed at ACHS with a documented SCP within the past year: ");//% X+9
            metricNames.Add("% of eligible CSD patients  managed at ACHS who have achieved remission:");//% X+9

            //Depression Care Management_Major_CSD 1 Year
            metricNames.Add("Total # of patients:"); //# X+7
            metricNames.Add("Total # of patients with a documented follow-up within 1-3 weeks:");//% X+8
            metricNames.Add("Total # of patients with a documented SCP within the past year:");//% X+8
            metricNames.Add("Total # of  eligible patients who have achieved remission: ");//% X+7

            //Depression Care Management_Major_5 Point_CSD 1 Year_v2
            metricNames.Add("Percent:");//% X+1

            //Depression Care Management_Major_FU PHQ_CSD 1 Year
            metricNames.Add("Total # of patients with a 3-12 week PHQ-9 score:");//%X+6
            metricNames.Add("Total # of patients with a 4-8 week PHQ-9 score:");//% X+6
        }

        private void findProvderName()
        {
            Point providerLocation = new Point(0, 0);
            //int fileNumber = 0;
            for (int metricNumber = 0; metricNumber < 13; metricNumber++)//total numebr of metrics = 13
            {
                if (metricNumber < 6)
                    workbook = workbooks[0];
                else if (metricNumber >= 6 && metricNumber < 10)
                    workbook = workbooks[1];
                else if (metricNumber == 10)
                    workbook = workbooks[2];
                else if (metricNumber > 10)
                    workbook = workbooks[3];

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
                            setMetricDataLocations(providerLocation, metricNumber);
                            break;
                        }
                    }
                }
            }
        }

        private void setMetricDataLocations(Point providerLocation, int metricNumber)
        {

            var sheet = workbook.Worksheet(1);
            String metricName = ""; //use this to match the cell from worksheet on which to grab the metrics data
            int xOffset = 0; //use this to tell the program how many cells on x to move over to find desired data
            //start the actual retrieval of data here
            switch (metricNumber)
            {
                case 0:
                    metricName = metricNames[0];
                    xOffset = 7;
                    break;
                case 1:
                    metricName = metricNames[1];
                    xOffset = 6;
                    break;
                case 2:
                    metricName = metricNames[2];
                    xOffset = 8;
                    break;
                case 3:
                    metricName = metricNames[3];
                    xOffset = 9;
                    break;
                case 4:
                    metricName = metricNames[4];
                    xOffset = 9;
                    break;
                case 5:
                    metricName = metricNames[5];
                    xOffset = 9;
                    break;
                case 6:
                    metricName = metricNames[6];
                    xOffset = 7;
                    break;
                case 7:
                    metricName = metricNames[7];
                    xOffset = 8;
                    break;
                case 8:
                    metricName = metricNames[8];
                    xOffset = 8;
                    break;
                case 9:
                    metricName = metricNames[9];
                    xOffset = 7;
                    break;
                case 10:
                    metricName = metricNames[10];
                    xOffset = 1;
                    break;
                case 11:
                    metricName = metricNames[11];
                    xOffset = 6;
                    break;
                case 12:
                    metricName = metricNames[12];
                    xOffset = 6;
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
                    var firp = curRow.Cell(c).Value;
                    if (firp.ToString() == metricName)
                    {
                        var value = curRow.Cell(c + xOffset).Value;
                        if (metricNumber == 0 || metricNumber == 2 || metricNumber == 6)
                        {
                            metrics.Add(value); //just need a straight number
                        }
                        else //need a percent
                        {
                            double percentValue = (double)value / 100;
                            metrics.Add(percentValue);
                        }
                        return;
                    }
                }
                curRow = curRow.RowBelow();
            }
        }
    }
}
