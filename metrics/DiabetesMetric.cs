using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ProviderDashboards.metrics
{
    class DiabetesMetric
    {
        List<XLWorkbook> workbooks;
        XLWorkbook workbook;
        List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        List<object> metrics = new List<object>(); // this is the actuall contnets of the metrics locations
        List<String> metricNames = new List<String>();
        String provider;

        public List<object> Metrics { get { return metrics; } set { return; } }

        public DiabetesMetric(String provider, List<XLWorkbook> workbooks)
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
            //Diabetes Care Management_Metrics
            metricNames.Add("Total # of diabetic patients: "); //# X+4
            metricNames.Add("HgbA1c x1:");//% X+3
            metricNames.Add("Average Hgb A1c:"); //# X+4
            metricNames.Add("HgbA1c < 7:"); //% X+3
            metricNames.Add("Self Care Plan:");//% X+4
            metricNames.Add("BP <130/80:");//% X+2
            metricNames.Add("Eye Exam12 Months:");//% X+3
            metricNames.Add("Eye Exam18 Months:");//% X+3
            metricNames.Add("Annual Foot Exam:");//% X+4
            metricNames.Add("LDL <100:");//% X+ 4

            //Diabetes Measure: HgbA1c x2 Non-Compliance_Metrics
            if (provider == "Agency")
                metricNames.Add("Total # of Patients with two HgbA1c's during the past year and at least 90 days apart:"); //% X+1 = provider X + 13 = agency
            else //every one else
                metricNames.Add("Percent:");

            //Diabetes Measure: High Risk DRE Non-Compliance_1 Year_Metrics
            metricNames.Add("Percent with DRE done within past year:");// % X+5 /X+6
            metricNames.Add("Percent  with DRE done within past 18 months:");// % X+5 / X+6
            metricNames.Add("Total # of  High Risk Patients:");//# X+5/X+4

            //Diabetes Measure: Statin Non-Compliance_Metrics
            metricNames.Add("Percent:");//%  x+2/X+1

            //Diabetes Measure: ACEI_ARB Non-Compliance_Metrics
            metricNames.Add("Percent:");//%  X+2/X+1
        }

        private void findProvderName()
        {
            Point providerLocation = new Point(0, 0);
            for (int metricNumber = 0; metricNumber < 16; metricNumber++)//total numebr of metrics = 15
            {
                if (metricNumber < 10)
                    workbook = workbooks[0];
                else if (metricNumber == 10)
                    workbook = workbooks[1];
                else if (metricNumber > 10 && metricNumber < 14)
                    workbook = workbooks[2];
                else if (metricNumber == 14)
                    workbook = workbooks[3];
                else if (metricNumber == 15)
                    workbook = workbooks[4];

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
            bool agency = provider == "Agency";
            var sheet = workbook.Worksheet(1);
            String metricName = ""; //use this to match the cell from worksheet on which to grab the metrics data
            int xOffset = 0; //use this to tell the program how many cells on x to move over to find desired data
            //start the actual retrieval of data here
            switch (metricNumber)
            {
                case 0:
                    metricName = metricNames[0];
                    xOffset = 4;
                    break;
                case 1:
                    metricName = metricNames[1];
                    xOffset = 3;
                    break;
                case 2:
                    metricName = metricNames[2];
                    xOffset = 4;
                    break;
                case 3:
                    metricName = metricNames[3];
                    xOffset = 3;
                    break;
                case 4:
                    metricName = metricNames[4];
                    xOffset = 4;
                    break;
                case 5:
                    metricName = metricNames[5];
                    xOffset = 2;
                    break;
                case 6:
                    metricName = metricNames[6];
                    xOffset = 3;
                    break;
                case 7:
                    metricName = metricNames[7];
                    xOffset = 3;
                    break;
                case 8:
                    metricName = metricNames[8];
                    if (agency)
                        xOffset = 3;
                    else
                        xOffset = 4;
                    break;
                case 9:
                    metricName = metricNames[9];
                    xOffset = 4;
                    break;
                case 10:
                    metricName = metricNames[10];
                    if (agency)
                        xOffset = 13;
                    else
                        xOffset = 1;
                    break;
                case 11:
                    metricName = metricNames[11];
                    if (agency)
                        xOffset = 6;
                    else
                        xOffset = 5;
                    break;
                case 12:
                    metricName = metricNames[12];
                    if (agency)
                        xOffset = 6;
                    else
                        xOffset = 5;
                    break;
                case 13:
                    metricName = metricNames[13];
                    if (agency)
                        xOffset = 4;
                    else
                        xOffset = 5;
                    break;
                case 14:
                    metricName = metricNames[14];
                    if (agency)
                        xOffset = 1;
                    else
                        xOffset = 2;
                    break;
                case 15:
                    metricName = metricNames[15];
                    if (agency)
                        xOffset = 1;
                    else
                        xOffset = 2;
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
                        if (metricNumber == 0 || metricNumber == 2 || metricNumber == 13)
                        {
                            metrics.Add(value); //just need a straight number
                        }
                        else //need a p[ercent
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
