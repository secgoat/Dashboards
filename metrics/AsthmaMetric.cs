using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ProviderDashboards.metrics
{
    class AsthmaMetric
    {
        List<XLWorkbook> workbooks;
        XLWorkbook workbook;
        List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        List<object> metrics = new List<object>(); // this is the actuall contnets of the metrics locations
        List<String> metricNames = new List<String>();
        String provider;

        public List<object> Metrics { get { return metrics; } set { return; } }

        public AsthmaMetric(String provider, List<XLWorkbook> workbooks)
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
            //from Ashtma care management metrics
            metricNames.Add("Total # of patients with persistent + unclassified asthma diagnosis: "); //# X+8
            metricNames.Add("Total # of patients with a classified asthma diagnosis: ");//% X+10
            metricNames.Add("Total # of patients with persistent asthma: "); //# X+6
            metricNames.Add("Total # of  patients w/persistent or unclassified asthma with a documented ED visit w/in past year: "); //% X+18
            metricNames.Add("Total # of  patients w/ persistent or unclassified asthma with a documented Asthma Action Plan within the past year: ");//% X+18
            metricNames.Add("Total # of patients w/ persistent or unclassified asthma with an annual Asthma review: ");//% X+18
            //from Leukotrine
            metricNames.Add("Total # of patients on an inhaled corticosteroid or leukotriene inhibitor: "); //% X+3
        }

        private void findProvderName()
        {
            Point providerLocation = new Point(0, 0);
            //int fileNumber = 0;
            for (int metricNumber = 0; metricNumber < 7; metricNumber++)//total numebr of metrics = 6
            {
                if (metricNumber == 6)
                    workbook = workbooks[1];
                else
                    workbook = workbooks[0];

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
                    xOffset = 8;
                    break;
                case 1:
                    metricName = metricNames[1];
                    xOffset = 10;
                    break;
                case 2:
                    metricName = metricNames[2];
                    xOffset = 6;
                    break;
                case 3:
                    metricName = metricNames[3];
                    xOffset = 18;
                    break;
                case 4:
                    metricName = metricNames[4];
                    xOffset = 18;
                    break;
                case 5:
                    metricName = metricNames[5];
                    xOffset = 18;
                    break;
                case 6:
                    metricName = metricNames[6];
                    xOffset = 3;
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
                        if (metricNumber == 0 || metricNumber == 2)
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
