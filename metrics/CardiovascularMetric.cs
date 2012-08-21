using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ProviderDashboards.metrics
{
    class CardiovascularMetric
    {
        protected List<XLWorkbook> workbooks;
        protected XLWorkbook workbook;
        protected List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        protected List<object> metrics = new List<object>(); // this is the actuall contnets of the metrics locations
        protected List<String> metricNames = new List<String>();
        protected String provider;

        public List<object> Metrics { get { return metrics; } set { return; } }

         public CardiovascularMetric(String provider, List<XLWorkbook> workbooks)
        {
            this.provider = provider;
            this.workbooks = workbooks;
            setmetricNames();
            findProvderName();
            //use this so date is only inserted once, couldnt get it right in the find provider loops
            System.DateTime now = DateTime.Today;
            metrics.Insert(0, now);
        }

         private void findProvderName()
         {
             Point providerLocation = new Point(0, 0);

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
                         for (int metricNumber = 0; metricNumber < 8; metricNumber++)
                         {
                             setMetricDataLocations(providerLocation, metricNumber);
                         }
                         break;
                     }
                 }
             }
         }

         private void setmetricNames()
         {
             metricNames.Add("Total # of patients:");//# X+7
             metricNames.Add("Total # of patients with a documented LDL within the last 12 months:");//# X +14
             metricNames.Add("Total # of patients with a documented  LDL value < 100 within the last 12 months:");//% X+15
             metricNames.Add("Total # of patients with last BP < 140/90 within the past year:");//% X+14
             metricNames.Add("Total # with a documented smoking status within the past year: ");//% X+13
             metricNames.Add("Total # of smokers: ");//% X+11
             metricNames.Add("Total # of smokers who were counseled to quit within the past year:");//% X+13
             metricNames.Add("Total # of patients with a documented self care plan within the past year:");//% X+14
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
                     xOffset = 14;
                     break;
                 case 2:
                     metricName = metricNames[2];
                     xOffset = 15;
                     break;
                 case 3:
                     metricName = metricNames[3];
                     xOffset = 14;
                     break;
                 case 4:
                     metricName = metricNames[4];
                     xOffset = 13;
                     break;
                 case 5:
                     metricName = metricNames[5];
                     xOffset = 11;
                     break;
                 case 6:
                     metricName = metricNames[6];
                     xOffset = 13;
                     break;
                 case 7:
                     metricName = metricNames[7];
                     xOffset = 14;
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
                         if (metricNumber != 0)
                         {
                             double percentValue = (double)value / 100;
                             metrics.Add(percentValue);
                             return;
                         }
                         metrics.Add(value); // this should only be metric 0 = the # of patients every othe roen shoudlbe a %   
                         return;
                     }
                 }
                 curRow = curRow.RowBelow();
             }
         }
    }
}
