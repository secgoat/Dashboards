using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ProviderDashboards.metrics
{
    class AsthmaMetric
    {
        XLWorkbook workbook;
        List<Point> metricDataLocations = new List<Point>(); // where is it located in the array
        List<String> metrics = new List<String>(); // this is the actuall contnets of the metrics locations
        List<String> metricNames = new List<String>();
        String provider;
        Dictionary<Point, String> dashboardMetrics = new Dictionary<Point, String>();
        List<object[,]> metricsArrays;

        public List<String> Metrics { get { return metrics; } set{ return;  } }

        public AsthmaMetric(String provider, XLWorkbook workbook)
        {
            this.provider = provider;
            this.workbook = workbook;
            setmetricNames();
            findProvderName();
        }

        /// <summary>
        /// make  list of metrics names so we can match provider
        /// and then metric to get the x,y location of metrics that need to be moved across
        /// </summary>
        private void setmetricNames()
        {
            //x == provider name
            //from Ashtma care management metrics
            metricNames.Add("Total # of patients with persistent + unclassified asthma diagnosis: "); //lookign for number (x+2)(y+9)
            metricNames.Add("Total # of patients with a classified asthma diagnosis: ");//looking for percentage (x+3)(y+10)
            metricNames.Add("Total # of patients with persistent asthma: "); //number (x+2)(Y+6)
            metricNames.Add("Total # of  patients w/persistent or unclassified asthma with a documented ED visit w/in past year: "); //percent (x+5)(y+18)
            metricNames.Add("Total # of  patients w/ persistent or unclassified asthma with a documented Asthma Action Plan within the past year: ");//percent(x+6)(y+18)
            metricNames.Add("Total # of patients w/ persistent or unclassified asthma with an annual Asthma review: ");//percent (x+7)(y+18)
            //from Leukotrine
            metricNames.Add("Total # of patients on an inhaled corticosteroid or leukotriene inhibitor: "); //percent (x+2)(y+13)
        }

        private void findProvderName()
        {
            Point providerLocation = new Point(0,0);
            int fileNumber = 0;

            var sheet = workbook.Worksheet(1);
            var colRange = sheet.Range("A:A");
            foreach(var cell in colRange.CellsUsed())
            {
                if (cell.Value != null)
                {
                    String value = (String)cell.Value;
                    int cellRow = cell.Address.RowNumber;
                    if (value.Contains(provider))
                    {
                        providerLocation = new Point(1, cellRow);
                        setMetricDataLocations(providerLocation, fileNumber);
                        fileNumber++;
                    }
                }
            
            }
         
            /*foreach (object[,] metricsFile in metricsArrays) //loop through once for each file
            {
                int boundsX = metricsFile.GetUpperBound(0);
                int boundsY = metricsFile.GetUpperBound(1);
                //now loop through the arrays
                for (int y = 0; y <= boundsY; y++)
                {
                    for (int x = 0; x <= boundsX; x++)
                    {
                        if (metricsFile[x, y] != null)
                        {
                            string name = metricsFile[x, y].ToString();
                            if (name.Contains(provider))
                            {
                                providerLocation = new Point(x, y);
                                break;
                            }
                        }
                    }
                } //end array loop
                setMetricDataLocations(providerLocation, fileNumber);
                fileNumber++;
            }*/

        }

        private void setMetricDataLocations(Point providerLocation, int fileNumber)
        {

            var sheet = workbook.Worksheet(1);
            String contents = "";
            //just goign to see if i canloop through each damn cell and look at the value
            int usedRows = sheet.RowsUsed().Count();

            var firstRowUsed = sheet.FirstRowUsed();
            var currentRow = firstRowUsed.RowUsed();

            int lastColumn = sheet.LastColumnUsed().ColumnNumber();
            int lastRow = sheet.LastRowUsed().RowNumber(); //BAM! find the last row used

            var lastCell = currentRow.LastCellUsed(); //last cell gets set to nnull sometimes?

            object[,] sheetValue = new object[lastRow, lastColumn];

            while (lastCell == null)
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

            
            if (fileNumber == 0) //from asthma care management
            {
                
                metrics.Add(sheet.Cell(providerLocation.X + 16, providerLocation.Y + 2).ToString());
                contents = sheet.Cell(providerLocation.X + 16, providerLocation.Y + 2).Value.ToString();

                metrics.Add(sheet.Cell(providerLocation.X + 18, providerLocation.Y + 3).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 18, providerLocation.Y + 3).ToString();

                metrics.Add(sheet.Cell(providerLocation.X + 6, providerLocation.Y + 2).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 6, providerLocation.Y + 2).ToString();

                metrics.Add(sheet.Cell(providerLocation.X + 18, providerLocation.Y + 5).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 18, providerLocation.Y + 5).ToString();

                metrics.Add(sheet.Cell(providerLocation.X + 18, providerLocation.Y + 6).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 18, providerLocation.Y + 6).ToString();

                metrics.Add(sheet.Cell(providerLocation.X + 18, providerLocation.Y + 7).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 18, providerLocation.Y + 7).ToString();
            }

            else //from lleukotrine
            {
                metrics.Add(sheet.Cell(providerLocation.X + 13, providerLocation.Y + 3).Value.ToString());
                contents = sheet.Cell(providerLocation.X + 13, providerLocation.Y + 3).Value.ToString();
            }
        }

       

    }
}
