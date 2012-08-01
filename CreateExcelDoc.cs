/*
 * use this class when the Montlhy metrics excel worksheet does not exist
 * shoudl not be called every time, only needs to create the first time, or if it is missing etc.
 * 
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProviderDashboards
{
    class CreateExcelDoc
    {
        private Excel.Application app = null;
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range workSheet_range = null;

        //the following are used to hold the names of sheets, providers, metrics names etc.
        #region annoying excel header and cell names
        private List<String> sheetNames = new List<String>() { "Cardiovascular", "Asthma", "Depression", "Diabetes", "Preventive" };
        private List<String> providerNames = new List<String>()
        {
            "Agency Metrics",
            "Aaron Solnit, MD",
            "Alexandria Noble, APRN",
            "Barbara Ford, APRN",
            "Caitlin O'Donnell, MD",
            "Charles Wolcott, MD",
            "Danielle Beaulieu, PA-C",
            "David Nelson, DO",
            "Evelyn Hagan, APRN",
            "Nicole Fischler, APRN",
            "Philip Lawson, MD",
            "Sarah Young-Xu, MD"
        };

        private List<String> depessionMetricSubHeaders = new List<String>()
        {
            "All Patient with a Diagnosis of Depression Managed at ACHS",
            "Major Depression, Managed at ACHS w/ Initial PHQ-9 >= 10 within the past year"
        };
            
        private Dictionary<String, List<String>> metricNames = new Dictionary<string, List<string>>()
        {
           {"Diabetes", new List<String>()
                {   
                "Month", 
                "# of patients with DM",
                "# of patients w/ HgbA1c >=8 (High Risk)",
                "%of patients w/ 1 HgbA1c",
                "Average HgbA1c",
                "% of patients w/HgbA1c <7",
                "% of patients w/ 2 HgbA1c's",
                "% of patients with SCP",
                "% of patients 55+ y.o. on ACEI/ARB",
                "% of patients 40+ y.o. on Statin",
                "% of patients w/ BP <130/80",
                "% of patients w/ DRE within 12 months",
                "% of patients w/ DRE within 18 months",
                "% of High Risk patients w/ DRE within 12 months",
                "% of High Risk patients w/ DRE within 18 months",
                "% of patients with annual foot exam",
                "% of patients w/ LDL <100"
                }
           },
           {"Depression", new List<String>()
                {
                "Month",
                "#of patients w/ depression",
                "# of patients w/ CSD",
                "% CSD patients with 50% reduction in PHQ score",
                "% w/ documented PHQ (6 months)",
                "% with documented SCP (12 months)",
                "% of CSD patients in remission",
                " ",
                 //metrics names for right columns
                "#of patients w/ CSD",
                "% w/ 5 point reduction in PHQ w/in 6 months after initial PHQ",
                "% CSD patients with documented 1-3 wk",
                "% w/ follow-up PHQ in 3-12 wks after initial PHQ",
                "% w/ follow-up PHQ in 4-8 wks after initial PHQ",
                "% with documented SCP (12 months)",
                "% of CSD patients in remission"
                }
           },
           {"Asthma", new List<String>()
                {
                "Month",
                "Number of Patients with an Asthma Diagnosis of Persistent Asthma or Unclassified",
                "% of patients w/ a classified asthma dx",
                "# of patients w/ a persistent asthma dx",
                "% of patients with persistent or unclassified asthma w/ a documented ED visit w/in last 12 months",
                "% of patients with persistent or unclassified asthma w/ a documented asthma action plan w/in last 12 months",
                "% of patients with persistent or unclassified asthma w/ an annual review",
                "% of patients w/ persistent asthma on anti-inflammatory"
                }
           },
           {"Cardiovascular", new List<String>()
                {
                "Month",
                "# of patients w/ High Risk CVD (EXCLUDES PATIENTS W/DIABETES)",
                "% of patients w/ a documented LDL w/in last 12 months",
                "% of patients w/ LDL value <100 w/in last 12 months",
                "% of patients w/ BP < 140/90 w/in last 12 months",
                "% of patients w/documented smoking status w/in last 12 months",
                "% of patients who are current smokers",
                "% of current smokers who were counseled to quit smoking",
                "% of patients w/a documented SCP w/in last 12 months"
                }
           },
           {"Preventive", new List<String>()
                {
                "Month",
                "Percent Pediatric BMI >= 85th Percentile",
                "Percent Pediatric BMI >= 85th Percentile Counseled 5-2-1-0",
                "Percent Cervical Cancer Screening",
                "Percent Breast Cancer Screening",
                "Percent Colon Cancer Screening",
                "Percent Pneumococcal Immunization",
                "Percent 50+ y.o. AD on File"
                }
           }
        };

        private Dictionary<String, List<String>> metricGoals = new Dictionary<String, List<String>>()
        {
              {"Diabetes", new List<String>()
                {   
                " ", 
                " ",
                " ",
                " ",
                "Goal: 7.0",
                "Goal: > 58%",
                "Goal: > 90%",
                "Goal: > 70%",
                "Goal: > 75%",
                "Goal: > 60%",
                "Goal: > 40%",
                "Goal: > 70%",
                "Goal: > 70%",
                "Goal: > 70%",
                "Goal: > 70%",
                "Goal: > 90%",
                "Goal: > 60%"
                }
           },
           {"Depression", new List<String>()
                {
                //sub header for left column
                "All Patient with a Diagnosis of Depression Managed at ACHS",
                //metrics names for left columns
                " ",
                " ",
                " ",
                "Goal: > 50%",
                "Goal: > 70%",
                "Goal: > 70%",
                " ",
                //subheader for right columns
                "Major Depression, Managed at ACHS w/ Initial PHQ-9 >= 10 within the past year",
                //metrics names for right columns
                " ",
                "Goal: > 50%",
                "Goal: > 70%",
                " ",
                "Goal: > 70%",
                " ",
                "Goal: > 60%"
                }
           },
           
           {"Asthma", new List<String>()
                {
                " ",
                " ",
                "Goal: > 90%",
                " ",
                " ",
                " ",
                "Goal: > 70%",
                "Goal: > 90%"
                }
           },
           {"Cardiovascular", new List<String>()
                {
                " ",
                " ",
                " ",
                "Goal: > 60%",
                " ",
                " ",
                " ",
                " ",
                " "
                }
           },
           {"Preventive", new List<String>()
                {
                " ",
                " ",
                " ",
                "Goal: 93%",
                "Goal: > 80%",
                "Goal: > 70%",
                "Goal: 90%",
                " "
                }
           }
        };
        #endregion

        public CreateExcelDoc()
        {
            createDoc();
            populateSheets();
            //workbook.SaveAs(@"Provider Dashboards.xls");
            //app.Quit();
        }

        public void createDoc()
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1); //this initializes the workbook with 1 sheet.

                for (int i = 0; i < 4; i++) //use this to add 4 more sheets.
                {
                    worksheet = (Excel.Worksheet)workbook.Worksheets.Add();
                    worksheet.Name = sheetNames[i];
                }
                workbook.Sheets[5].Name = sheetNames[4]; //seems to be the easist way to name the last sheet Preventive
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

            }
            catch (Exception e)
            {
                Console.Write("Error: " + e);
            }
            finally
            {
            }
        }

        public void populateSheets()
        {

            for (int i = 1; i <= 5; i++) //number of sheets
            {
                int row = 1;
                int col = 1;
                worksheet = (Excel.Worksheet)workbook.Sheets[i];
                for (int j = 0; j < providerNames.Count; j++)
                {
                    row++;
                    this.createHeaders(row, 1, providerNames[j], "A" + row, "Q" + row, 2, "GRAY", true, 10, "n");
                    row++;
                    //check to see if sheet is 2 or depression and act accordingly
                    int depressionSubheader = 0;
                    if (i == 2 && depressionSubheader == 0)
                    {
                        workSheet_range = worksheet.get_Range("H:H"); // this is the way to select the entire column
                        workSheet_range.Interior.Color = System.Drawing.Color.Black.ToArgb();
                        workSheet_range.ColumnWidth = 0.58;

                        for (int x = 0; x <= 1; x++)
                        {
                            switch (x)
                            {
                                case 0:
                                    this.createHeaders(row, col, depessionMetricSubHeaders[0], "A" + row, "G" + row, 2, "WHITE", true, 10, "n");
                                    col++; 
                                    depressionSubheader++;
                                    break;
                                case 1:
                                    this.createHeaders(row, col + 7, depessionMetricSubHeaders[1], "I" + row, "O" + row, 2, "WHITE", true, 10, "n");
                                    row++;
                                    col--;//sets the next row back to the right column
                                    break;
                                default:
                                    break;

                            }
                        }
 
                    }

                    for (int k = 0; k < metricNames.ElementAt(i - 1).Value.Count; k++)
                    {
                        this.addData(row, col, metricNames.ElementAt(i - 1).Value[k],  "@");
                        col++;
                    }
                    col = 1;
                    row++;
                    for (int l = 0; l < metricGoals.ElementAt(i - 1).Value.Count; l++)
                    {
                        this.addData(row, col, metricGoals.ElementAt(i - 1).Value[l], "@");

                        col++;
                    }
                    col = 1;
                }
            }
        }

        public void createHeaders(int row, int col, string htext, string cell1,
        string cell2, int mergeColumns, string color, bool font, int size, string
        fcolor)
        {
            worksheet.Cells[row, col] = htext;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Merge(mergeColumns);

            switch (color)
            {
                case "YELLOW":
                    workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    workSheet_range.Interior.Color = System.Drawing.Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.PeachPuff.ToArgb();
                    break;
                default:
                    //workSheet_range.Interior.Color = System.Drawing.Color.White.ToArgb();
                    break;
            }

            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Font.Bold = font;
            workSheet_range.ColumnWidth = size;
            if (fcolor.Equals(""))
            {
                workSheet_range.Font.Color = System.Drawing.Color.White.ToArgb();
            }
            else
            {
                workSheet_range.Font.Color = System.Drawing.Color.Black.ToArgb();
            }
        }

        public void addData(int row, int col, string data, string format)
        {
            //workSheet_range = worksheet.get_Range(row, col); // this is the way to select the entire column
           // workSheet_range.WrapText = true;
            worksheet.Cells[row, col] = data;
            workSheet_range.WrapText = true;
            //workSheet_range.Font.Bold = data;
            workSheet_range.NumberFormat = format;
            
        }
    }
}
