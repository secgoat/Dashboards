using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ProviderDashboards
{
    public partial class ProviderDashboards : Form
    {
        List<String> dataLocations = new List<String>(2); //0 = metrics, 1 = dashboard
        List<String> providers = new List<String>();
        XMLSettings settings = new XMLSettings();
        //UpdateDashboards update;
        DashboardUpdate update;
        List<String> _metricsNamesFromDashboard = new List<string>();

        String settingsFileName = "Settings.xml";

        public ProviderDashboards()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists(settingsFileName))
            {
                settings.ReadConfigFile(settingsFileName, ref dataLocations, ref providers);
                //set the data locations etc, if the existed in the config file
                metricstextBox.Text = dataLocations[0];
                dashboardTextBox.Text = dataLocations[1];
               
                foreach (String provider in providers)
                {
                    metricsList.Items.Add(provider);
                }
            }
            //open the dashboards and read out the metrics names so that we can display them in the metrics match tabs
            // this will allow the user to open the excel file and match the field in which they want to populate the dshaboard position
            // this will help the program match the file name and field name so it is not hardcoded by me anymore.
            ReadProvidersFromAd();
            //update  = new UpdateDashboards();
            update = new DashboardUpdate(providers);
            GetMetricsNamesFromDashboard();
            ADList.SelectionMode = SelectionMode.MultiExtended;
            metricsList.SelectionMode = SelectionMode.MultiExtended;
        }

        private void createDashboards_Click(object sender, EventArgs e)
        {
            CreateExcelDoc excell_app = new CreateExcelDoc();
        }

        private void metricsBrowseButton_Click(object sender, EventArgs e)
        {
           if(metricsFolderBrowserDialog.ShowDialog() == DialogResult.OK)
               this.metricstextBox.Text = metricsFolderBrowserDialog.SelectedPath;
        }

        private void dashboardBrowseButton_Click(object sender, EventArgs e)
        {
            if (loadDashboardDialog.ShowDialog() == DialogResult.OK)
                this.dashboardTextBox.Text = loadDashboardDialog.FileName;
        }

        private void saveSettingsButton_Click(object sender, EventArgs e)
        {
            dataLocations.Clear();
            providers.Clear();
            dataLocations.Add(metricstextBox.Text);
            dataLocations.Add(dashboardTextBox.Text);
            foreach (var provider in metricsList.Items)
            {
                providers.Add(provider.ToString());
            }
            settings.WriteConfigFile(dataLocations, providers);
        }


        //MY functions to facilitate populating forms etc.

        private void ReadProvidersFromAd()
        {
            PrincipalContext domainName = new PrincipalContext(ContextType.Domain, "CAMPUS");
            GroupPrincipal group = GroupPrincipal.FindByIdentity(domainName, "Provider");
            PrincipalSearchResult<Principal> members = group.GetMembers();
            foreach (Principal provider in members)
            {
                ADList.Items.Add(provider);
            }


        }

        private void GetMetricsNamesFromDashboard()
        {
            //TODO: right now i am only trying to fill the diabetes dash metric list, need to add the logic to do all.
            //roll the the different tabs in the dashboard 0 -5 correspond to the
            // location of each tab: 0:diabetes, 1:Depressiojn, 2:Asthma, 3:Cardio, 4:Preventive
            for (int i = 1; i <= 5; i++) 
            {
                _metricsNamesFromDashboard = update.GetMetricsNames(dataLocations[1], i);
                foreach (string metric in _metricsNamesFromDashboard)
                {
                    switch (i)
                    {
                        case 1:
                            diabetesDashMetrics.Items.Add(metric);
                            break;
                        case 2:
                            depressionDashMetrics.Items.Add(metric);
                            break;
                        case 3:
                            asthmaDashMetrics.Items.Add(metric);
                            break;
                        case 4:
                            cardioDashMetrics.Items.Add(metric);
                            break;
                        case 5:
                            preventiveDashMetrics.Items.Add(metric);
                            break;

                    }
                    
                }
            }
        }


        //back to form programming
        //buttons for Active directory list
        private void addAllFromAd_Click(object sender, EventArgs e)
        {
            foreach (var provider in ADList.Items)
            {
                metricsList.Items.Add(provider);
            }
        }
        private void addOneFromAD_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ADList.SelectedItems.Count; i++)
            {
                metricsList.Items.Add(ADList.SelectedItems[i]);
            }
        }

        /* 
         * buttons for metrics list
         * 
         */
        private void removeAllFromMetrics_Click(object sender, EventArgs e)
        {
            metricsList.Items.Clear();
        }

        private void removeOneFromMetrics_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < metricsList.SelectedItems.Count; i++)
            {
                metricsList.Items.Remove(metricsList.SelectedItems[i]);
                --i;
            }
        }

        private void updateDashboardsButton_Click(object sender, EventArgs e)
        {
            update._openExcel(dataLocations[1], dataLocations[0]);
        }

        private void DashboardUpdate()
        {
            update._openExcel(dataLocations[1], dataLocations[0]);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void updateMetrics_Click(object sender, EventArgs e)
        {
            UpdateMetricsToXLSX change = new UpdateMetricsToXLSX();
            change.Convert(dataLocations[0]);
        }

        private void emailButton_Click(object sender, EventArgs e)
        {

        }

        

       
    }
}
