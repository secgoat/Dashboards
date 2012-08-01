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
            readProvidersFromAD();
            //update  = new UpdateDashboards();
            update = new DashboardUpdate(providers);
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

        private void readProvidersFromAD()
        {
            PrincipalContext domainName = new PrincipalContext(ContextType.Domain, "CAMPUS");
            GroupPrincipal group = GroupPrincipal.FindByIdentity(domainName, "Provider");
            PrincipalSearchResult<Principal> members = group.GetMembers();
            foreach (Principal provider in members)
            {
                ADList.Items.Add(provider);
            }


        }
       
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
           /* Thread thread = new Thread(DashboardUpdate);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join(); */
            /* move this top some where else to change the excel files automagically
             * UpdateMetricsToXLSX change = new UpdateMetricsToXLSX();
             *   change.Convert(dataLocations[0]);
             */ 
            
            update._openExcel(dataLocations[1], dataLocations[0]);
            
        }

        private void DashboardUpdate()
        {
            update._openExcel(dataLocations[1], dataLocations[0]);
        }
    }
}
