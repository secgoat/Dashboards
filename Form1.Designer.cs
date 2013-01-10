namespace ProviderDashboards
{
    partial class ProviderDashboards
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.MainForm = new System.Windows.Forms.TabControl();
            this.MainPage = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.emailButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.updateDashboardsButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.updateMetrics = new System.Windows.Forms.Button();
            this.createDashboardsButton = new System.Windows.Forms.Button();
            this.SettingsPage = new System.Windows.Forms.TabPage();
            this.saveSettingsButton = new System.Windows.Forms.Button();
            this.providersBox = new System.Windows.Forms.GroupBox();
            this.removeOneFromMetrics = new System.Windows.Forms.Button();
            this.removeAllFromMetrics = new System.Windows.Forms.Button();
            this.addOneFromAD = new System.Windows.Forms.Button();
            this.addAllFromAd = new System.Windows.Forms.Button();
            this.metricsList = new System.Windows.Forms.ListBox();
            this.ADList = new System.Windows.Forms.ListBox();
            this.metricsProviders = new System.Windows.Forms.Label();
            this.ProvidersAD = new System.Windows.Forms.Label();
            this.DataLocationBox = new System.Windows.Forms.GroupBox();
            this.dashboardBrowseButton = new System.Windows.Forms.Button();
            this.metricsBrowseButton = new System.Windows.Forms.Button();
            this.dashboardTextBox = new System.Windows.Forms.TextBox();
            this.dashboardLocationLabel = new System.Windows.Forms.Label();
            this.metricsLocationLabel = new System.Windows.Forms.Label();
            this.metricstextBox = new System.Windows.Forms.TextBox();
            this.matchMetrics = new System.Windows.Forms.TabPage();
            this.matchMetricsTabs = new System.Windows.Forms.TabControl();
            this.diabetesMetricsMatch = new System.Windows.Forms.TabPage();
            this.diabetesMatchedMetrics = new System.Windows.Forms.ListBox();
            this.diabetesDashMetrics = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.depressionMetricsMatch = new System.Windows.Forms.TabPage();
            this.depressionMatchedMetrics = new System.Windows.Forms.ListBox();
            this.depressionDashMetrics = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.asthmaMetricsMatch = new System.Windows.Forms.TabPage();
            this.asthmaMatchedMetrics = new System.Windows.Forms.ListBox();
            this.asthmaDashMetrics = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cardiovascularMetricsMatch = new System.Windows.Forms.TabPage();
            this.cardioMatchedMetrics = new System.Windows.Forms.ListBox();
            this.cardioDashMetrics = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.preventivemetricsMatch = new System.Windows.Forms.TabPage();
            this.preventiveMatchedMetrics = new System.Windows.Forms.ListBox();
            this.preventiveDashMetrics = new System.Windows.Forms.ListBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.metricsFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.dashboardFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.loadDashboardDialog = new System.Windows.Forms.OpenFileDialog();
            this.selectMetricFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.MainForm.SuspendLayout();
            this.MainPage.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SettingsPage.SuspendLayout();
            this.providersBox.SuspendLayout();
            this.DataLocationBox.SuspendLayout();
            this.matchMetrics.SuspendLayout();
            this.matchMetricsTabs.SuspendLayout();
            this.diabetesMetricsMatch.SuspendLayout();
            this.depressionMetricsMatch.SuspendLayout();
            this.asthmaMetricsMatch.SuspendLayout();
            this.cardiovascularMetricsMatch.SuspendLayout();
            this.preventivemetricsMatch.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainForm
            // 
            this.MainForm.Controls.Add(this.MainPage);
            this.MainForm.Controls.Add(this.SettingsPage);
            this.MainForm.Controls.Add(this.matchMetrics);
            this.MainForm.Location = new System.Drawing.Point(1, 2);
            this.MainForm.Name = "MainForm";
            this.MainForm.SelectedIndex = 0;
            this.MainForm.Size = new System.Drawing.Size(550, 692);
            this.MainForm.TabIndex = 0;
            // 
            // MainPage
            // 
            this.MainPage.Controls.Add(this.groupBox3);
            this.MainPage.Controls.Add(this.groupBox2);
            this.MainPage.Controls.Add(this.groupBox1);
            this.MainPage.Controls.Add(this.createDashboardsButton);
            this.MainPage.Location = new System.Drawing.Point(4, 22);
            this.MainPage.Name = "MainPage";
            this.MainPage.Padding = new System.Windows.Forms.Padding(3);
            this.MainPage.Size = new System.Drawing.Size(542, 666);
            this.MainPage.TabIndex = 0;
            this.MainPage.Text = "Main";
            this.MainPage.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.emailButton);
            this.groupBox3.Location = new System.Drawing.Point(7, 209);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(145, 64);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Email Dashboards";
            // 
            // emailButton
            // 
            this.emailButton.Location = new System.Drawing.Point(6, 35);
            this.emailButton.Name = "emailButton";
            this.emailButton.Size = new System.Drawing.Size(75, 23);
            this.emailButton.TabIndex = 1;
            this.emailButton.Text = "Email";
            this.emailButton.UseVisualStyleBackColor = true;
            this.emailButton.Click += new System.EventHandler(this.emailButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.updateDashboardsButton);
            this.groupBox2.Location = new System.Drawing.Point(7, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(145, 64);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Update Provider Dashboards";
            // 
            // updateDashboardsButton
            // 
            this.updateDashboardsButton.Location = new System.Drawing.Point(6, 35);
            this.updateDashboardsButton.Name = "updateDashboardsButton";
            this.updateDashboardsButton.Size = new System.Drawing.Size(75, 23);
            this.updateDashboardsButton.TabIndex = 1;
            this.updateDashboardsButton.Text = "Update Dashboards";
            this.updateDashboardsButton.UseVisualStyleBackColor = true;
            this.updateDashboardsButton.Click += new System.EventHandler(this.updateDashboardsButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.updateMetrics);
            this.groupBox1.Location = new System.Drawing.Point(7, 19);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(145, 64);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Update Excel 2003 to 2007";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // updateMetrics
            // 
            this.updateMetrics.Location = new System.Drawing.Point(6, 35);
            this.updateMetrics.Name = "updateMetrics";
            this.updateMetrics.Size = new System.Drawing.Size(75, 23);
            this.updateMetrics.TabIndex = 2;
            this.updateMetrics.Text = "Update";
            this.updateMetrics.UseVisualStyleBackColor = true;
            this.updateMetrics.Click += new System.EventHandler(this.updateMetrics_Click);
            // 
            // createDashboardsButton
            // 
            this.createDashboardsButton.Location = new System.Drawing.Point(13, 597);
            this.createDashboardsButton.Name = "createDashboardsButton";
            this.createDashboardsButton.Size = new System.Drawing.Size(190, 41);
            this.createDashboardsButton.TabIndex = 0;
            this.createDashboardsButton.Text = "Create New Dashboards";
            this.createDashboardsButton.UseVisualStyleBackColor = true;
            this.createDashboardsButton.Click += new System.EventHandler(this.createDashboards_Click);
            // 
            // SettingsPage
            // 
            this.SettingsPage.Controls.Add(this.saveSettingsButton);
            this.SettingsPage.Controls.Add(this.providersBox);
            this.SettingsPage.Controls.Add(this.DataLocationBox);
            this.SettingsPage.Location = new System.Drawing.Point(4, 22);
            this.SettingsPage.Name = "SettingsPage";
            this.SettingsPage.Padding = new System.Windows.Forms.Padding(3);
            this.SettingsPage.Size = new System.Drawing.Size(542, 666);
            this.SettingsPage.TabIndex = 1;
            this.SettingsPage.Text = "Settings";
            this.SettingsPage.UseVisualStyleBackColor = true;
            // 
            // saveSettingsButton
            // 
            this.saveSettingsButton.Location = new System.Drawing.Point(377, 598);
            this.saveSettingsButton.Name = "saveSettingsButton";
            this.saveSettingsButton.Size = new System.Drawing.Size(146, 54);
            this.saveSettingsButton.TabIndex = 2;
            this.saveSettingsButton.Text = "Save Settings";
            this.saveSettingsButton.UseVisualStyleBackColor = true;
            this.saveSettingsButton.Click += new System.EventHandler(this.saveSettingsButton_Click);
            // 
            // providersBox
            // 
            this.providersBox.Controls.Add(this.removeOneFromMetrics);
            this.providersBox.Controls.Add(this.removeAllFromMetrics);
            this.providersBox.Controls.Add(this.addOneFromAD);
            this.providersBox.Controls.Add(this.addAllFromAd);
            this.providersBox.Controls.Add(this.metricsList);
            this.providersBox.Controls.Add(this.ADList);
            this.providersBox.Controls.Add(this.metricsProviders);
            this.providersBox.Controls.Add(this.ProvidersAD);
            this.providersBox.Location = new System.Drawing.Point(7, 142);
            this.providersBox.Name = "providersBox";
            this.providersBox.Size = new System.Drawing.Size(529, 450);
            this.providersBox.TabIndex = 1;
            this.providersBox.TabStop = false;
            this.providersBox.Text = "Add / Remove Providers";
            // 
            // removeOneFromMetrics
            // 
            this.removeOneFromMetrics.Location = new System.Drawing.Point(256, 159);
            this.removeOneFromMetrics.Name = "removeOneFromMetrics";
            this.removeOneFromMetrics.Size = new System.Drawing.Size(75, 23);
            this.removeOneFromMetrics.TabIndex = 7;
            this.removeOneFromMetrics.Text = "<";
            this.removeOneFromMetrics.UseVisualStyleBackColor = true;
            this.removeOneFromMetrics.Click += new System.EventHandler(this.removeOneFromMetrics_Click);
            // 
            // removeAllFromMetrics
            // 
            this.removeAllFromMetrics.Location = new System.Drawing.Point(256, 130);
            this.removeAllFromMetrics.Name = "removeAllFromMetrics";
            this.removeAllFromMetrics.Size = new System.Drawing.Size(75, 23);
            this.removeAllFromMetrics.TabIndex = 6;
            this.removeAllFromMetrics.Text = "<<";
            this.removeAllFromMetrics.UseVisualStyleBackColor = true;
            this.removeAllFromMetrics.Click += new System.EventHandler(this.removeAllFromMetrics_Click);
            // 
            // addOneFromAD
            // 
            this.addOneFromAD.Location = new System.Drawing.Point(200, 75);
            this.addOneFromAD.Name = "addOneFromAD";
            this.addOneFromAD.Size = new System.Drawing.Size(75, 23);
            this.addOneFromAD.TabIndex = 5;
            this.addOneFromAD.Text = ">";
            this.addOneFromAD.UseVisualStyleBackColor = true;
            this.addOneFromAD.Click += new System.EventHandler(this.addOneFromAD_Click);
            // 
            // addAllFromAd
            // 
            this.addAllFromAd.Location = new System.Drawing.Point(200, 46);
            this.addAllFromAd.Name = "addAllFromAd";
            this.addAllFromAd.Size = new System.Drawing.Size(75, 23);
            this.addAllFromAd.TabIndex = 4;
            this.addAllFromAd.Text = ">>";
            this.addAllFromAd.UseVisualStyleBackColor = true;
            this.addAllFromAd.Click += new System.EventHandler(this.addAllFromAd_Click);
            // 
            // metricsList
            // 
            this.metricsList.FormattingEnabled = true;
            this.metricsList.Location = new System.Drawing.Point(349, 46);
            this.metricsList.Name = "metricsList";
            this.metricsList.Size = new System.Drawing.Size(176, 394);
            this.metricsList.TabIndex = 3;
            // 
            // ADList
            // 
            this.ADList.FormattingEnabled = true;
            this.ADList.Location = new System.Drawing.Point(9, 46);
            this.ADList.Name = "ADList";
            this.ADList.Size = new System.Drawing.Size(176, 394);
            this.ADList.TabIndex = 2;
            // 
            // metricsProviders
            // 
            this.metricsProviders.AutoSize = true;
            this.metricsProviders.Location = new System.Drawing.Point(346, 30);
            this.metricsProviders.Name = "metricsProviders";
            this.metricsProviders.Size = new System.Drawing.Size(88, 13);
            this.metricsProviders.TabIndex = 1;
            this.metricsProviders.Text = "Metrics Providers";
            // 
            // ProvidersAD
            // 
            this.ProvidersAD.AutoSize = true;
            this.ProvidersAD.Location = new System.Drawing.Point(9, 30);
            this.ProvidersAD.Name = "ProvidersAD";
            this.ProvidersAD.Size = new System.Drawing.Size(129, 13);
            this.ProvidersAD.TabIndex = 0;
            this.ProvidersAD.Text = "Active Directory Providers";
            // 
            // DataLocationBox
            // 
            this.DataLocationBox.Controls.Add(this.dashboardBrowseButton);
            this.DataLocationBox.Controls.Add(this.metricsBrowseButton);
            this.DataLocationBox.Controls.Add(this.dashboardTextBox);
            this.DataLocationBox.Controls.Add(this.dashboardLocationLabel);
            this.DataLocationBox.Controls.Add(this.metricsLocationLabel);
            this.DataLocationBox.Controls.Add(this.metricstextBox);
            this.DataLocationBox.Location = new System.Drawing.Point(7, 6);
            this.DataLocationBox.Name = "DataLocationBox";
            this.DataLocationBox.Size = new System.Drawing.Size(532, 130);
            this.DataLocationBox.TabIndex = 0;
            this.DataLocationBox.TabStop = false;
            this.DataLocationBox.Text = "Set Data Locations";
            // 
            // dashboardBrowseButton
            // 
            this.dashboardBrowseButton.Location = new System.Drawing.Point(441, 56);
            this.dashboardBrowseButton.Name = "dashboardBrowseButton";
            this.dashboardBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.dashboardBrowseButton.TabIndex = 13;
            this.dashboardBrowseButton.Text = "Browse";
            this.dashboardBrowseButton.UseVisualStyleBackColor = true;
            this.dashboardBrowseButton.Click += new System.EventHandler(this.dashboardBrowseButton_Click);
            // 
            // metricsBrowseButton
            // 
            this.metricsBrowseButton.Location = new System.Drawing.Point(441, 16);
            this.metricsBrowseButton.Name = "metricsBrowseButton";
            this.metricsBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.metricsBrowseButton.TabIndex = 12;
            this.metricsBrowseButton.Text = "Browse";
            this.metricsBrowseButton.UseVisualStyleBackColor = true;
            this.metricsBrowseButton.Click += new System.EventHandler(this.metricsBrowseButton_Click);
            // 
            // dashboardTextBox
            // 
            this.dashboardTextBox.Location = new System.Drawing.Point(124, 59);
            this.dashboardTextBox.Name = "dashboardTextBox";
            this.dashboardTextBox.Size = new System.Drawing.Size(311, 20);
            this.dashboardTextBox.TabIndex = 11;
            this.dashboardTextBox.Text = " ";
            // 
            // dashboardLocationLabel
            // 
            this.dashboardLocationLabel.AutoSize = true;
            this.dashboardLocationLabel.Location = new System.Drawing.Point(6, 66);
            this.dashboardLocationLabel.Name = "dashboardLocationLabel";
            this.dashboardLocationLabel.Size = new System.Drawing.Size(103, 13);
            this.dashboardLocationLabel.TabIndex = 10;
            this.dashboardLocationLabel.Text = "Dashboard Location";
            // 
            // metricsLocationLabel
            // 
            this.metricsLocationLabel.AutoSize = true;
            this.metricsLocationLabel.Location = new System.Drawing.Point(6, 26);
            this.metricsLocationLabel.Name = "metricsLocationLabel";
            this.metricsLocationLabel.Size = new System.Drawing.Size(85, 13);
            this.metricsLocationLabel.TabIndex = 9;
            this.metricsLocationLabel.Text = "Metrics Location";
            // 
            // metricstextBox
            // 
            this.metricstextBox.Location = new System.Drawing.Point(124, 19);
            this.metricstextBox.Name = "metricstextBox";
            this.metricstextBox.Size = new System.Drawing.Size(311, 20);
            this.metricstextBox.TabIndex = 8;
            this.metricstextBox.Text = " ";
            // 
            // matchMetrics
            // 
            this.matchMetrics.Controls.Add(this.matchMetricsTabs);
            this.matchMetrics.Location = new System.Drawing.Point(4, 22);
            this.matchMetrics.Name = "matchMetrics";
            this.matchMetrics.Padding = new System.Windows.Forms.Padding(3);
            this.matchMetrics.Size = new System.Drawing.Size(542, 666);
            this.matchMetrics.TabIndex = 2;
            this.matchMetrics.Text = "matchMetrics";
            this.matchMetrics.UseVisualStyleBackColor = true;
            // 
            // matchMetricsTabs
            // 
            this.matchMetricsTabs.Controls.Add(this.diabetesMetricsMatch);
            this.matchMetricsTabs.Controls.Add(this.depressionMetricsMatch);
            this.matchMetricsTabs.Controls.Add(this.asthmaMetricsMatch);
            this.matchMetricsTabs.Controls.Add(this.cardiovascularMetricsMatch);
            this.matchMetricsTabs.Controls.Add(this.preventivemetricsMatch);
            this.matchMetricsTabs.Location = new System.Drawing.Point(3, 3);
            this.matchMetricsTabs.Name = "matchMetricsTabs";
            this.matchMetricsTabs.SelectedIndex = 0;
            this.matchMetricsTabs.Size = new System.Drawing.Size(536, 660);
            this.matchMetricsTabs.TabIndex = 0;
            // 
            // diabetesMetricsMatch
            // 
            this.diabetesMetricsMatch.Controls.Add(this.diabetesMatchedMetrics);
            this.diabetesMetricsMatch.Controls.Add(this.diabetesDashMetrics);
            this.diabetesMetricsMatch.Controls.Add(this.label2);
            this.diabetesMetricsMatch.Controls.Add(this.label1);
            this.diabetesMetricsMatch.Location = new System.Drawing.Point(4, 22);
            this.diabetesMetricsMatch.Name = "diabetesMetricsMatch";
            this.diabetesMetricsMatch.Padding = new System.Windows.Forms.Padding(3);
            this.diabetesMetricsMatch.Size = new System.Drawing.Size(528, 634);
            this.diabetesMetricsMatch.TabIndex = 0;
            this.diabetesMetricsMatch.Text = "Diabetes";
            this.diabetesMetricsMatch.UseVisualStyleBackColor = true;
            // 
            // diabetesMatchedMetrics
            // 
            this.diabetesMatchedMetrics.FormattingEnabled = true;
            this.diabetesMatchedMetrics.Location = new System.Drawing.Point(356, 20);
            this.diabetesMatchedMetrics.Name = "diabetesMatchedMetrics";
            this.diabetesMatchedMetrics.Size = new System.Drawing.Size(166, 602);
            this.diabetesMatchedMetrics.TabIndex = 5;
            this.diabetesMatchedMetrics.SelectedIndexChanged += new System.EventHandler(this.diabetesMatchedMetrics_SelectedIndexChanged);
            // 
            // diabetesDashMetrics
            // 
            this.diabetesDashMetrics.FormattingEnabled = true;
            this.diabetesDashMetrics.Location = new System.Drawing.Point(9, 20);
            this.diabetesDashMetrics.Name = "diabetesDashMetrics";
            this.diabetesDashMetrics.Size = new System.Drawing.Size(150, 602);
            this.diabetesDashMetrics.TabIndex = 4;
            this.diabetesDashMetrics.SelectedIndexChanged += new System.EventHandler(this.diabetesDashMetrics_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(353, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Matched Metric Name";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Metrics From Dashboard";
            // 
            // depressionMetricsMatch
            // 
            this.depressionMetricsMatch.Controls.Add(this.depressionMatchedMetrics);
            this.depressionMetricsMatch.Controls.Add(this.depressionDashMetrics);
            this.depressionMetricsMatch.Controls.Add(this.label3);
            this.depressionMetricsMatch.Controls.Add(this.label4);
            this.depressionMetricsMatch.Location = new System.Drawing.Point(4, 22);
            this.depressionMetricsMatch.Name = "depressionMetricsMatch";
            this.depressionMetricsMatch.Padding = new System.Windows.Forms.Padding(3);
            this.depressionMetricsMatch.Size = new System.Drawing.Size(528, 634);
            this.depressionMetricsMatch.TabIndex = 1;
            this.depressionMetricsMatch.Text = "Depression";
            this.depressionMetricsMatch.UseVisualStyleBackColor = true;
            // 
            // depressionMatchedMetrics
            // 
            this.depressionMatchedMetrics.FormattingEnabled = true;
            this.depressionMatchedMetrics.Location = new System.Drawing.Point(356, 25);
            this.depressionMatchedMetrics.Name = "depressionMatchedMetrics";
            this.depressionMatchedMetrics.Size = new System.Drawing.Size(166, 602);
            this.depressionMatchedMetrics.TabIndex = 9;
            // 
            // depressionDashMetrics
            // 
            this.depressionDashMetrics.FormattingEnabled = true;
            this.depressionDashMetrics.Location = new System.Drawing.Point(9, 25);
            this.depressionDashMetrics.Name = "depressionDashMetrics";
            this.depressionDashMetrics.Size = new System.Drawing.Size(150, 602);
            this.depressionDashMetrics.TabIndex = 8;
            this.depressionDashMetrics.SelectedIndexChanged += new System.EventHandler(this.depressionDashMetrics_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(353, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Matched Metric Name";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(122, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Metrics From Dashboard";
            // 
            // asthmaMetricsMatch
            // 
            this.asthmaMetricsMatch.Controls.Add(this.asthmaMatchedMetrics);
            this.asthmaMetricsMatch.Controls.Add(this.asthmaDashMetrics);
            this.asthmaMetricsMatch.Controls.Add(this.label5);
            this.asthmaMetricsMatch.Controls.Add(this.label6);
            this.asthmaMetricsMatch.Location = new System.Drawing.Point(4, 22);
            this.asthmaMetricsMatch.Name = "asthmaMetricsMatch";
            this.asthmaMetricsMatch.Padding = new System.Windows.Forms.Padding(3);
            this.asthmaMetricsMatch.Size = new System.Drawing.Size(528, 634);
            this.asthmaMetricsMatch.TabIndex = 2;
            this.asthmaMetricsMatch.Text = "Asthma";
            this.asthmaMetricsMatch.UseVisualStyleBackColor = true;
            // 
            // asthmaMatchedMetrics
            // 
            this.asthmaMatchedMetrics.FormattingEnabled = true;
            this.asthmaMatchedMetrics.Location = new System.Drawing.Point(356, 25);
            this.asthmaMatchedMetrics.Name = "asthmaMatchedMetrics";
            this.asthmaMatchedMetrics.Size = new System.Drawing.Size(166, 602);
            this.asthmaMatchedMetrics.TabIndex = 9;
            // 
            // asthmaDashMetrics
            // 
            this.asthmaDashMetrics.FormattingEnabled = true;
            this.asthmaDashMetrics.Location = new System.Drawing.Point(9, 25);
            this.asthmaDashMetrics.Name = "asthmaDashMetrics";
            this.asthmaDashMetrics.Size = new System.Drawing.Size(150, 602);
            this.asthmaDashMetrics.TabIndex = 8;
            this.asthmaDashMetrics.SelectedIndexChanged += new System.EventHandler(this.asthmaDashMetrics_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(353, 7);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(112, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Matched Metric Name";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 8);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(122, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Metrics From Dashboard";
            // 
            // cardiovascularMetricsMatch
            // 
            this.cardiovascularMetricsMatch.Controls.Add(this.cardioMatchedMetrics);
            this.cardiovascularMetricsMatch.Controls.Add(this.cardioDashMetrics);
            this.cardiovascularMetricsMatch.Controls.Add(this.label7);
            this.cardiovascularMetricsMatch.Controls.Add(this.label8);
            this.cardiovascularMetricsMatch.Location = new System.Drawing.Point(4, 22);
            this.cardiovascularMetricsMatch.Name = "cardiovascularMetricsMatch";
            this.cardiovascularMetricsMatch.Padding = new System.Windows.Forms.Padding(3);
            this.cardiovascularMetricsMatch.Size = new System.Drawing.Size(528, 634);
            this.cardiovascularMetricsMatch.TabIndex = 3;
            this.cardiovascularMetricsMatch.Text = "Cardiovascular";
            this.cardiovascularMetricsMatch.UseVisualStyleBackColor = true;
            // 
            // cardioMatchedMetrics
            // 
            this.cardioMatchedMetrics.FormattingEnabled = true;
            this.cardioMatchedMetrics.Location = new System.Drawing.Point(356, 25);
            this.cardioMatchedMetrics.Name = "cardioMatchedMetrics";
            this.cardioMatchedMetrics.Size = new System.Drawing.Size(166, 602);
            this.cardioMatchedMetrics.TabIndex = 9;
            // 
            // cardioDashMetrics
            // 
            this.cardioDashMetrics.FormattingEnabled = true;
            this.cardioDashMetrics.Location = new System.Drawing.Point(9, 25);
            this.cardioDashMetrics.Name = "cardioDashMetrics";
            this.cardioDashMetrics.Size = new System.Drawing.Size(150, 602);
            this.cardioDashMetrics.TabIndex = 8;
            this.cardioDashMetrics.SelectedIndexChanged += new System.EventHandler(this.cardioDashMetrics_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(353, 7);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(112, 13);
            this.label7.TabIndex = 7;
            this.label7.Text = "Matched Metric Name";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 8);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(122, 13);
            this.label8.TabIndex = 6;
            this.label8.Text = "Metrics From Dashboard";
            // 
            // preventivemetricsMatch
            // 
            this.preventivemetricsMatch.Controls.Add(this.preventiveMatchedMetrics);
            this.preventivemetricsMatch.Controls.Add(this.preventiveDashMetrics);
            this.preventivemetricsMatch.Controls.Add(this.label9);
            this.preventivemetricsMatch.Controls.Add(this.label10);
            this.preventivemetricsMatch.Location = new System.Drawing.Point(4, 22);
            this.preventivemetricsMatch.Name = "preventivemetricsMatch";
            this.preventivemetricsMatch.Padding = new System.Windows.Forms.Padding(3);
            this.preventivemetricsMatch.Size = new System.Drawing.Size(528, 634);
            this.preventivemetricsMatch.TabIndex = 4;
            this.preventivemetricsMatch.Text = "Preventive";
            this.preventivemetricsMatch.UseVisualStyleBackColor = true;
            // 
            // preventiveMatchedMetrics
            // 
            this.preventiveMatchedMetrics.FormattingEnabled = true;
            this.preventiveMatchedMetrics.Location = new System.Drawing.Point(356, 25);
            this.preventiveMatchedMetrics.Name = "preventiveMatchedMetrics";
            this.preventiveMatchedMetrics.Size = new System.Drawing.Size(166, 602);
            this.preventiveMatchedMetrics.TabIndex = 9;
            // 
            // preventiveDashMetrics
            // 
            this.preventiveDashMetrics.FormattingEnabled = true;
            this.preventiveDashMetrics.Location = new System.Drawing.Point(9, 25);
            this.preventiveDashMetrics.Name = "preventiveDashMetrics";
            this.preventiveDashMetrics.Size = new System.Drawing.Size(150, 602);
            this.preventiveDashMetrics.TabIndex = 8;
            this.preventiveDashMetrics.SelectedIndexChanged += new System.EventHandler(this.preventiveDashMetrics_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(353, 7);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(112, 13);
            this.label9.TabIndex = 7;
            this.label9.Text = "Matched Metric Name";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 8);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(122, 13);
            this.label10.TabIndex = 6;
            this.label10.Text = "Metrics From Dashboard";
            // 
            // loadDashboardDialog
            // 
            this.loadDashboardDialog.FileName = "loadDashboiardDialog";
            // 
            // selectMetricFileDialog
            // 
            this.selectMetricFileDialog.FileName = "selectMetricFileDialog";
            // 
            // ProviderDashboards
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(549, 688);
            this.Controls.Add(this.MainForm);
            this.Name = "ProviderDashboards";
            this.Text = "Provider Dashboards";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.MainForm.ResumeLayout(false);
            this.MainPage.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.SettingsPage.ResumeLayout(false);
            this.providersBox.ResumeLayout(false);
            this.providersBox.PerformLayout();
            this.DataLocationBox.ResumeLayout(false);
            this.DataLocationBox.PerformLayout();
            this.matchMetrics.ResumeLayout(false);
            this.matchMetricsTabs.ResumeLayout(false);
            this.diabetesMetricsMatch.ResumeLayout(false);
            this.diabetesMetricsMatch.PerformLayout();
            this.depressionMetricsMatch.ResumeLayout(false);
            this.depressionMetricsMatch.PerformLayout();
            this.asthmaMetricsMatch.ResumeLayout(false);
            this.asthmaMetricsMatch.PerformLayout();
            this.cardiovascularMetricsMatch.ResumeLayout(false);
            this.cardiovascularMetricsMatch.PerformLayout();
            this.preventivemetricsMatch.ResumeLayout(false);
            this.preventivemetricsMatch.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl MainForm;
        private System.Windows.Forms.TabPage MainPage;
        private System.Windows.Forms.TabPage SettingsPage;
        private System.Windows.Forms.Button createDashboardsButton;
        private System.Windows.Forms.GroupBox DataLocationBox;
        private System.Windows.Forms.TextBox dashboardTextBox;
        private System.Windows.Forms.Label dashboardLocationLabel;
        private System.Windows.Forms.Label metricsLocationLabel;
        private System.Windows.Forms.TextBox metricstextBox;
        private System.Windows.Forms.Button saveSettingsButton;
        private System.Windows.Forms.GroupBox providersBox;
        private System.Windows.Forms.Button updateDashboardsButton;
        private System.Windows.Forms.Button dashboardBrowseButton;
        private System.Windows.Forms.Button metricsBrowseButton;
        private System.Windows.Forms.FolderBrowserDialog metricsFolderBrowserDialog;
        private System.Windows.Forms.FolderBrowserDialog dashboardFolderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog loadDashboardDialog;
        private System.Windows.Forms.Button removeOneFromMetrics;
        private System.Windows.Forms.Button removeAllFromMetrics;
        private System.Windows.Forms.Button addOneFromAD;
        private System.Windows.Forms.Button addAllFromAd;
        private System.Windows.Forms.ListBox metricsList;
        private System.Windows.Forms.ListBox ADList;
        private System.Windows.Forms.Label metricsProviders;
        private System.Windows.Forms.Label ProvidersAD;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button updateMetrics;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button emailButton;
        private System.Windows.Forms.TabPage matchMetrics;
        private System.Windows.Forms.TabControl matchMetricsTabs;
        private System.Windows.Forms.TabPage diabetesMetricsMatch;
        private System.Windows.Forms.TabPage depressionMetricsMatch;
        private System.Windows.Forms.TabPage asthmaMetricsMatch;
        private System.Windows.Forms.TabPage cardiovascularMetricsMatch;
        private System.Windows.Forms.TabPage preventivemetricsMatch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox diabetesMatchedMetrics;
        private System.Windows.Forms.ListBox diabetesDashMetrics;
        private System.Windows.Forms.ListBox depressionMatchedMetrics;
        private System.Windows.Forms.ListBox depressionDashMetrics;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox asthmaMatchedMetrics;
        private System.Windows.Forms.ListBox asthmaDashMetrics;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox cardioMatchedMetrics;
        private System.Windows.Forms.ListBox cardioDashMetrics;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ListBox preventiveMatchedMetrics;
        private System.Windows.Forms.ListBox preventiveDashMetrics;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.OpenFileDialog selectMetricFileDialog;
    }
}

