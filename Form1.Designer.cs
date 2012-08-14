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
            this.updateDashboardsButton = new System.Windows.Forms.Button();
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
            this.metricsFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.dashboardFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.loadDashboardDialog = new System.Windows.Forms.OpenFileDialog();
            this.updateMetrics = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.MainForm.SuspendLayout();
            this.MainPage.SuspendLayout();
            this.SettingsPage.SuspendLayout();
            this.providersBox.SuspendLayout();
            this.DataLocationBox.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainForm
            // 
            this.MainForm.Controls.Add(this.MainPage);
            this.MainForm.Controls.Add(this.SettingsPage);
            this.MainForm.Location = new System.Drawing.Point(1, 2);
            this.MainForm.Name = "MainForm";
            this.MainForm.SelectedIndex = 0;
            this.MainForm.Size = new System.Drawing.Size(550, 692);
            this.MainForm.TabIndex = 0;
            // 
            // MainPage
            // 
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
            // loadDashboardDialog
            // 
            this.loadDashboardDialog.FileName = "loadDashboiardDialog";
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
            this.SettingsPage.ResumeLayout(false);
            this.providersBox.ResumeLayout(false);
            this.providersBox.PerformLayout();
            this.DataLocationBox.ResumeLayout(false);
            this.DataLocationBox.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
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
    }
}

