namespace ExcelToAzure
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.topPanel = new System.Windows.Forms.Panel();
            this.btnData = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnLogin = new System.Windows.Forms.Button();
            this.Dashboard = new System.Windows.Forms.Panel();
            this.topPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // topPanel
            // 
            this.topPanel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.topPanel.Controls.Add(this.btnData);
            this.topPanel.Controls.Add(this.btnImport);
            this.topPanel.Controls.Add(this.btnLogin);
            this.topPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.topPanel.Location = new System.Drawing.Point(0, 0);
            this.topPanel.Margin = new System.Windows.Forms.Padding(0);
            this.topPanel.MaximumSize = new System.Drawing.Size(0, 100);
            this.topPanel.MinimumSize = new System.Drawing.Size(0, 100);
            this.topPanel.Name = "topPanel";
            this.topPanel.Size = new System.Drawing.Size(1002, 100);
            this.topPanel.TabIndex = 1;
            // 
            // btnData
            // 
            this.btnData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.btnData.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnData.BackgroundImage")));
            this.btnData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnData.FlatAppearance.BorderSize = 0;
            this.btnData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnData.Location = new System.Drawing.Point(236, 18);
            this.btnData.Margin = new System.Windows.Forms.Padding(0);
            this.btnData.MaximumSize = new System.Drawing.Size(64, 64);
            this.btnData.MinimumSize = new System.Drawing.Size(64, 64);
            this.btnData.Name = "btnData";
            this.btnData.Size = new System.Drawing.Size(64, 64);
            this.btnData.TabIndex = 2;
            this.btnData.UseVisualStyleBackColor = true;
            this.btnData.Click += new System.EventHandler(this.btnData_Click);
            // 
            // btnImport
            // 
            this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.btnImport.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnImport.BackgroundImage")));
            this.btnImport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnImport.FlatAppearance.BorderSize = 0;
            this.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImport.Location = new System.Drawing.Point(136, 18);
            this.btnImport.Margin = new System.Windows.Forms.Padding(0);
            this.btnImport.MaximumSize = new System.Drawing.Size(64, 64);
            this.btnImport.MinimumSize = new System.Drawing.Size(64, 64);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(64, 64);
            this.btnImport.TabIndex = 1;
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnLogin
            // 
            this.btnLogin.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLogin.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnLogin.BackgroundImage")));
            this.btnLogin.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnLogin.FlatAppearance.BorderSize = 0;
            this.btnLogin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLogin.Location = new System.Drawing.Point(36, 18);
            this.btnLogin.Margin = new System.Windows.Forms.Padding(0);
            this.btnLogin.MaximumSize = new System.Drawing.Size(64, 64);
            this.btnLogin.MinimumSize = new System.Drawing.Size(64, 64);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(64, 64);
            this.btnLogin.TabIndex = 0;
            this.btnLogin.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnLogin.UseVisualStyleBackColor = false;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // Dashboard
            // 
            this.Dashboard.BackColor = System.Drawing.Color.White;
            this.Dashboard.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Dashboard.Location = new System.Drawing.Point(0, 100);
            this.Dashboard.Name = "Dashboard";
            this.Dashboard.Size = new System.Drawing.Size(1002, 612);
            this.Dashboard.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1002, 712);
            this.Controls.Add(this.Dashboard);
            this.Controls.Add(this.topPanel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1024, 768);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Data Viewer";
            this.topPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel topPanel;
        private System.Windows.Forms.Button btnData;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.Panel Dashboard;
    }
}

