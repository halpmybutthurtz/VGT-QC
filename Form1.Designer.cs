namespace Discovery_Exporter
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
            this.refreshButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.listView = new System.Windows.Forms.ListView();
            this.Device = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.IP = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.logButton = new System.Windows.Forms.Button();
            this.serializeButton = new System.Windows.Forms.Button();
            this.resetButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // refreshButton
            // 
            resources.ApplyResources(this.refreshButton, "refreshButton");
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.UseVisualStyleBackColor = true;
            this.refreshButton.Click += new System.EventHandler(this.refreshButton_Click);
            // 
            // exportButton
            // 
            resources.ApplyResources(this.exportButton, "exportButton");
            this.exportButton.Name = "exportButton";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // listView
            // 
            resources.ApplyResources(this.listView, "listView");
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Device,
            this.IP});
            this.listView.HideSelection = false;
            this.listView.Name = "listView";
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            // 
            // Device
            // 
            resources.ApplyResources(this.Device, "Device");
            // 
            // IP
            // 
            resources.ApplyResources(this.IP, "IP");
            // 
            // logButton
            // 
            resources.ApplyResources(this.logButton, "logButton");
            this.logButton.Name = "logButton";
            this.logButton.UseVisualStyleBackColor = true;
            this.logButton.Click += new System.EventHandler(this.logButton_Click);
            // 
            // serializeButton
            // 
            resources.ApplyResources(this.serializeButton, "serializeButton");
            this.serializeButton.Name = "serializeButton";
            this.serializeButton.UseVisualStyleBackColor = true;
            this.serializeButton.Click += new System.EventHandler(this.serializeButton_Click);
            // 
            // resetButton
            // 
            resources.ApplyResources(this.resetButton, "resetButton");
            this.resetButton.Name = "resetButton";
            this.resetButton.UseVisualStyleBackColor = true;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.serializeButton);
            this.Controls.Add(this.logButton);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.refreshButton);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button refreshButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.ColumnHeader Device;
        private System.Windows.Forms.ColumnHeader IP;
        private System.Windows.Forms.Button logButton;
        private System.Windows.Forms.Button serializeButton;
        private System.Windows.Forms.Button resetButton;
    }
}

