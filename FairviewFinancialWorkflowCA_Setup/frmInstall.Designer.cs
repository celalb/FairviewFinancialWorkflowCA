
namespace AddOnInstaller
{
    partial class frmInstall
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
            this.chkDefaultFolder = new System.Windows.Forms.CheckBox();
            this.chkRestart = new System.Windows.Forms.CheckBox();
            this.txtDest = new System.Windows.Forms.TextBox();
            this.Label2 = new System.Windows.Forms.Label();
            this.cmdInstall = new System.Windows.Forms.Button();
            this.Label1 = new System.Windows.Forms.Label();
            this.FileWatcher = new System.IO.FileSystemWatcher();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btClose = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.lblSymbol = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.FileWatcher)).BeginInit();
            this.SuspendLayout();
            // 
            // chkDefaultFolder
            // 
            this.chkDefaultFolder.Checked = true;
            this.chkDefaultFolder.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDefaultFolder.Enabled = false;
            this.chkDefaultFolder.Location = new System.Drawing.Point(37, 108);
            this.chkDefaultFolder.Name = "chkDefaultFolder";
            this.chkDefaultFolder.Size = new System.Drawing.Size(160, 24);
            this.chkDefaultFolder.TabIndex = 22;
            this.chkDefaultFolder.Text = "Set the place recommended by SBO";
            this.chkDefaultFolder.CheckedChanged += new System.EventHandler(this.chkDefaultFolder_CheckedChanged_1);
            // 
            // chkRestart
            // 
            this.chkRestart.Location = new System.Drawing.Point(37, 131);
            this.chkRestart.Name = "chkRestart";
            this.chkRestart.Size = new System.Drawing.Size(104, 24);
            this.chkRestart.TabIndex = 21;
            this.chkRestart.Text = "Start over";
            this.chkRestart.CheckedChanged += new System.EventHandler(this.chkRestart_CheckedChanged);
            // 
            // txtDest
            // 
            this.txtDest.Enabled = false;
            this.txtDest.Location = new System.Drawing.Point(37, 67);
            this.txtDest.Name = "txtDest";
            this.txtDest.Size = new System.Drawing.Size(472, 20);
            this.txtDest.TabIndex = 20;
            this.txtDest.Visible = false;
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(37, 43);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(256, 23);
            this.Label2.TabIndex = 19;
            this.Label2.Text = "location to install the application";
            this.Label2.Visible = false;
            // 
            // cmdInstall
            // 
            this.cmdInstall.Location = new System.Drawing.Point(314, 137);
            this.cmdInstall.Name = "cmdInstall";
            this.cmdInstall.Size = new System.Drawing.Size(96, 23);
            this.cmdInstall.TabIndex = 18;
            this.cmdInstall.Text = "Install Add-on";
            this.cmdInstall.Visible = false;
            // 
            // Label1
            // 
            this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.Label1.Location = new System.Drawing.Point(33, 9);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(416, 24);
            this.Label1.TabIndex = 17;
            this.Label1.Text = "LSoft AddOn Install";
            // 
            // FileWatcher
            // 
            this.FileWatcher.EnableRaisingEvents = true;
            this.FileWatcher.SynchronizingObject = this;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(37, 90);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(478, 12);
            this.progressBar1.TabIndex = 26;
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // btClose
            // 
            this.btClose.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btClose.Location = new System.Drawing.Point(420, 137);
            this.btClose.Name = "btClose";
            this.btClose.Size = new System.Drawing.Size(89, 26);
            this.btClose.TabIndex = 25;
            this.btClose.Text = "OK";
            this.btClose.UseVisualStyleBackColor = true;
            // 
            // lblInfo
            // 
            this.lblInfo.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.Location = new System.Drawing.Point(34, 163);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(432, 29);
            this.lblInfo.TabIndex = 24;
            this.lblInfo.Text = "Process :";
            this.lblInfo.Click += new System.EventHandler(this.lblInfo_Click);
            // 
            // lblSymbol
            // 
            this.lblSymbol.AutoSize = true;
            this.lblSymbol.Font = new System.Drawing.Font("Webdings", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.lblSymbol.ForeColor = System.Drawing.Color.Maroon;
            this.lblSymbol.Location = new System.Drawing.Point(9, 158);
            this.lblSymbol.Name = "lblSymbol";
            this.lblSymbol.Size = new System.Drawing.Size(31, 24);
            this.lblSymbol.TabIndex = 23;
            this.lblSymbol.Text = "Â";
            // 
            // frmInstall
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(625, 229);
            this.Controls.Add(this.chkDefaultFolder);
            this.Controls.Add(this.chkRestart);
            this.Controls.Add(this.txtDest);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.cmdInstall);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btClose);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.lblSymbol);
            this.Name = "frmInstall";
            this.Text = "frmInstall";
            ((System.ComponentModel.ISupportInitialize)(this.FileWatcher)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.CheckBox chkDefaultFolder;
        internal System.Windows.Forms.CheckBox chkRestart;
        internal System.Windows.Forms.TextBox txtDest;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.Button cmdInstall;
        internal System.Windows.Forms.Label Label1;
        internal System.IO.FileSystemWatcher FileWatcher;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btClose;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Label lblSymbol;
    }
}