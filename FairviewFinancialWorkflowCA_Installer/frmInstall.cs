//  SAP MANAGE DI API 2006 SDK Sample
//****************************************************************************
//
//  File:      frmInstall.cs
//
//  Copyright (c) SAP MANAGE
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
//****************************************************************************
// This sample creates an add-on installer for SBO.
// An installation for SBO should be build in a spesific way.
// 1) It should be able to accept a command line parameter from SBO.
//    This parameter is a string built from 2 strings devided by "|".
//    The first string is the path recommended by SBO for installation folder.
//    The second string is the location of "AddOnInstallAPI.dll".
//    For example, a command line parameter that looks like this:
//    "C:\MyAddon|C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
//    Means that the recommended installation folder for this addon is "C:\MyAddon"
//    and the location of "AddOnInstallAPI.dll" is - 
//                 "C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
// 2) When the installation is complete the installer must call the function 
//    "EndInstall" from "AddOnInstallAPI.dll" to inform SBO the installation is complete.
//    This dll contains 3 functions that can be used during the installation.
//    The functions are: 
//         1) EndInstall - Signals SBO that the installation is complete.
//         2) SetAddOnFolder - Use it if you want to change the installation folder.
//         3) RestartNeeded - Use it if your installation requires a restart, it will cause
//            the SBO application to close itself after the installation is complete.
//    All 3 functions return a 32 bit integer. There are 2 possible values for this integer.
//    0 - Success, 1 - Failure.
// 3) The installer must be one executable file.
// 4) After your installer is ready you need to create an add-on registration file.
//    In order to create it you have a utility - "Add-On Registration Data Creator"
//    you can find it in -
//       "..\SAP Manage\SAP Business One SDK\Tools\AddOnRegDataGen\AddOnRegDataGen.exe".
//    This utility creates a file with the extention 'ard', you will be asked to 
//    point to this file when you register your addon.

using System; 
using System.Runtime.InteropServices; 
using Microsoft.Win32; 


using Microsoft.VisualBasic;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
namespace AddOnInstaller {
    public class frmInstall : System.Windows.Forms.Form {

        byte iStartInstall = 0;
        bool bError = false;

        #region ' Windows Form Designer generated code ' 
        
        public frmInstall() { 
            
            
            // This call is required by the Windows Form Designer.
            InitializeComponent(); 
            
            // Add any initialization after the InitializeComponent() call
            
        } 
        
        // Form overrides dispose to clean up the component list.
        protected override void Dispose( bool disposing ) { 
            if ( disposing ) { 
                if ( !( ( components == null ) ) ) { 
                    components.Dispose(); 
                } 
            } 
            base.Dispose( disposing ); 
        } 
        
        
        // Required by the Windows Form Designer
        private System.ComponentModel.IContainer components; 
        
        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        internal System.Windows.Forms.Label lblHeadLine; 
        internal System.Windows.Forms.Label Label1; 
        internal System.Windows.Forms.Label Label2; 
        internal System.Windows.Forms.TextBox txtDest; 
        internal System.Windows.Forms.CheckBox chkRestart; 
        internal System.Windows.Forms.CheckBox chkDefaultFolder; 
        internal System.Windows.Forms.Button cmdInstall;
        private Label lblInfo;
        private PictureBox pictureBox1;
        private Panel panel1;
        private Button btClose;
        private Label lblSymbol;
        private ProgressBar progressBar1; 
        internal System.IO.FileSystemWatcher FileWatcher; 
        [ System.Diagnostics.DebuggerStepThrough() ]
        private void InitializeComponent() { 
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmInstall));
            this.cmdInstall = new System.Windows.Forms.Button();
            this.lblHeadLine = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.txtDest = new System.Windows.Forms.TextBox();
            this.chkRestart = new System.Windows.Forms.CheckBox();
            this.chkDefaultFolder = new System.Windows.Forms.CheckBox();
            this.FileWatcher = new System.IO.FileSystemWatcher();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblInfo = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btClose = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblSymbol = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.FileWatcher)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdInstall
            // 
            this.cmdInstall.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdInstall.Location = new System.Drawing.Point(391, 182);
            this.cmdInstall.Name = "cmdInstall";
            this.cmdInstall.Size = new System.Drawing.Size(120, 23);
            this.cmdInstall.TabIndex = 1;
            this.cmdInstall.Text = "Install";
            this.cmdInstall.Visible = false;
            this.cmdInstall.Click += new System.EventHandler(this.cmdInstall_Click);
            // 
            // lblHeadLine
            // 
            this.lblHeadLine.Enabled = false;
            this.lblHeadLine.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeadLine.Location = new System.Drawing.Point(12, 339);
            this.lblHeadLine.Name = "lblHeadLine";
            this.lblHeadLine.Size = new System.Drawing.Size(416, 24);
            this.lblHeadLine.TabIndex = 2;
            this.lblHeadLine.Text = "This Installer is a sample for Sap Business One. ";
            this.lblHeadLine.Visible = false;
            // 
            // Label1
            // 
            this.Label1.Enabled = false;
            this.Label1.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.Location = new System.Drawing.Point(12, 363);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(416, 24);
            this.Label1.TabIndex = 3;
            this.Label1.Text = "It will install a \" & sAddonName & \" add-on";
            this.Label1.Visible = false;
            // 
            // Label2
            // 
            this.Label2.Enabled = false;
            this.Label2.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label2.Location = new System.Drawing.Point(20, 411);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(256, 23);
            this.Label2.TabIndex = 4;
            this.Label2.Text = "Installation Folder recieved from SBO application";
            this.Label2.Visible = false;
            // 
            // txtDest
            // 
            this.txtDest.Enabled = false;
            this.txtDest.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDest.Location = new System.Drawing.Point(20, 435);
            this.txtDest.Name = "txtDest";
            this.txtDest.Size = new System.Drawing.Size(472, 21);
            this.txtDest.TabIndex = 5;
            this.txtDest.Visible = false;
            // 
            // chkRestart
            // 
            this.chkRestart.Enabled = false;
            this.chkRestart.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkRestart.Location = new System.Drawing.Point(20, 499);
            this.chkRestart.Name = "chkRestart";
            this.chkRestart.Size = new System.Drawing.Size(209, 24);
            this.chkRestart.TabIndex = 6;
            this.chkRestart.Text = "Ask for a restart";
            this.chkRestart.Visible = false;
            // 
            // chkDefaultFolder
            // 
            this.chkDefaultFolder.Checked = true;
            this.chkDefaultFolder.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDefaultFolder.Enabled = false;
            this.chkDefaultFolder.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDefaultFolder.Location = new System.Drawing.Point(20, 467);
            this.chkDefaultFolder.Name = "chkDefaultFolder";
            this.chkDefaultFolder.Size = new System.Drawing.Size(324, 24);
            this.chkDefaultFolder.TabIndex = 7;
            this.chkDefaultFolder.Text = "Use path supplied by SBO";
            this.chkDefaultFolder.Visible = false;
            this.chkDefaultFolder.CheckedChanged += new System.EventHandler(this.chkDefaultFolder_CheckedChanged);
            // 
            // FileWatcher
            // 
            this.FileWatcher.EnableRaisingEvents = true;
            this.FileWatcher.SynchronizingObject = this;
            this.FileWatcher.Renamed += new System.IO.RenamedEventHandler(this.FileWatcher_Renamed);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
            this.pictureBox1.Location = new System.Drawing.Point(11, 11);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(478, 94);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // lblInfo
            // 
            this.lblInfo.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.Location = new System.Drawing.Point(32, 123);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(432, 29);
            this.lblInfo.TabIndex = 9;
            this.lblInfo.Text = "FairviewFinancialWorkflowCA.exe is locked and can\'t be deleted.\r\nClose other Business One Clients or terminate the process.";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btClose);
            this.panel1.Controls.Add(this.lblInfo);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.lblSymbol);
            this.panel1.Location = new System.Drawing.Point(1, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(499, 160);
            this.panel1.TabIndex = 10;
            // 
            // btClose
            // 
            this.btClose.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btClose.Location = new System.Drawing.Point(398, 118);
            this.btClose.Name = "btClose";
            this.btClose.Size = new System.Drawing.Size(89, 26);
            this.btClose.TabIndex = 0;
            this.btClose.Text = "OK";
            this.btClose.UseVisualStyleBackColor = true;
            this.btClose.Click += new System.EventHandler(this.btClose_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 106);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(478, 12);
            this.progressBar1.TabIndex = 10;
            // 
            // lblSymbol
            // 
            this.lblSymbol.AutoSize = true;
            this.lblSymbol.Font = new System.Drawing.Font("Webdings", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.lblSymbol.ForeColor = System.Drawing.Color.Maroon;
            this.lblSymbol.Location = new System.Drawing.Point(7, 118);
            this.lblSymbol.Name = "lblSymbol";
            this.lblSymbol.Size = new System.Drawing.Size(31, 24);
            this.lblSymbol.TabIndex = 1;
            this.lblSymbol.Text = "Â";
            // 
            // frmInstall
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(500, 160);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.chkDefaultFolder);
            this.Controls.Add(this.chkRestart);
            this.Controls.Add(this.txtDest);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.lblHeadLine);
            this.Controls.Add(this.cmdInstall);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmInstall";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "IndyDutch FairviewFinancialWorkflowCA - Setup";
            this.TopMost = true;
            this.Activated += new System.EventHandler(this.frmInstall_Activated);
            this.Load += new System.EventHandler(this.frmInstall_Load);
            ((System.ComponentModel.ISupportInitialize)(this.FileWatcher)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        } 
        
        
        #endregion 
        
        #region 'Data members' 
        private string sAddonName = "FairviewFinancialWorkflowCA"; 
        private string strDll; //  The path of "AddOnInstallAPI.dll"
        private string strDest; //  Installation target path
        private bool bFileCreated; //  True if the file was created
        #endregion 
        
        #region 'Declarations' 
        
        /*
        //EndInstall - Signals SBO that the installation is complete.
        [DllImport("AddOnInstallAPI.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int EndInstallEx(string str, bool b);
        //EndUnInstall - Signals SBO that the uninstallation is complete.
        [DllImport("AddOnInstallAPI.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int EndUninstall(string str, bool b);
        //SetAddOnFolder - Use it if you want to change the installation folder.
        [DllImport("AddOnInstallAPI.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int SetAddOnFolder(string srrPath);
        //RestartNeeded - Use it if your installation requires a restart, it will cause
        //the SBO application to close itself after the installation is complete.
        [DllImport("AddOnInstallAPI.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int RestartNeeded();
        //the SBO application to close itself after the installation is complete.
        [DllImport("AddOnInstallAPI.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int B1Info(ref string lpBuffer, int length);
         */

        //EndInstall - Signals SBO that the installation is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern int EndInstallEx(string str, bool b);

        //EndUnInstall - Signals SBO that the uninstallation is complete.
        [DllImport("AddOnInstallAPI.dll")]
        public static extern int EndUninstall(string str, bool b);

		//  Declaring the functions inside "AddOnInstallAPI.dll"
        [DllImport("AddOnInstallAPI.dll")]
		static extern Int32 EndInstall();

		// EndInstall - Signals SBO that the installation is complete.
        // SetAddOnFolder - Use it if you want to change the installation folder.
        [DllImport("AddOnInstallAPI.dll")]
		static extern Int32 SetAddOnFolder(string srrPath);

		// RestartNeeded - Use it if your installation requires a restart, it will cause
        // the SBO application to close itself after the installation is complete.
        [DllImport("AddOnInstallAPI.dll")]
		static extern Int32 RestartNeeded();
        
		#endregion 
        
        #region 'Methods' 
        
        //  Read the addon path from the registry
        public string ReadPath() { 
            string readPathReturn = null;
            string sAns = null; 
            string sErr = "";

            try
            {
                sAns = RegValue(RegistryHive.LocalMachine, "SOFTWARE", "FairFinCA", ref sErr);
                readPathReturn = sAns;
                if (!((sAns != "")))
                {
                    this.lblSymbol.Text = "r";
                    this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                    this.lblInfo.Text = "[-3] " + sErr;
                    this.btClose.Visible = true;
                    this.progressBar1.Visible = false;
                    bError = true;
                    //MessageBox.Show( "This error occurred: " + sErr ); 
                }
            }
            catch
            {
            }
            return readPathReturn;
        } 
        
        
        //  This Function reads values to the registry
        public string RegValue( RegistryHive Hive, string Key, string ValueName, [System.Runtime.InteropServices.Optional] ref string ErrInfo  ) { 
            
            RegistryKey objParent = null; 
            RegistryKey objSubkey = null; 
            string sAns = null; 
            switch ( Hive ) {
                case RegistryHive.ClassesRoot:
                    objParent = Registry.ClassesRoot; 
                    break;
                case RegistryHive.CurrentConfig:
                    objParent = Registry.CurrentConfig; 
                    break;
                case RegistryHive.CurrentUser:
                    objParent = Registry.CurrentUser; 
                    break;
                case RegistryHive.DynData:
                    objParent = Registry.DynData; 
                    break;
                case RegistryHive.LocalMachine:
                    objParent = Registry.LocalMachine; 
                    break;
                case RegistryHive.PerformanceData:
                    objParent = Registry.PerformanceData; 
                    break;
                case RegistryHive.Users:
                    objParent = Registry.Users; 
                    
                    break;
            }
            
            
            try { 
                objSubkey = objParent.OpenSubKey( Key ); 
                // if can't be found, object is not initialized
                if ( !( objSubkey == null ) ) { 
                    sAns = System.Convert.ToString( ( objSubkey.GetValue( ValueName ) ) ); 
                } 
                
            } 
            catch ( Exception ex ) { 
                
                ErrInfo = ex.Message; 
            } 
            finally { 
                
                // if no error but value is empty, populate errinfo
                if ( ErrInfo == "" & sAns == "" ) { 
                    ErrInfo = "No value found for requested registry key"; 
                } 
            } 
            return sAns; 
        } 
        
        
        //  This Function writes values to the registry
        public bool WriteToRegistry( RegistryHive ParentKeyHive, string SubKeyName, string ValueName, object Value ) { 
            
            RegistryKey objSubKey = null; 
            //string sException = null; 
            RegistryKey objParentKey = null; 
            //bool bAns = false; 
            
            try { 
                
                switch ( ParentKeyHive ) {
                    case RegistryHive.ClassesRoot:
                        objParentKey = Registry.ClassesRoot; 
                        break;
                    case RegistryHive.CurrentConfig:
                        objParentKey = Registry.CurrentConfig; 
                        break;
                    case RegistryHive.CurrentUser:
                        objParentKey = Registry.CurrentUser; 
                        break;
                    case RegistryHive.DynData:
                        objParentKey = Registry.DynData; 
                        break;
                    case RegistryHive.LocalMachine:
                        objParentKey = Registry.LocalMachine; 
                        break;
                    case RegistryHive.PerformanceData:
                        objParentKey = Registry.PerformanceData; 
                        break;
                    case RegistryHive.Users:
                        objParentKey = Registry.Users; 
                        break;
                }
                
                
                // Open 
                objSubKey = objParentKey.OpenSubKey( SubKeyName, true ); 
                // create if doesn't exist
                if ( objSubKey == null ) { 
                    objSubKey = objParentKey.CreateSubKey( SubKeyName ); 
                } 
                
                
                objSubKey.SetValue( ValueName, Value ); 
                //bAns = true; 
            } 
            catch ( Exception ex ) { 
                //bAns = false; 
                
            } 
            
            return true; 
            
        }

        private void ExtractFile(string sPath, string sFileName, string sExtention)
        {
            ExtractFile_Resource(sPath, sFileName, sExtention);

            int iRetry = 0;

            while (bFileCreated == false)
            {
                Application.DoEvents();

                System.Threading.Thread.Sleep(1000);

                Application.DoEvents();

                iRetry++;

                if (iRetry == 5)
                {
                    this.lblSymbol.Text = "Â";
                    this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                    this.lblInfo.Text = sFileName + " is locked and can't be copied.\r\n" +
                                        "Close other Business One Clients or terminate the process.";
                    this.btClose.Visible = false;
                    this.progressBar1.Visible = false;
                }
                // Don't continue running until the file is copied...
            }

            this.lblSymbol.Text = "@";
            this.lblSymbol.ForeColor = Color.FromArgb(0, 0, 128);
            this.lblInfo.Text = "Installing FairViewFinancialWorkflow ...";

            Application.DoEvents();
        }

        
        //  This function extracts the given add-on into the path specified
        private void ExtractFile_Resource( string sPath, string sFileName, string sExtention ) { 
            try { 
                System.IO.FileStream AddonExeFile = null; 
                System.Reflection.Assembly thisExe = null; 
                thisExe = System.Reflection.Assembly.GetExecutingAssembly(); 
                object sTargetPath = sPath + @"\" + sFileName + "." + sExtention; 
                object sSourcePath = sPath + @"\" + sFileName + ".tmp"; 
                
                System.IO.Stream file = null;

                file = thisExe.GetManifestResourceStream("FairviewFinancialWorkflowCA_Setup." + sFileName + "." + sExtention); 
                
                //  Create a tmp file first, after file is extracted change to exe
                if ( System.IO.File.Exists( System.Convert.ToString( sSourcePath ) ) ) { 
                    System.IO.File.Delete( System.Convert.ToString( sSourcePath ) ); 
                } 
                AddonExeFile = System.IO.File.Create( System.Convert.ToString( sSourcePath ) ); 
                
                byte[] buffer = null; 
                buffer = new byte[ file.Length ]; 
                
                file.Read( buffer, 0, System.Convert.ToInt32( file.Length ) ); 
                AddonExeFile.Write( buffer, 0, System.Convert.ToInt32( file.Length ) ); 
                AddonExeFile.Close(); 
                
                if ( System.IO.File.Exists( System.Convert.ToString( sTargetPath ) ) ) { 
                    System.IO.File.Delete( System.Convert.ToString( sTargetPath ) ); 
                } 
                //  Change file extension to exe
                System.IO.File.Move( System.Convert.ToString( sSourcePath ), System.Convert.ToString( sTargetPath ) ); 
                
            } 
            catch ( Exception ex ) {
                this.lblSymbol.Text = "r";
                this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                this.lblInfo.Text = "[-4] " + ex.Message + " {" + sFileName + "}";
                this.btClose.Visible = true;
                this.progressBar1.Visible = false;
                bError = true;
                //Interaction.MsgBox( ex.Message, (Microsoft.VisualBasic.MsgBoxStyle)(0), null ); 
            } 
        } 
        

        private void DeleteFile(string sPath, string sFileName)
        {
            bool bDeleted = false;

            int iRetry = 0;

            if (System.IO.File.Exists(sPath + sFileName))
            {
                while (bDeleted == false)
                {
                    Application.DoEvents();

                    try
                    {
                        System.IO.File.Delete(sPath + sFileName);
                        bDeleted = true;
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(1000);

                        Application.DoEvents();

                        iRetry++;

                        if (iRetry == 5)
                        {
                            this.lblSymbol.Text = "Â";
                            this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                            this.lblInfo.Text = sFileName + " is locked and can't be deleted.\r\n" +
                                                "Close other Business One Clients or terminate the process.";
                            this.btClose.Visible = false;
                            this.progressBar1.Visible = false;
                        }

                        Application.DoEvents();

                        System.Threading.Thread.Sleep(1000);

                        Application.DoEvents();
                    }
                }
                //MessageBox.Show( path + @"\B1Assistant.exe was deleted" ); 
            }
            else
            {
                //MessageBox.Show( path + @"\B1Assistant.exe was not found" ); 
            }

            this.lblSymbol.Text = "i";
            this.lblSymbol.ForeColor = Color.FromArgb(0, 0, 128);
            this.lblInfo.Text = "Uninstalling FairViewFinancialWorkflow ...";
            this.btClose.Visible = false;
            this.progressBar1.Visible = false;

            Application.DoEvents();
        }

        private void DeleteFolder(string sFolder)
        {
            try
            {
                System.IO.Directory.Delete(sFolder, true);
            }
            catch
            {
            }
        }
        
        //  This procedure delets the addon files
        private void UnInstall() {            
            string path = null; 
            path = ReadPath(); //  Reads the addon path from the registry
            
            // Use .exe path if registry fails
            if(path.Length == 0)
            {
                path = @"C:\Program Files (x86)\SAP\SAP Business One\AddOns\IndyDutch\FairViewFinancialWorkflow\";
                bError = false;
            }

            if (bError == false && path != "" ) { 
                try { 
                    //  Delete the addon EXE file
                    DeleteFile(path + @"\", "FairviewFinancialWorkflowCA.exe");
                    try
                    {
                        DeleteFile(path + @"\", "FairviewFinancialCA.dll");
                        DeleteFile(path + @"\", "Shared.dll");
                    } catch (Exception e)
                    {

                    }
                    // DeleteFile(path + @"\", "PrePAyment.xml");
                    // DeleteFile(path + @"\", "ServiceConfig.xml");

                } 
                catch  { 
                    //MessageBox.Show("ERROR UNINSTALLING" );
                    this.lblSymbol.Text = "r";
                    this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                    this.lblInfo.Text = "[-2] Error uninstalling";
                    this.btClose.Visible = true;
                    this.progressBar1.Visible = false;
                    bError = true;
                } 
            } 
            else { 
                //MessageBox.Show( "Path not found" );

                //this.lblSymbol.Text = "r";
                //this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                //this.lblInfo.Text = "[-1] Uninstall finished (limited)";

                this.lblSymbol.Text = "r";
                this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                this.lblInfo.Text = "[-1] Path not found in the registry " + path;
                this.btClose.Visible = true;
                this.progressBar1.Visible = false;
                bError = true;
            }            

            if (bError == false)
            {
                progressBar1.Visible = true;
                progressBar1.Minimum = 1;
                progressBar1.Maximum = 4;

                progressBar1.Value = 1;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 2;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 3;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 4;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                EndUninstall(path, true);

                //  Terminate the application
                GC.Collect();
                System.Environment.Exit(0);                
            } else
            {
                EndUninstall(path, false);
            }
        } 
        
        
        //  This procedure copies the addon exe file to the installation folder        
        private void Install() { 
            try { 
                Environment.CurrentDirectory = strDll; //  For Dll function calls will work
                
                if ( chkDefaultFolder.Checked == false ) { //  Change the installation folder
                    string strTemp = txtDest.Text; 
                    SetAddOnFolder( strTemp ); 
                    strDest = txtDest.Text; 
                } 
                
                if ( !( ( System.IO.Directory.Exists( strDest ) ) ) ) { 
                    System.IO.Directory.CreateDirectory( strDest ); //  Create installation folder
                } 
                
                FileWatcher.Path = strDest; 
                FileWatcher.EnableRaisingEvents = true; 
                
                //  Extract add-on to installation folder
                ExtractFile(strDest, "FairviewFinancialWorkflowCA", "exe");
                try
                {
                    ExtractFile(strDest, "FairviewFinancialCA", "dll");
                    ExtractFile(strDest, "Shared", "dll");
                } catch (Exception e)
                {

                }
                //  ExtractFile(strDest, "PrePAyment", "xml");
                //ExtractFile(strDest, "ServiceConfig", "xml");


                if ( chkRestart.Checked ) { 
                    RestartNeeded(); //  Inform SBO the restart is needed
                } 
                EndInstallEx(strDest, true); //  Inform SBO the installation ended
                // Write installation Folder to registry
                bool bAns = false; 
                
                // WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", "path", "c:\folder")
                bAns = WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", "FairFinCA", strDest);
                
            } 
            catch ( Exception ex ) { 
                //Interaction.MsgBox( ex.Message, MsgBoxStyle.Information, "Addon Installer" ); 
                this.lblSymbol.Text = "r";
                this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                this.lblInfo.Text = "[-5] " + ex.Message;
                this.btClose.Visible = true;
                this.progressBar1.Visible = false;
                bError = true;
            }

            if (bError == false)
            {
                progressBar1.Visible = true;
                progressBar1.Minimum = 1;
                progressBar1.Maximum = 4;

                progressBar1.Value = 1;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 2;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 3;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                progressBar1.Value = 4;
                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                //  Terminate the application
                GC.Collect();

                //MessageBox.Show( "Finished Installing", "Installation ended", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                System.Windows.Forms.Application.Exit(); //  Exit the installer
            }
            else
            {
                EndInstallEx(strDest, false);
            }
        } 
        
        
        #endregion 
        
        #region 'Events' 
        private void frmInstall_Load( System.Object sender, System.EventArgs e ) { 
            try {
                // Dim strAppPath As String

                //  The command line parameters, seperated by '|' will be broken to this array
                string[] strCmdLineElements = new string[ 3 ]; 
                
                string strCmdLine = null; //  The whole command line
                
                int NumOfParams = 0; // The number of parameters in the command line (should be 2)
                
                
                NumOfParams = Environment.GetCommandLineArgs().Length; 
                
                if ( NumOfParams == 2 ) { 
                    strCmdLine = System.Convert.ToString( Environment.GetCommandLineArgs().GetValue( 1 ) );
                    if (strCmdLine.ToUpper() == "/U")
                    {                        
                        iStartInstall = 2;
                        this.lblSymbol.Text = "i";
                        this.lblSymbol.ForeColor = Color.FromArgb(0, 0, 128);
                        this.lblInfo.Text = "Uninstalling FairviewFinancialWorkflowCA ...";
                        this.btClose.Visible = false;
                        this.progressBar1.Visible = false;
                    }
                    else
                    {
                        strCmdLineElements = strCmdLine.Split(char.Parse("|"));

                        //  Get Install destination Folder
                        strDest = System.Convert.ToString(strCmdLineElements.GetValue(0));
                        txtDest.Text = strDest;

                        //  Get the "AddOnInstallAPI.dll" path
                        strDll = System.Convert.ToString(strCmdLineElements.GetValue(1));
                        strDll = strDll.Remove((strDll.Length - 19), 19); //  Only the path is needed

                        iStartInstall = 1;
                        this.lblSymbol.Text = "@";
                        this.lblSymbol.ForeColor = Color.FromArgb(0, 0, 128);
                        this.lblInfo.Text = "Installing FairviewFinancialWorkflowCA ...";
                        this.btClose.Visible = false;
                        this.progressBar1.Visible = false;
                    }
                } 
                else {
                    iStartInstall = 0;
                    this.lblSymbol.Text = "r";
                    this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                    this.lblInfo.Text = "This installer must be run from Sap Business One";
                    this.btClose.Visible = true;
                    this.progressBar1.Visible = false;
                    //MessageBox.Show( "This installer must be run from Sap Business One", "Incorrect installation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation ); 
                    //System.Windows.Forms.Application.Exit(); 
                }

                this.Visible = true;
                this.Update();
                Application.DoEvents();

                if (iStartInstall == 1)
                {
                    iStartInstall = 99;
                    Install();
                }

                if (iStartInstall == 2)
                {
                    iStartInstall = 99;
                    UnInstall();
                }      
            } 
            catch ( Exception ex ) { 
                ShowError( ex ); 
            } 
        } 
        
        
        private void cmdInstall_Click( System.Object sender, System.EventArgs e ) { 
            //Install(); 
        } 
        
        private void chkDefaultFolder_CheckedChanged( System.Object sender, System.EventArgs e ) { 
            txtDest.Enabled = !( ( chkDefaultFolder.Checked ) ); 
        } 
        
        
        //  This event happens when the addon exe file is renamed to exe extention
        private void FileWatcher_Renamed( object sender, System.IO.RenamedEventArgs e ) { 
            bFileCreated = true; 
            FileWatcher.EnableRaisingEvents = false; 
        } 
        
        
        public void ShowError( Exception ex ) { 
            Interaction.MsgBox( ex.Message + Constants.vbNewLine + "Source:" + ex.StackTrace, MsgBoxStyle.Information, "Add-on Setup" ); 
        } 
        
        #endregion 
        
        [STAThread]
        public static void Main() { Application.Run( new frmInstall() ); }

        private void frmInstall_Activated(object sender, EventArgs e)
        {
                  
        }

        private void btClose_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

    } 
    
    
    
} 
