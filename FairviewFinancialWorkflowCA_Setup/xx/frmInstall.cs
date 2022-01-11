using Microsoft.VisualBasic;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace AddOnInstaller
{
    public partial class frmInstall : Form
    {
        byte iStartInstall = 0;
        bool bError = false;
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
        public string ReadPath()
        {
            string readPathReturn = null;
            string sAns = null;
            string sErr = "";

            try
            {
                sAns = RegValue(RegistryHive.LocalMachine, "SOFTWARE", "FairviewFinancialWorkflowCA", ref sErr);
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
        public string RegValue(RegistryHive Hive, string Key, string ValueName, [System.Runtime.InteropServices.Optional] ref string ErrInfo)
        {

            RegistryKey objParent = null;
            RegistryKey objSubkey = null;
            string sAns = null;
            switch (Hive)
            {
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


            try
            {
                objSubkey = objParent.OpenSubKey(Key);
                // if can't be found, object is not initialized
                if (!(objSubkey == null))
                {
                    sAns = System.Convert.ToString((objSubkey.GetValue(ValueName)));
                }

            }
            catch (Exception ex)
            {

                ErrInfo = ex.Message;
            }
            finally
            {

                // if no error but value is empty, populate errinfo
                if (ErrInfo == "" & sAns == "")
                {
                    ErrInfo = "No value found for requested registry key";
                }
            }
            return sAns;
        }


        //  This Function writes values to the registry
        public bool WriteToRegistry(RegistryHive ParentKeyHive, string SubKeyName, string ValueName, object Value)
        {

            RegistryKey objSubKey = null;
            //string sException = null; 
            RegistryKey objParentKey = null;
            //bool bAns = false; 

            try
            {

                switch (ParentKeyHive)
                {
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
                objSubKey = objParentKey.OpenSubKey(SubKeyName, true);
                // create if doesn't exist
                if (objSubKey == null)
                {
                    objSubKey = objParentKey.CreateSubKey(SubKeyName);
                }


                objSubKey.SetValue(ValueName, Value);
                //bAns = true; 
            }
            catch (Exception ex)
            {
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
                    this.lblInfo.Text = sFileName + " Error 5 " +
                                        " Close All Business One Clients ";
                    this.btClose.Visible = false;
                    this.progressBar1.Visible = false;
                }
                // Don't continue running until the file is copied...
            }

            this.lblSymbol.Text = "@";
            this.lblSymbol.ForeColor = Color.FromArgb(0, 0, 128);
            this.lblInfo.Text = "Installing FairviewFinancialWorkflowCA ...";

            Application.DoEvents();
        }


        //  This function extracts the given add-on into the path specified
        private void ExtractFile_Resource(string sPath, string sFileName, string sExtention)
        {
            try
            {
                System.IO.FileStream AddonExeFile = null;
                System.Reflection.Assembly thisExe = null;

                thisExe = System.Reflection.Assembly.GetExecutingAssembly();

                if (thisExe == null)
                {
                    lblInfo.Text = "GetExecutingAssembly null";

                }
                object sTargetPath = sPath + @"\" + sFileName + "." + sExtention;
                object sSourcePath = sPath + @"\" + sFileName + ".tmp";

                System.IO.Stream file = null;

                file = thisExe.GetManifestResourceStream("LSoft_Setup." + sFileName + "." + sExtention);

                //  Create a tmp file first, after file is extracted change to exe
                if (System.IO.File.Exists(System.Convert.ToString(sSourcePath)))
                {

                    System.IO.File.Delete(System.Convert.ToString(sSourcePath));
                }

                AddonExeFile = System.IO.File.Create(System.Convert.ToString(sSourcePath));

                if (file == null)
                {
                    lblInfo.Text = "Manifest file null";
                    MessageBox.Show("LSoft_Setup." + sFileName + "." + sExtention + "manifest null");
                }
                byte[] buffer = null;
                buffer = new byte[file.Length];

                file.Read(buffer, 0, System.Convert.ToInt32(file.Length));

                AddonExeFile.Write(buffer, 0, System.Convert.ToInt32(file.Length));

                AddonExeFile.Close();

                if (System.IO.File.Exists(System.Convert.ToString(sTargetPath)))
                {

                    System.IO.File.Delete(System.Convert.ToString(sTargetPath));
                }
                //  Change file extension to exe

                System.IO.File.Move(System.Convert.ToString(sSourcePath), System.Convert.ToString(sTargetPath));


            }
            catch (Exception ex)
            {
                this.lblSymbol.Text = "r";
                this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                this.lblInfo.Text = "[-4] " + ex.Message + " {" + sFileName + "}";
                MessageBox.Show("[-4] " + ex.Message + " {" + sFileName + "}");
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
                            this.lblInfo.Text = sFileName + " The file cannot be deleted because it is in use.\r\n" +
                                                "Close all Business One Clients ";
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
            this.lblInfo.Text = "Uninstalling FairviewFinancialWorkflowCA ...";
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
        private void UnInstall()
        {
            string path = null;
            path = ReadPath(); //  Reads the addon path from the registry

            // Use .exe path if registry fails
            if (path.Length == 0)
            {
                path = @"C:\Program Files\SAP\SAP Business One\AddOns\FairviewFinancialWorkflowCA\FairviewFinancialWorkflowCA\";
                bError = false;
            }

            if (bError == false && path != "")
            {
                try
                {
                    //  Delete the addon EXE file
                    DeleteFile(path + @"\", "FairviewFinancialWorkflowCA.exe");
                    DeleteFile(path + @"\", "PrePAyment.xml");
                    DeleteFile(path + @"\", "ServiceConfig.xml");
                    DeleteFile(path + @"\", "FairviewFinancialWorkflowCA.pbd");
                    DeleteFile(path + @"\", "FairviewFinancialWorkflowCA.exe.config");
                }
                catch
                {
                    //MessageBox.Show("ERROR UNINSTALLING" );
                    this.lblSymbol.Text = "r";
                    this.lblSymbol.ForeColor = Color.FromArgb(192, 0, 0);
                    this.lblInfo.Text = "[-2] Error uninstalling";
                    this.btClose.Visible = true;
                    this.progressBar1.Visible = false;
                    bError = true;
                }
            }
            else
            {
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
            }
            else
            {
                EndUninstall(path, false);
            }
        }


        //  This procedure copies the addon exe file to the installation folder        
        private void Install()
        {
            try
            {
                Environment.CurrentDirectory = strDll; //  For Dll function calls will work

                if (chkDefaultFolder.Checked == false)
                { //  Change the installation folder
                    string strTemp = txtDest.Text;
                    SetAddOnFolder(strTemp);
                    strDest = txtDest.Text;
                }

                if (!((System.IO.Directory.Exists(strDest))))
                {

                    System.IO.Directory.CreateDirectory(strDest); //  Create installation folder
                }

                FileWatcher.Path = strDest;
                FileWatcher.EnableRaisingEvents = true;

                //  Extract add-on to installation folder
                ExtractFile(strDest, "FairviewFinancialWorkflowCA", "exe");
                ExtractFile(strDest, "PrePAyment","xml");
                ExtractFile(strDest, "ServiceConfig","xml");
                ExtractFile(strDest, "FairviewFinancialWorkflowCA.exe", "config");
                ExtractFile(strDest, "FairviewFinancialWorkflowCA", "pbd");

                if (chkRestart.Checked)
                {
                    RestartNeeded(); //  Inform SBO the restart is needed
                }
                lblInfo.Text = "End Install";

                EndInstallEx(strDest, true); //  Inform SBO the installation ended
                // Write installation Folder to registry
                bool bAns = false;
                lblInfo.Text = " Write registery ";

                // WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", "path", "c:\folder")
                bAns = WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", "FairviewFinancialWorkflowCA", strDest);

            }
            catch (Exception ex)
            {
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

        public frmInstall()
        {
            InitializeComponent();
        }






        private void frmInstall_Load(object sender, EventArgs e)
        {
            try
            {
                // Dim strAppPath As String

                //  The command line parameters, seperated by '|' will be broken to this array
                string[] strCmdLineElements = new string[3];

                string strCmdLine = null; //  The whole command line

                int NumOfParams = 0; // The number of parameters in the command line (should be 2)


                NumOfParams = Environment.GetCommandLineArgs().Length;

                if (NumOfParams == 2)
                {
                    strCmdLine = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
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
                else
                {
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
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }
        private void cmdInstall_Click(object sender, EventArgs e)
        {
            Install();
        }
        public void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message + ex.StackTrace);
            //     Interaction.MsgBox(ex.Message + Constants.vbNewLine + "Source:" + ex.StackTrace, MsgBoxStyle.Information, "Addon Installer");
        }

        private void FileWatcher_Renamed(object sender, System.IO.RenamedEventArgs e)
        {
            bFileCreated = true;
            FileWatcher.EnableRaisingEvents = false;
        }

        private void chkDefaultFolder_CheckedChanged(object sender, EventArgs e)
        {
            txtDest.Enabled = !((chkDefaultFolder.Checked));
        }

    }
}
