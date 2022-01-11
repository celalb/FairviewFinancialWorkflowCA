using FairviewFinancialCA;

using Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace FairviewFinancialWorkflowCA
{
    public class B1Starter
    {
        PrePayment prepayment;
        SAPbouiCOM.Form oForm;

        public B1Starter()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = null;
                string sConnectionString = null;
                if (Environment.GetCommandLineArgs().Length > 1)
                    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                else
                    sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                SboGuiApi = new SAPbouiCOM.SboGuiApi();
                SboGuiApi.Connect(sConnectionString);
                ProgData.B1Application = SboGuiApi.GetApplication(-1);

                ProgData.B1Company = (SAPbobsCOM.Company)ProgData.B1Application.Company.GetDICompany();
                //    ProgData.B1Application.StatusBar.SetText(string.Format("{0} loading ...", "LBAddon"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                ProgData.B1Application.AppEvent += B1Application_AppEvent;

                Logger.firm = ProgData.B1Company.CompanyDB;
                Logger.exeFolder = AppDomain.CurrentDomain.BaseDirectory;
                if (false == ProgData.B1Application.Menus.Item("2816").SubMenus.Exists("PP_001"))
                    ProgData.B1Application.Menus.Item("2816").SubMenus.Add("PP_001", "Consolidated Payments", SAPbouiCOM.BoMenuType.mt_STRING, 1);
                XmlDocument oXmlDoc = new XmlDocument();
                oXmlDoc.Load("ServiceConfig.xml");

                string srv = oXmlDoc.SelectSingleNode("/IndyDutch/Database/Server").InnerText;


                string dbname = ProgData.B1Company.CompanyDB;
                string dbuser = oXmlDoc.SelectSingleNode("/IndyDutch/Database/UID").InnerText;
                string dbpassw = B1Starter.Decrypt(oXmlDoc.SelectSingleNode("/IndyDutch/Database/PWD").InnerText, true);

                ProgData.sqlConnectionString = $"Server = {srv}; Database = {dbname}; User Id = {dbuser}; Password = {dbpassw};";
                ProgData.Forms = new Dictionary<string, IForm>();
                DbUtility dbutility = new DbUtility();
                //   dbutility.CreateTables();


                Subscribe();

                ProgData.B1Application.StatusBar.SetText(string.Format("{0} loading ...", "Pre Payment"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);


            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                Environment.Exit(0);
            }

        }
        private static string Decrypt(string cipherString, bool useHashing)
        {
            try
            {
                byte[] keyArray;
                //get the byte code of the string

                byte[] toEncryptArray = Convert.FromBase64String(cipherString);

                System.Configuration.AppSettingsReader settingsReader =
                                                    new AppSettingsReader();
                //Get your key from config file to open the lock!
                string key = "Linux123#"; //(string)settingsReader.GetValue("SecurityKey",
                                          //                   typeof(String));

                if (useHashing)
                {
                    //if hashing was used get the hash code with regards to your key
                    MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                    keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                    //release any resource held by the MD5CryptoServiceProvider

                    hashmd5.Clear();
                }
                else
                {
                    //if hashing was not implemented get the byte code of the key
                    keyArray = UTF8Encoding.UTF8.GetBytes(key);
                }

                TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
                //set the secret key for the tripleDES algorithm
                tdes.Key = keyArray;

                tdes.Mode = CipherMode.ECB;

                tdes.Padding = PaddingMode.PKCS7;

                ICryptoTransform cTransform = tdes.CreateDecryptor();
                byte[] resultArray = cTransform.TransformFinalBlock(
                                     toEncryptArray, 0, toEncryptArray.Length);
                //Release resources held by TripleDes Encryptor                
                tdes.Clear();
                //return the Clear decrypted TEXT
                return UTF8Encoding.UTF8.GetString(resultArray);
            }
            catch (Exception er)
            {

            }
            return cipherString;
        }

        public void Subscribe()
        {
            ProgData.B1Application.ItemEvent += B1Application_ItemEvent;
            ProgData.B1Application.MenuEvent += B1Application_MenuEvent;
            ProgData.B1Application.RightClickEvent += B1Application_RightClickEvent;
            ProgData.B1Application.AppEvent += B1Application_AppEvent;
        }
        private static void RemoveMenu()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = ProgData.B1Application.Menus;

            //SAPbouiCOM.MenuCreationParams oCreationPackage = null;


            try
            {
                oMenuItem = ProgData.B1Application.Menus.Item("43540");
                oMenus = oMenuItem.SubMenus;
                SAPbouiCOM.MenuItem MenuI = oMenus.Item("47700");
                oMenus.Remove(MenuI);
            }
            catch (Exception err)
            {
            }
        }
        private static void B1Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        RemoveMenu();
                        ProgData.B1Company.Disconnect();
                        System.Environment.Exit(0);

                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        RemoveMenu();
                        ProgData.B1Company.Disconnect();
                        System.Environment.Exit(0);
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        RemoveMenu();
                        ProgData.B1Company.Disconnect();
                        System.Environment.Exit(0);
                        break;
                }
            }
            catch (Exception ex)
            {

            }


        }
        private void B1Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        private void B1Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "PP_001" && pVal.BeforeAction == false)
                {
                    prepayment = new PrePayment();
                }
                else
                    if (pVal.MenuUID == "1282" || pVal.MenuUID == "1281" || pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291")
                {
                    string uId = ProgData.B1Application.Forms.ActiveForm.UniqueID;
                    if (ProgData.Forms.ContainsKey(uId))
                    {

                        IForm form = ProgData.Forms[uId];
                        form.Menu(ref pVal);
                    }
                }

            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }


        private void B1Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (ProgData.Forms.ContainsKey(FormUID))
                {
                    oForm = ProgData.B1Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    IForm form = ProgData.Forms[FormUID];

                    if (pVal.BeforeAction == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.CharPressed == 13
                         && oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        BubbleEvent = false;
                        form.EventAll(pVal);
                    }
                    else
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    {
                        try
                        {
                            ProgData.Forms.Remove(oForm.UniqueID);
                        }
                        catch (Exception e)
                        {

                        }
                    }
                    else
                        form.EventAll(pVal);
                }

            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }


    }

}
