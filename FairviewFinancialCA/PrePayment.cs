using SAPbobsCOM;
using SAPbouiCOM;
using Shared;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialCA
{
     public class PrePayment : Shared.IForm
    {
        protected SAPbouiCOM.Form oForm = null;
        SAPbouiCOM.Button savebtn, loadbt;
        SAPbouiCOM.Matrix ogrpMatrix, oMatrix;
        DataTable Matrixdt, grpDt;
        Recordset oRec;
        int crow;
        Plist plist;
        bool Linked;
        int HDocEntry;
        UserDataSource UD_8, UD_6, UD_5, UD_10, UD_11, UD_9, UD_DocNum,UD_13;

        public Dictionary<JesKey, List<JE>> JES = new Dictionary<JesKey, List<JE>>();

        List<Payment> allpayments = new List<Payment>();
        List<Grouppay> Grouppays = new List<Grouppay>();
        int[] cellcolor;

        private void firstprepare()
        {
            Matrixdt.Columns.Add("Sel", BoFieldsType.ft_AlphaNumeric, 1);
            Matrixdt.Columns.Add("CardCode", BoFieldsType.ft_AlphaNumeric, 15);
            Matrixdt.Columns.Add("DocNum", BoFieldsType.ft_Integer, 11);
            Matrixdt.Columns.Add("DocEntry", BoFieldsType.ft_Integer, 11);
            Matrixdt.Columns.Add("DocType", BoFieldsType.ft_AlphaNumeric, 5);
            Matrixdt.Columns.Add("DocDate", BoFieldsType.ft_Date);
            Matrixdt.Columns.Add("DueDate", BoFieldsType.ft_Date);
            Matrixdt.Columns.Add("DPastDue", BoFieldsType.ft_AlphaNumeric, 4);
            Matrixdt.Columns.Add("DocTotal", BoFieldsType.ft_Sum);
            Matrixdt.Columns.Add("BalDue", BoFieldsType.ft_Sum);
            Matrixdt.Columns.Add("TotalDisc", BoFieldsType.ft_Sum);
            Matrixdt.Columns.Add("TotalPay", BoFieldsType.ft_Sum);
            Matrixdt.Columns.Add("PostedDocEntry", BoFieldsType.ft_Integer);
            Matrixdt.Columns.Add("Processed", BoFieldsType.ft_AlphaNumeric, 1);
           
            Matrixdt.Columns.Add("Rownum", BoFieldsType.ft_Integer);
            Matrixdt.Columns.Add("Linenum", BoFieldsType.ft_Integer);
            oMatrix.Columns.Item("cRownum").Visible = false;
            UD_13 = oForm.DataSources.UserDataSources.Item("Ud_13");
            UD_13.ValueEx = Logger.exeFolder + @"\addfield.jpg";
            setBindCols();
            ogrpMatrix.Columns.Item("Rownum").Visible = false;
            ogrpMatrix.Columns.Item("cCodex").TitleObject.Sortable = true;
            ogrpMatrix.Columns.Item("cNamex").TitleObject.Sortable = true;

            grpDt = oForm.DataSources.DataTables.Add("GrpPay");

            grpDt.Columns.Add("gCardCode", BoFieldsType.ft_AlphaNumeric, 20);
            grpDt.Columns.Add("gCardName", BoFieldsType.ft_AlphaNumeric, 100);
            grpDt.Columns.Add("gDocNum", BoFieldsType.ft_AlphaNumeric, 20);
            grpDt.Columns.Add("gTotal", BoFieldsType.ft_Sum);
            grpDt.Columns.Add("gDiscount", BoFieldsType.ft_Sum);
            grpDt.Columns.Add("gPayment", BoFieldsType.ft_Sum);
            grpDt.Columns.Add("gPaymonAcc", BoFieldsType.ft_Sum);
            grpDt.Columns.Add("gSelD", BoFieldsType.ft_AlphaNumeric, 1);
            grpDt.Columns.Add("gSelJ", BoFieldsType.ft_AlphaNumeric,5);
          
            grpDt.Columns.Add("Rownum", BoFieldsType.ft_Integer);
            grpDt.Columns.Add("Linenum", BoFieldsType.ft_Integer);
            grpDt.Columns.Add("JMess", BoFieldsType.ft_AlphaNumeric, 50);
            grpDt.Columns.Add("JDocentry", BoFieldsType.ft_Integer);
            grpDt.Columns.Add("DPMess", BoFieldsType.ft_AlphaNumeric, 50);
            grpDt.Columns.Add("DPDocentry", BoFieldsType.ft_Integer);
            ogrpMatrix.Columns.Item("cCodex").DataBind.Bind(grpDt.UniqueID, "gCardCode");
            ogrpMatrix.Columns.Item("cNamex").DataBind.Bind(grpDt.UniqueID, "gCardName");
            ogrpMatrix.Columns.Item("cPDocNox").DataBind.Bind(grpDt.UniqueID, "gDocNum");
            ogrpMatrix.Columns.Item("cTotalx").DataBind.Bind(grpDt.UniqueID, "gTotal");
            ogrpMatrix.Columns.Item("cTotDiscx").DataBind.Bind(grpDt.UniqueID, "gDiscount");
            Column col = ogrpMatrix.Columns.Item("cTotDiscx");
            col.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
            ogrpMatrix.Columns.Item("cPaytotx").DataBind.Bind(grpDt.UniqueID, "gPayment");
            col = ogrpMatrix.Columns.Item("cPaytotx");
            col.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
            ogrpMatrix.Columns.Item("cPayaccx").DataBind.Bind(grpDt.UniqueID, "gPaymonAcc");
            col = ogrpMatrix.Columns.Item("cPayaccx");
            col.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
            ogrpMatrix.Columns.Item("Rownum").DataBind.Bind(grpDt.UniqueID, "Linenum");
            ogrpMatrix.Columns.Item("SelJ").DataBind.Bind(grpDt.UniqueID, "gSelJ");
            ogrpMatrix.Columns.Item("cPMessx").DataBind.Bind(grpDt.UniqueID, "JMess");
            ogrpMatrix.Columns.Item("SelD").DataBind.Bind(grpDt.UniqueID, "gSelD");
            ogrpMatrix.Columns.Item("DPMess").DataBind.Bind(grpDt.UniqueID, "DPMess");
        }

        public PrePayment()
        {
            try
            {
                if (false == string.IsNullOrEmpty(Logger.exeFolder + @"\PrePayment.xml"))
                {
                    string srf = string.Empty;
                    System.Xml.XmlDocument oXMLDoc;

                    string sPath = System.Environment.CurrentDirectory + "\\Forms\\";

                    oXMLDoc = new System.Xml.XmlDocument();

                    oXMLDoc.Load(Logger.exeFolder + @"\PrePayment.xml");

                    string uid = oXMLDoc.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                    uid = uid + Guid.NewGuid().ToString().Substring(0, 6);
                    oXMLDoc.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value = uid;

                    srf = oXMLDoc.InnerXml;
                    // ProgData.B1Application.LoadBatchActions(ref srf);

                    // oForm = ProgData.B1Application.Forms.Item(uid);
                    SAPbouiCOM.FormCreationParams oCP = null;
                    oCP = ((SAPbouiCOM.FormCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                    oCP.XmlData = oXMLDoc.InnerXml;//load the form with XML 


                    oForm = ProgData.B1Application.Forms.AddEx(oCP);

                    ProgData.Forms.Add(oForm.UniqueID, this);
                    oForm.AutoManaged = true;
                    oForm.Visible = true;
                    oForm.EnableMenu("1282", true);
                    oForm.EnableMenu("1288", true);
                    oForm.EnableMenu("1289", true);
                    oForm.EnableMenu("1290", true);
                    oForm.EnableMenu("1291", true);
                    ogrpMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("mtgrpMtx").Specific);
                    oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("mtpreMtx").Specific);
                    Matrixdt = oForm.DataSources.DataTables.Add("PrePay");
                    this.firstprepare();
                    oMatrix.Item.Enabled = false;
                    ogrpMatrix.Item.Enabled = false;
                    // ChooseFromList Blist =   oForm.ChooseFromLists.Item("BList");
                    // Conditions cons = Blist.GetConditions();
                    
                    Version version = System.Reflection.Assembly.GetEntryAssembly().GetName().Version;
                    //   DateTime buildDate = new DateTime(2000, 1, 1)
                    //                         .AddDays(version.Build).AddSeconds(version.Revision * 2);
                    string displayableVersion = $"({version})";
                    oForm.Title = $"Pre Payments {displayableVersion}";

                    oForm.ActiveItem = "txtCCode";

                    UD_8 = oForm.DataSources.UserDataSources.Item("UD_8");
                    UD_8.ValueEx = DateTime.Now.ToString("yyyyMMdd");
                    UD_6 = oForm.DataSources.UserDataSources.Item("UD_6");
                    UD_6.ValueEx = DateTime.Now.ToString("yyyyMMdd");
                    UD_5 = oForm.DataSources.UserDataSources.Item("UD_5");
                    UD_10 = oForm.DataSources.UserDataSources.Item("UD_10");
                    UD_11 = oForm.DataSources.UserDataSources.Item("UD_11");
                    UD_9 = oForm.DataSources.UserDataSources.Item("UD_9");
                    UD_DocNum = oForm.DataSources.UserDataSources.Item("UD_DocNum");
                    savebtn = (Button)oForm.Items.Item("1").Specific;
                    savebtn.Item.Enabled = false;
                    loadbt = (Button)oForm.Items.Item("loadbt").Specific;
                    loadbt.Item.Enabled = false;
                    oForm.Items.Item("txtDocNum").Click();
                    oForm.Items.Item("txtCCode").Enabled = false;
                    oForm.Items.Item("txtCName").Enabled = false;
                    CheckBox fchk = (CheckBox)oForm.Items.Item("fselect").Specific;
                    fchk.Checked = false;
                    oForm.Mode = BoFormMode.fm_FIND_MODE;
                }


            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

        }
        private void clickbtnlist(ItemEvent pVal)
        {
            try
            {
                plist = new Plist();
                plist.CreatenewForm(oForm.Items.Item("txtDocNum").Left - 250, oForm.Items.Item("txtDocNum").Top + 22);
                plist.ListClosed += Plist_ListClosed;
            }
            catch (Exception er)
            {

            }
        }
        private void Plist_ListClosed(object sender, DataTable e)
        {
            plist.ListClosed -= Plist_ListClosed;
            plist = null;
            if (e == null)
            {

            }
            else
            {
                string val = e.GetValue("DocNum", e.Rows.Offset).ToString();

                if (!string.IsNullOrEmpty(val))
                {
                    Find(val);
                }
            }
        }
        private void EditText_ChooseFromListAfter(ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;

            /* SAPbouiCOM.ISBOChooseFromListEventArg oCFLEvento = null;

             oCFLEvento = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;*/
            try
            {
                oForm.Select();
                oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                string sCFL_ID = null;
                sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.DataTable oDataTable = null;
                oDataTable = oCFLEvento.SelectedObjects;
                string itemCode = null;
                string itemName = "";
                SAPbouiCOM.EditText EditText = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                SAPbouiCOM.EditText EditText1 = (SAPbouiCOM.EditText)oForm.Items.Item("txtCCode").Specific;
                SAPbouiCOM.EditText EditText2 = (SAPbouiCOM.EditText)oForm.Items.Item("txtCName").Specific;
                if (pVal.ItemUID == "txtCCode")
                {
                    itemCode = System.Convert.ToString(oDataTable.GetValue("CardCode", 0));
                    itemName = System.Convert.ToString(oDataTable.GetValue("CardName", 0));
                    UD_5.ValueEx = itemCode;
                    UD_10.ValueEx = itemName;
                    loadbt.Item.Enabled = true;
                }
                else
               if (pVal.ItemUID == "txtCName")
                {
                    itemCode = System.Convert.ToString(oDataTable.GetValue("CardCode", 0));
                    itemName = System.Convert.ToString(oDataTable.GetValue("CardName", 0));
                    UD_5.ValueEx = itemCode;
                    UD_10.ValueEx = itemName;
                    loadbt.Item.Enabled = true;
                }
                else
                {
                    itemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));

                    EditText.Value = itemCode;
                }
            }
            catch (Exception e)
            {
                Logger.Log(e);

            }
        }

        public void Save()
        {
            DbUtility dbutil = new DbUtility();
            try
            {

                string cardCode = ((SAPbouiCOM.EditText)oForm.Items.Item("txtCCode").Specific).Value;
                string val = ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocDate").Specific).Value;
                DateTime docDate = DateTime.ParseExact(val, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);
                val = ((SAPbouiCOM.EditText)oForm.Items.Item("txtPosDate").Specific).Value;
                DateTime dueDate = DateTime.ParseExact(val, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);
                string docRef = ((SAPbouiCOM.EditText)oForm.Items.Item("txtRef").Specific).Value;
                int user = ProgData.B1Company.UserSignature;
                int docEntry =0;
                int docNum = 0;
                DateTime dDate = DateTime.Today;
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {
                    docNum=  GetLastDocNum();
                    UD_DocNum.ValueEx =docNum.ToString();
                    HDocEntry = dbutil.InsertHeader(cardCode, docDate, dueDate, docRef, user,docNum);
                }
                else
                if (oForm.Mode == BoFormMode.fm_UPDATE_MODE)

                    dbutil.UpdateHeader(cardCode, docDate, dueDate, docRef, HDocEntry);

                
                string cardName = "";
                string docType = "";
                string daysPastDue = "";
                Double docTotal = 0.0;
                Double balDue = 0.0;
                Double discTotal = 0.0;
                Double payTotal = 0.0;
            //    int posteddocentry = 0;
              //  int wposteddocentry = 0;
                int linenum = 0;
                string selected = "N";
                // string xCardCode = "";
                foreach (var payment in allpayments)
                {
                    
               
                    selected = "N";
                  //  Matrixdt.Rows.Offset = payment.Linenum;
                    
                    selected = payment.Sel;
           //         if (selected == "N")
             //           continue;
                    docNum = Convert.ToInt32(payment.DocNum);
                    cardCode = payment.CardCode;


                    docEntry = Convert.ToInt32(payment.DocEntry);
                    docType = payment.DocType;
                    dDate = payment.DocDate;
                   // docDate = DateTime.ParseExact(val, "yyyyMMdd",
                     //             CultureInfo.InvariantCulture);
                    
                    dueDate = payment.DueDate;
                    daysPastDue = payment.DPastDue;
                   
                    docTotal = payment.DocTotal;
                    balDue = payment.BalDue;
                    discTotal = payment.TotalDisc;
                    payTotal =payment.TotalPay;
                   
                    dbutil.InsertTrans(HDocEntry, docEntry, payment.Linenum, docNum, cardCode, cardName, docType, dDate,
                                   dueDate, daysPastDue, docTotal, balDue, discTotal, payTotal, selected,payment.PostedDocEntry>0?"Y":"N");
                

                }
                int pJentry = -1;
                string posted = "N";
                string pJmess = "";
                int pDentry = -1;
             
                string pDmess = "";
                for (int ix = 0; ix < Grouppays.Count; ix++)
                {
                    Grouppay grp = Grouppays[ix];
                    if (grp.paymentdetail.Count == 0)
                    {

                         DeleteSummary(grp,dbutil);
                        continue;
                    }
                    pJmess = "";
                    pJentry = -1;

                    string[] ret = new string[2] { "-1", "" };
                    pJmess = grp.JMess;
                    pJentry = grp.PostedJDocEntry;
                    if (grp.PostedJDocEntry < 1)
                    {
                        ret = SaveGroup(grp.CardCode, dbutil, grp);
                        try
                        {
                            pJentry = Convert.ToInt32(ret[0]);
                            pJmess = ret[1];

                        }
                        catch (Exception e)
                        {
                            pJentry = -1;
                        }

                    }
                    pDentry = grp.PostedDDocEntry;
                    pDmess = grp.DMess;
                    if (grp.PostedDDocEntry>0)
                    {
                        dbutil.UpdateDraftPayment(grp.PostedDDocEntry, grp.CardCode, grp.TotalPay, "_SYS00000001261", docDate, docRef, grp.paymentdetail);

                    } else                 
                   
                    if (grp.PostedDDocEntry < 1)
                    {
                        ret[0] = "-1";
                        ret[1] = "";
                        if (grp.SelP == "Y")
                        ret = dbutil.CreateDraftPayment(grp.CardCode, grp.TotalPay, "_SYS00000001261", docDate, docRef, grp.paymentdetail);
                        try
                        {
                            pDentry = Convert.ToInt32(ret[0]);
                        } catch (Exception e)
                        { }
                        try
                        {
                            pDmess = ret[1];
                        }
                        catch (Exception e)
                        { }
                    }
                    posted = "N";
                    if (pJentry > 0)
                        posted = "Y";
                    dbutil.InsertSumm(HDocEntry, grp.Linenum, grp.CardCode, pJentry, pJmess, grp.SelJ, grp.PayAcc,grp.SelP,pDentry,pDmess);
                    dbutil.UpdateTransPost(HDocEntry, grp.CardCode, "Y", posted);

                }
                dbutil.Close(true);
                /*Matrixdt.Rows.Clear();
                grpDt.Rows.Clear();
                Grouppays = new List<Grouppay>();
                JES = new Dictionary<JesKey, List<JE>>();

                allpayments = new List<Payment>();
                savebtn = (Button)oForm.Items.Item("1").Specific;
                savebtn.Item.Enabled = false;
                loadbt = (Button)oForm.Items.Item("loadbt").Specific;
                loadbt.Item.Enabled = false;
               
                oMatrix.Clear();
                ogrpMatrix.Clear();*/
                oForm.Mode = BoFormMode.fm_OK_MODE;
               
                oForm.Items.Item("txtCCode").Enabled = false;
                oForm.Items.Item("txtCName").Enabled = false;
                oForm.PaneLevel = 1;
                oForm.Items.Item("fldpPaym").Click();
                oForm.Items.Item("txtDocNum").Click();
            }
            catch (Exception e)
            {
                Logger.Log(e);
                dbutil.Close(false);
               ProgData.B1Application.StatusBar.SetText(e.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            }


        }
        private void DeleteSummary(Grouppay grp,DbUtility dbutil)
        {
            if (grp.PostedDDocEntry > 0)
                dbutil.DeleteDraftPayment(grp.PostedDDocEntry);

            dbutil.DeletePayAcct(HDocEntry, grp.Linenum);
            dbutil.deleteSumm(HDocEntry, grp.Linenum);


        }

        private string[] SaveGroup(string cardcode,DbUtility dbutil, Grouppay gpay)
        {
           
            string[] ret = new string[2] { "-1", "" };
            
                JesKey jkey = new JesKey();
                jkey.CardCode = cardcode;
                jkey.DocEntry = gpay.Linenum;
                List<JE> ljes = new List<JE>();
               
                if (JES.ContainsKey(jkey))
                {
                    ljes = JES[jkey];
                    if (ljes.Count > 0)
                    {
                        if (ljes[0].Posted == "A" && gpay.SelJ == "Y")
                       {
                          ret = dbutil.CreateJEntry(ljes, cardcode);
                       }
                    }
                    foreach (var item in ljes)
                    {
                        
                        dbutil.InsertPayAcct(HDocEntry, item.Docnum, item.LineNum, item.CardCode, item.Account, item.DocType, item.PostDate, item.LineMemo, item.Debit, item.Credit,ProgData.B1Company.UserSignature ,item.Posted,gpay.Linenum);
                    }
                }
             
            return ret;

        }
        private void setBindCols()
        {
            string fldname = "";
            cellcolor = new int[oMatrix.Columns.Count];

            for (int ix = 1; ix < oMatrix.Columns.Count; ix++)
            {
                try
                {
                    fldname = oMatrix.Columns.Item(ix).UniqueID;
                    fldname = fldname.Substring(1, fldname.Length - 1);

                    oMatrix.Columns.Item(ix).DataBind.Bind(Matrixdt.UniqueID, fldname);
                    if (fldname == "CardCode" || fldname == "DocNum" || fldname == "DocTotal" || fldname == "BalDue")
                    {
                        oMatrix.Columns.Item(ix).TitleObject.Sortable = true;

                    }
                    if (fldname == "DocTotal" || fldname == "BalDue" || fldname == "TotalDisc" || fldname == "TotalPay")
                    {
                        oMatrix.Columns.Item(ix).ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                    }


                }
                catch (Exception e)
                {
                    Logger.Log(new Exception(e.Message+"**"+fldname));
                }
            }
        }
        private void setField(DataTable dt, int row, string fldname, object value)
        {
            try
            {
                dt.SetValue(fldname, row, value);
            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message+"--"+fldname+"**"+row.ToString()+"**"+value.ToString()));
            }
        }
        private void GetoldDataNav()
        {
            string qry = "";
            try
            {
                qry = "Select  T0.CardCode as HCardCode,T0.DocDate as HPayDate,T0.DocDueDate as HPostDate,T0.DocEntry as HDocEntry ,"; ;
                qry += " T0.DocNum as HDocNum,T0.DocRef as HDocRef,T0.UserSign as HUserSign,T0.CreateDate as HCreateDate,T1.*";
                qry += ",T2.PostedDocEntry,PostOption,PayOnAcc ,T2.LineNum as SumLinenum,PostedMess,T2.SelDraft as SelDraft,T2.PostDPay,T2.PostDPayMessage";
                qry += " from redi_ConsBP_Header T0 INNER JOIN redi_ConsBP_Trans T1 ON T0.DocEntry = T1.HDocEntry ";
                qry += " LEFT OUTER JOIN redi_ConsBP_Summ T2 ON T2.DocEntry = T1.HDocEntry AND T2.CardCode = T1.CardCode";
                qry += " Order By T0.DocEntry ";
                oRec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(qry);
            } catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + "**" + qry));
            }
        }
        private void Find(string DocNum)
        {
            string qry = "";
            try
            {
                qry = "Select  T0.CardCode as HCardCode,T0.DocDate as HPayDate,T0.DocDueDate as HPostDate,T0.DocEntry as HDocEntry ,";
                qry += "T0.DocNum as HDocNum,T0.DocRef as HDocRef,T0.UserSign as HUserSign,T0.CreateDate as HCreateDate,T1.*";
                qry += ",T2.PostedDocEntry,PostOption,PayOnAcc ,T2.LineNum as SumLinenum,PostedMess,T2.SelDraft as SelDraft,T2.PostDPay,T2.PostDPayMessage";
                qry += " from redi_ConsBP_Header T0 INNER JOIN redi_ConsBP_Trans T1 ON T0.DocEntry = T1.HDocEntry ";
                qry += "LEFT OUTER JOIN redi_ConsBP_Summ T2 ON T2.DocEntry = T1.HDocEntry AND T2.CardCode = T1.CardCode";
                qry += " WHERE T0.DocNum = " + DocNum + " Order By T0.DocEntry ";

                SAPbobsCOM.Recordset oRecx = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecx.DoQuery(qry);
                var task = Task.Run(async () => await recpos(DocNum));
                LoadData(oRecx);
            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + "**" + qry));
            }

        }
        private async Task recpos(string DocNum)
        {

            if (oRec == null)
            {
                GetoldDataNav();

            }
            oRec.MoveFirst();
            while (!oRec.EoF)
            {
                if (oRec.Fields.Item("DocNum").Value.ToString() == DocNum)
                {
                    break;
                }

                oRec.MoveNext();
            }
            Task t = Task.Delay(500);
            await t;


        }
        private void NavRec(string NavId)
        {
            if (oRec == null)
            {
                GetoldDataNav();

            }
            oForm.Items.Item("txtCCode").Enabled = false;
            oForm.Items.Item("txtCName").Enabled = false;

            switch (NavId)
            {

                case "1281":
                    oForm.Mode = BoFormMode.fm_FIND_MODE;
                    UD_6.ValueEx = "";
                    UD_8.ValueEx ="";
                    UD_9.ValueEx = "";
                    UD_DocNum.ValueEx = "0";
                    UD_5.ValueEx = "";
                    UD_10.ValueEx = "";
                    return;
                case "1282":
                    oForm.Mode = BoFormMode.fm_ADD_MODE;
                    oForm.Items.Item("txtCCode").Enabled = true;
                    oForm.Items.Item("txtCName").Enabled = true;
                    oForm.Items.Item("txtCCode").Click();
                    return;
                case "1288":
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oRec.MoveNext();
                    if (oRec.EoF)
                        oRec.MoveFirst();
                    break;
                case "1289":
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oRec.MovePrevious();
                    if (oRec.BoF)
                    {
                        oRec.MoveLast();
                    }
                    break;
                case "1290":
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oRec.MoveFirst();
                    break;
                case "1291":
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oRec.MoveLast();
                    break;
                default:
                    return;
            }
            LoadData(oRec);
        }
        private void LoadData(SAPbobsCOM.Recordset Rec)
        {
            try
            {
                HDocEntry = Convert.ToInt32(Rec.Fields.Item("HDocEntry").Value);
                UD_5.ValueEx = Rec.Fields.Item("HCardCode").Value.ToString();
                string HCardCode = Rec.Fields.Item("HCardCode").Value.ToString();
                SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qry = $"Select \"CardName\" from \"OCRD\" Where \"CardCode\" = '{HCardCode}'";
                orec.DoQuery(qry);
                try
                {
                    UD_10.ValueEx = orec.Fields.Item(0).Value.ToString();
                }
                catch (Exception e)
                {

                }
               
                UD_6.ValueEx = Convert.ToDateTime(Rec.Fields.Item("HPostDate").Value).ToString("yyyyMMdd");
                UD_8.ValueEx = Convert.ToDateTime(Rec.Fields.Item("HPayDate").Value).ToString("yyyyMMdd");
                UD_9.ValueEx = Rec.Fields.Item("HDocRef").Value.ToString();
                UD_DocNum.ValueEx = Rec.Fields.Item("HDocNum").Value.ToString();
                oForm.Mode = BoFormMode.fm_OK_MODE;
                List<Grouppay> oldpayments = new List<Grouppay>();
                string cardcode = "";
                Grouppay paym = new Grouppay() ;
                Matrixdt.Rows.Clear();
                grpDt.Rows.Clear();
                Grouppays = new List<Grouppay>();
                JES = new Dictionary<JesKey, List<JE>>();

                allpayments = new List<Payment>();
                savebtn = (Button)oForm.Items.Item("1").Specific;
                savebtn.Item.Enabled = false;
                loadbt = (Button)oForm.Items.Item("loadbt").Specific;
                loadbt.Item.Enabled = false;

                oMatrix.Clear();
                ogrpMatrix.Clear();
               
                oForm.Mode = BoFormMode.fm_FIND_MODE;

                oForm.Items.Item("txtCCode").Enabled = false;
                oForm.Items.Item("txtCName").Enabled = false;
                oForm.PaneLevel = 1;
                oForm.Items.Item("fldpPaym").Click();
                oForm.Items.Item("txtDocNum").Click();

                while (!Rec.EoF)
                {

                    if (cardcode != Rec.Fields.Item("CardCode").Value.ToString())                   
                    {
                        if (paym.paymentdetail!=null && paym.paymentdetail.Count>0)
                            oldpayments.Add(paym);
                        paym = new Grouppay();
                        paym.paymentdetail = new List<grpPaymentdetail>();
                        paym.Linenum = Convert.ToInt32(Rec.Fields.Item("SumLineNum").Value);
                        paym.PostedJDocEntry = Convert.ToInt32(Rec.Fields.Item("PostedDocEntry").Value);
                        paym.PostedDDocEntry = Convert.ToInt32(Rec.Fields.Item("PostDPay").Value);
                        paym.SelJ = Rec.Fields.Item("PostOption").Value.ToString();
                        paym.SelP = Rec.Fields.Item("SelDraft").Value.ToString();
                        paym.DMess = Rec.Fields.Item("PostDPayMessage").Value.ToString();
                        paym.JMess = Rec.Fields.Item("PostedMess").Value.ToString();
                        paym.PayAcc = Convert.ToDouble(Rec.Fields.Item("PayOnAcc").Value, CultureInfo.InvariantCulture);
                        paym.DocEntry = Convert.ToInt32(Rec.Fields.Item("DocEntry").Value);
                        cardcode = Rec.Fields.Item("CardCode").Value.ToString();
                        paym.CardCode = cardcode;
                       
                        
                    }
                    if (cardcode == Rec.Fields.Item("CardCode").Value.ToString())
                    {
                        grpPaymentdetail detpaym = new grpPaymentdetail();
                        detpaym.DocDocEntry = Convert.ToInt32(Rec.Fields.Item("DocEntry").Value);
                        detpaym.DocNum = Rec.Fields.Item("DocNum").Value.ToString();
                        detpaym.DocType = Rec.Fields.Item("DocType").Value.ToString();
                        detpaym.BalDue = Convert.ToDouble(Rec.Fields.Item("BalDue").Value, CultureInfo.InvariantCulture);
                        detpaym.TotalDisc = Convert.ToDouble(Rec.Fields.Item("TotalDisc").Value, CultureInfo.InvariantCulture);
                        detpaym.TotalPay = Convert.ToDouble(Rec.Fields.Item("TotalPay").Value, CultureInfo.InvariantCulture);
                        paym.paymentdetail.Add(detpaym);
                    }
                    Rec.MoveNext();
                }
                if (paym.paymentdetail != null && paym.paymentdetail.Count > 0)
                    oldpayments.Add(paym);
                Bind(HCardCode, oldpayments);
               /* FillMatrixdt(Rec,oldpayments);
                
                oMatrix.LoadFromDataSourceEx();

                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                if (oMatrix.RowCount > 0 && cellcolor.Length == 0)
                {
                    for (int ix = 1; ix < oMatrix.Columns.Count; ix++)
                    {

                        cellcolor[ix - 1] = oMatrix.CommonSetting.GetCellBackColor(1, ix);
                    }
                }*/
            } catch (Exception e)
            {
                Logger.Log(e);
            }
        }
        private int GetLastDocNum()
        {
            int ret = 0;
            string qry = "Select Max(DocNum) From redi_ConsBP_Header";
            Recordset oRecSet = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecSet.DoQuery(qry);
            try
            {
                ret = Convert.ToInt32(oRecSet.Fields.Item(0).Value);
            }
            catch (Exception e)
            {
                ret = 0;
            }
            ret++;
            return ret;
        }
        public void Bind(string ConsCode,List<Grouppay> oldpayments=null)
        {
            //string qry = $"EXEC rediSP_BPOpenTrans '{ConsCode}'";
            ProgData.B1Application.StatusBar.SetText("Fetching open transactions... please wait..", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            SAPbouiCOM.StaticText msg = (StaticText)oForm.Items.Item("msg").Specific;
            /* msg.Item.BackColor = 0x000f0000;
             msg.Caption = "Fetching open transactions... please wait..";
             msg.Item.Visible = true;*/
            oForm.Freeze(true);
            string qry = "";
            if (ProgData.B1Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                qry = "Select TOP 50 \"DocNum\",\"DocEntry\",\"CardCode\",\"CardName\",\"DocType\",\"DocDate\",\"DocDueDate\",\"DaysPastDue\",\"DocTotal\",\"BalDue\",\"DiscAmt\" as \"TotalDisc\",\"BalDue\"-\"DiscAmt\" as \"TotalPay\" ";
                qry += " FROM (select t.\"DocNum\", t.\"DocEntry\", t.\"CardCode\", t.\"CardName\", 'IN' as \"DocType\", t.\"DocDate\", t.\"DocDueDate\",";
                qry += " case when cast(DAYS_BETWEEN(CURRENT_DATE, t.\"DocDueDate\") as int) < 0 then '*' else cast(DAYS_BETWEEN(CURRENT_DATE , t.\"DocDueDate\")  as varchar) end as \"DaysPastDue\",t.\"DocTotal\", ";
                qry += " t.\"DocTotal\" - t.\"PaidToDate\" as \"BalDue\", case when DAYS_BETWEEN( t.\"DocDate\", CURRENT_DATE) <= di.\"NumOfDays\"   then(t.\"DocTotal\" - t.\"PaidToDate\") * (di.\"Discount\" / 100) else 0 end as \"DiscAmt\"  from";
                qry += " \"OINV\" t  inner join \"OCRD\" c  on c.\"CardCode\" = t.\"CardCode\"    Left join \"OCTG\" pt on pt.\"GroupNum\" = c.\"GroupNum\"   ";

                qry += " left join \"OCDC\" d on pt.\"DiscCode\" = d.\"Code\"    left join \"CDC1\" di on d.\"Code\" = di.\"CdcCode\" ) X   ORDER BY \"CardCode\"  ";
            }
            else
            {
                if (String.IsNullOrEmpty(ConsCode))
                {
                    return;
                }
                string val = ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocDate").Specific).Value;
                DateTime docDate = DateTime.ParseExact(val, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);

                qry = $"EXEC rediSP_BPOpenTrans '{ConsCode}','{ docDate.ToString("yyyyMMdd")}'";
            }


            Recordset oRecSet = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecSet.DoQuery(qry);
            if (ProgData.B1Company.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB && oldpayments == null)
                UD_DocNum.ValueEx = GetLastDocNum().ToString();
                  
            FillMatrixdt(oRecSet,oldpayments);
            oMatrix.LoadFromDataSourceEx();
            setBindCols();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            if (oForm.Mode != BoFormMode.fm_ADD_MODE)
            oForm.Mode = BoFormMode.fm_OK_MODE;

            cellcolor[0] = -1;
            
            cellcolor[1] = 15724527;
            cellcolor[2] = 15724527;
            cellcolor[3] = 15724527;
            cellcolor[4] = 15724527;
            cellcolor[5] = 15724527;
            cellcolor[6] = 15724527;
            cellcolor[7] = 15724527;
            cellcolor[8] = 15724527;
            cellcolor[9] = -1;
            cellcolor[10] = -1;

            renk = true;
            var task = Task.Run(async () => await renklendir(1));
            if (Matrixdt.Rows.Count > 0)
            {
                savebtn.Item.Enabled = true;
                oMatrix.Item.Enabled = true;
            }
            oForm.Freeze(false);

            msg.Caption = ".";
            msg.Item.Visible = false;
            ProgData.B1Application.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);

        }
        private void FillMatrixdt(Recordset oRecSet, List<Grouppay> allpaymentx)
        {
            oForm.Freeze(true);
            int ix = -1;
            string _cardcode = "";
            loadbt.Item.Enabled = false;
            ((CheckBox)oForm.Items.Item("AllRec").Specific).Checked=true;
            allpayments = new List<Payment>();
            Grouppays = new List<Grouppay>();
            JES = new Dictionary<JesKey, List<JE>>();
            
            try
            {

                int yx = 0;
                Matrixdt.Rows.Clear();
                Payment payment = new Payment();
                while (!oRecSet.EoF)
                {
                    ix++;
                    Matrixdt.Rows.Add();
                    payment = new Payment();
                    Matrixdt.Rows.Offset = ix;
                   /* if (find)
                    {
                        setField(Matrixdt, ix, "Sel", oRecSet.Fields.Item("Selected").Value);
                        payment.Sel = oRecSet.Fields.Item("Selected").Value.ToString();
                        setField(Matrixdt, ix, "Linenum", oRecSet.Fields.Item("Linenum").Value);
                        payment.Linenum = Convert.ToInt32(oRecSet.Fields.Item("Linenum").Value);
                    }
                    else
                    {
                        setField(Matrixdt, ix, "Sel", "N");
                        payment.Sel = "N";
                         
                        
                        setField(Matrixdt, ix, "Linenum", ix);
                        payment.Linenum = ix;

                    }*/
                    setField(Matrixdt, ix, "Sel", "N");
                    payment.Sel = "N";


                    setField(Matrixdt, ix, "Linenum", ix);
                    payment.Linenum = ix;
                    setField(Matrixdt, ix, "DocEntry", oRecSet.Fields.Item("DocEntry").Value);
                    payment.DocEntry = (int)oRecSet.Fields.Item("DocEntry").Value;
                   
                    setField(Matrixdt, ix, "CardCode", oRecSet.Fields.Item("CardCode").Value.ToString());
                    payment.CardCode = oRecSet.Fields.Item("CardCode").Value.ToString();
                    setField(Matrixdt, ix, "DocNum", oRecSet.Fields.Item("DocNum").Value);
                    payment.DocNum = oRecSet.Fields.Item("DocNum").Value.ToString();
                  
                    setField(Matrixdt, ix, "DocType", oRecSet.Fields.Item("DocType").Value);
                    payment.DocType = oRecSet.Fields.Item("DocType").Value.ToString();
                    setField(Matrixdt, ix, "DocDate", oRecSet.Fields.Item("DocDate").Value);
                    payment.DocDate = Convert.ToDateTime(oRecSet.Fields.Item("DocDate").Value);
                    setField(Matrixdt, ix, "DueDate", oRecSet.Fields.Item("DocDueDate").Value);
                    payment.DueDate = (DateTime)oRecSet.Fields.Item("DocDueDate").Value;
                    setField(Matrixdt, ix, "DPastDue", oRecSet.Fields.Item("DaysPastDue").Value);
                    payment.DPastDue = oRecSet.Fields.Item("DaysPastDue").Value.ToString();
                    setField(Matrixdt, ix, "DocTotal", oRecSet.Fields.Item("DocTotal").Value);
                    payment.DocTotal = (double)oRecSet.Fields.Item("DocTotal").Value;
                    setField(Matrixdt, ix, "BalDue", oRecSet.Fields.Item("BalDue").Value);
                    payment.BalDue = (double)oRecSet.Fields.Item("BalDue").Value;
                    setField(Matrixdt, ix, "TotalDisc", oRecSet.Fields.Item("TotalDisc").Value);
                    payment.TotalDisc = (double)oRecSet.Fields.Item("TotalDisc").Value;
                    setField(Matrixdt, ix, "TotalPay", oRecSet.Fields.Item("TotalPay").Value);
                    payment.TotalPay = (double)oRecSet.Fields.Item("TotalPay").Value;
                    setField(Matrixdt, ix, "Processed", "N");
                    payment.PostedDocEntry = -1;
                    if (allpaymentx != null)
                    {
                        Grouppay px = new Grouppay();
                        grpPaymentdetail it= new grpPaymentdetail();
                        foreach (var item in allpaymentx)
                        {
                            it = item.paymentdetail.Find(x => x.DocNum == payment.DocNum && x.DocDocEntry == payment.DocEntry);
                            if (it.DocDocEntry != 0)
                            {
                                px = item;
                                break;
                            } 

                        }
             //           Grouppay px = allpaymentx.Find(x => x.paymentdetail.Find(y=>y.DocDocEntry == payment.DocEntry && y.DocNum == payment.DocNum).DocDocEntry == payment.DocEntry && x.CardCode == payment.CardCode);
                        if (px.paymentdetail != null && px.paymentdetail.Count > 0)
                        {

                            setField(Matrixdt, ix, "Sel", "Y");
                            setField(Matrixdt, ix, "TotalDisc", it.TotalDisc);
                            payment.TotalDisc = it.TotalDisc;
                            setField(Matrixdt, ix, "TotalPay", it.TotalPay);
                            payment.TotalPay = it.TotalPay;
                            payment.Sel = "Y";
                            try
                            {
                                setField(Matrixdt, ix, "PostedDocEntry", px.PostedJDocEntry);
                                payment.PostedDocEntry = px.PostedJDocEntry;
                            }
                            catch (Exception e)
                            {
                                setField(Matrixdt, ix, "PostedDocEntry", -1);
                                payment.PostedDocEntry = -1;
                            }
                            try
                            {

                                if (px.PostedJDocEntry > 0)
                                    setField(Matrixdt, ix, "Processed", "Y");


                            }
                            catch (Exception e)
                            {

                            }
                            if (payment.CardCode != _cardcode)
                                FillgrpMatrix(px);
                            _cardcode = payment.CardCode;
                        }
                    }
                    setField(Matrixdt, ix, "Rownum", ix);
                    payment.Row = ix;
                   
                   
                    allpayments.Add(payment);
                    yx++;
                    /* if (yx > 50)
                     {

                         break;

                     }*/
                    oRecSet.MoveNext();
                }

            }
            catch (Exception e)
            {
                Logger.Log(e);
            }

            oMatrix.LoadFromDataSourceEx();
            ogrpMatrix.LoadFromDataSourceEx();
            oForm.PaneLevel = 1;
            oForm.Freeze(false);
        }
        bool renk;
        private async Task renklendir(int startrow = 1, int say = 99999)
        {

            int errcolor = 252 | (221 << 8) | (130 << 16);

            int color = 255 | (255 << 8) | (255 << 16);


            int pcolor = 180 | (200 << 8) | (130 << 16);
            string pId = "";
            int ix = 0;
            if (startrow > Matrixdt.Rows.Count)
                startrow = 1;
            if (startrow < 1)
                startrow = 1;
            try
            {
                for (int rownum = startrow; rownum <= Matrixdt.Rows.Count; rownum++)
                {
                    if (!renk)
                        break;
                    int row = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cRownum").Cells.Item(rownum).Specific).Value);

                    if (row < 0)
                        row = 1;
                    if (row > Matrixdt.Rows.Count - 1)
                        row = rownum - 1;
                    if (oMatrix.CommonSetting.GetCellBackColor(rownum,10) == errcolor)
                    {
                        for (int fx = 1; fx < oMatrix.Columns.Count; fx++)
                        {
                            oMatrix.CommonSetting.SetCellBackColor(rownum, fx, cellcolor[fx-1]);
                        }
                    }
                    
                    if (Matrixdt.GetValue("Processed", row).ToString() == "Y")
                    {

                        oMatrix.CommonSetting.SetRowBackColor(rownum, pcolor);
                      
                        oMatrix.CommonSetting.SetRowEditable(rownum, false);
                    }
                    else
                    if (Matrixdt.GetValue("Sel", row).ToString() == "Y")
                    {

                        oMatrix.CommonSetting.SetRowBackColor(rownum, errcolor);

                        oMatrix.CommonSetting.SetRowEditable(rownum, true);
                    }
                   
                    ix++;
                    if (ix > say)
                        break;

                }
            }
            catch (Exception e)
            {
                Logger.Log(e);
            }
            Task t = Task.Delay(500);
            await t;
        }
        private void FillgrpMatrix(Grouppay gpay, bool find = false)
        {
            
            grpDt.Rows.Add();
            int ix = grpDt.Rows.Count - 1;
            grpDt.Rows.Offset = ix;
            grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, gpay.PayAcc);
           
            grpDt.SetValue("Rownum", grpDt.Rows.Offset, ix+1);
            gpay.Row = ix + 1;

            grpDt.SetValue("gCardCode", grpDt.Rows.Offset, gpay.CardCode);
          
            grpDt.SetValue("Linenum", grpDt.Rows.Offset, gpay.Linenum);
          
            SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string qry = $"Select \"CardName\" from \"OCRD\" Where \"CardCode\" = '{gpay.CardCode}'";
            orec.DoQuery(qry);
            try
            {
                grpDt.SetValue("gCardName", grpDt.Rows.Offset, orec.Fields.Item(0).Value);
            }
            catch (Exception e)
            {

            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orec);

            GC.Collect();
            try
            {
                grpDt.SetValue("gSelJ", grpDt.Rows.Offset, gpay.SelJ);
               
            }
            catch (Exception e)
            {
                grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "N");
                gpay.SelJ = "N";
            }
            try
            {
                
                grpDt.SetValue("JDocentry", grpDt.Rows.Offset, gpay.PostedJDocEntry);
            }
            catch (Exception e)
            {
                grpDt.SetValue("JDocentry", grpDt.Rows.Offset, "-1");
                gpay.PostedJDocEntry = -1;
            }

            try
            {
                grpDt.SetValue("gSelD", grpDt.Rows.Offset, gpay.SelP);
               
            }
            catch (Exception e)
            {
                grpDt.SetValue("gSelD", grpDt.Rows.Offset, "N");
                gpay.SelP = "N";
            }
            try
            {
                
                grpDt.SetValue("DPDocentry", grpDt.Rows.Offset, gpay.PostedDDocEntry);
            }
            catch (Exception e)
            {
                grpDt.SetValue("DPDocentry", grpDt.Rows.Offset, "-1");
                gpay.PostedJDocEntry = -1;
            }
            try
            {
              
                grpDt.SetValue("JMess", grpDt.Rows.Offset, gpay.JMess);
            }
            catch (Exception e)
            {
                grpDt.SetValue("JMess", grpDt.Rows.Offset, "");
              
            }
            try
            {

                grpDt.SetValue("DPMess", grpDt.Rows.Offset, gpay.DMess);
            }
            catch (Exception e)
            {
                grpDt.SetValue("DPMess", grpDt.Rows.Offset, "");

            }
            Grouppays.Add(gpay);
           // if (gpay.SelJ == "Y")
            {
                fillJERec(gpay.Linenum.ToString(),gpay.CardCode);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orec);
            
            GC.Collect();

        }
        private void fillJERec(string sumLinenum,string cardcode)
        {
            SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string qry = $"SELECT * FROM redi_ConsBP_PayAcct WHERE SumLinenum =  {sumLinenum} AND DocEntry = {HDocEntry.ToString()} ORDER BY LineNum";
            orec.DoQuery(qry);
            JesKey jkey = new JesKey();
            jkey.CardCode = cardcode;
            jkey.DocEntry = grpDt.Rows.Offset;
            List<JE> ljes = new List<JE>();
            while (!orec.EoF)
            {
                JE je = new JE();
                int ix = -1;
                je.Account = orec.Fields.Item("Account").Value.ToString();
                je.Debit = Convert.ToDouble(orec.Fields.Item("Debit").Value,CultureInfo.InvariantCulture);
                je.Credit = Convert.ToDouble(orec.Fields.Item("Credit").Value, CultureInfo.InvariantCulture);
                je.LineMemo = orec.Fields.Item("LineMemo").Value.ToString();
                je.DocType = orec.Fields.Item("DocType").Value.ToString();
                je.Docnum = Convert.ToInt32(orec.Fields.Item("DocNum").Value.ToString());
                je.LineNum = Convert.ToInt32(orec.Fields.Item("LineNum").Value.ToString());
                je.CardCode = orec.Fields.Item("CardCode").Value.ToString();                
                je.SumLinenum = Convert.ToInt32(orec.Fields.Item("SumLinenum").Value); 
                je.Posted = orec.Fields.Item("Posted").Value.ToString();
                je.PostDate = Convert.ToDateTime(orec.Fields.Item("PostDate").Value,CultureInfo.InvariantCulture);
                je.DocEntry = Convert.ToInt32(orec.Fields.Item("DocEntry").Value.ToString());
                ljes.Add(je);
                
                
                orec.MoveNext();
            }
            JES.Add(jkey, ljes);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orec);

            GC.Collect();
        }

        public void grpMatrix_DblClick(ItemEvent pVal)
        {
            string Poption = "N";
            string pdraft = "N";
            grpDt.Rows.Offset = Convert.ToInt32(((EditText)ogrpMatrix.GetCellSpecific("Rownum", pVal.Row)).Value);
            try
            {
                if (((CheckBox)ogrpMatrix.GetCellSpecific("SelJ", pVal.Row)).Checked)
                Poption = "Y";
            }
            catch (Exception e)
            {
                Poption = "N";
            }
            try
            {
                if (((CheckBox)ogrpMatrix.GetCellSpecific("SelD", pVal.Row)).Checked)
                    pdraft = "Y";
            }
            catch (Exception e)
            {
                pdraft = "N";
            }
            string Jref = grpDt.GetValue("JDocentry", grpDt.Rows.Offset).ToString();
            string Dref = grpDt.GetValue("DPDocentry", grpDt.Rows.Offset).ToString();
            if (Convert.ToInt32(Jref) > 0)
                ProgData.B1Application.OpenForm(BoFormObjectEnum.fo_JournalPosting, "", Jref);
            if (Convert.ToInt32(Dref) > 0)
            {

                ProgData.B1Application.OpenForm((BoFormObjectEnum)140, "", Dref);
            }
           
            if (Convert.ToInt32(Jref) < 1 )
            {
                string cardCode = ((EditText)ogrpMatrix.GetCellSpecific("cCodex", pVal.Row)).Value;
                int ftop = ogrpMatrix.Item.Top + (pVal.Row * 20);
                double payonacctot = Convert.ToDouble(((EditText)ogrpMatrix.GetCellSpecific("cPayaccx", pVal.Row)).Value, CultureInfo.InvariantCulture);
                int gix = Grouppays.FindIndex(x => x.CardCode == cardCode && x.PostedJDocEntry < 1);

                if (gix > 0)
                {

                    Grouppay gpay = new Grouppay();
                    grpDt.Rows.Offset = pVal.Row - 1;

                    var jkey = buildJE(cardCode, payonacctot, gpay.Linenum);

                    JEPopup jepopup = new JEPopup();
                    jepopup.AddControls(JES[jkey], jkey, ftop, this,true);
                }
            }
        }
        public void grpMatrix_LostFocusAfter(ItemEvent pVal)
        {
            string cardCode = ((EditText)ogrpMatrix.GetCellSpecific("cCodex", pVal.Row)).Value;
            if (oForm.Mode != BoFormMode.fm_ADD_MODE)
            {
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
            }
            try
            {
                UD_13.ValueEx = Logger.exeFolder + @"\addfield.jpg";
             //   ((PictureBox)ogrpMatrix.Columns.Item("cbtn").Cells.Item(pVal.Row).Specific).Picture = Logger.exeFolder + @"\addfield.jpg";
               
            } catch (Exception e)
            {

            }
            grpDt.Rows.Offset = Convert.ToInt32(((EditText)ogrpMatrix.GetCellSpecific("Rownum", pVal.Row)).Value);
            double payonacctot = Convert.ToDouble(((EditText)ogrpMatrix.GetCellSpecific("cPayaccx", pVal.Row)).Value, CultureInfo.InvariantCulture);
            ogrpMatrix.GetLineData(pVal.Row);
            if (payonacctot == 0 && pVal.ColUID == "cPayaccx")
            {
                ogrpMatrix.CommonSetting.SetCellEditable(pVal.Row, 9, false);
                ogrpMatrix.CommonSetting.SetCellEditable(pVal.Row, 10, false);
                grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "N");
            } else
            if (payonacctot > 0 && pVal.ColUID == "cPayaccx")
            {
                ogrpMatrix.CommonSetting.SetCellEditable(pVal.Row, 9, true);
                ogrpMatrix.CommonSetting.SetCellEditable(pVal.Row, 10, true);

            }
            if (pVal.ColUID == "SelJ")
            {
                if (((CheckBox)ogrpMatrix.GetCellSpecific("SelJ", pVal.Row)).Checked)
                {

                    grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "Y");
                } else
                    grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "N");
            }
            grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, payonacctot);
            ogrpMatrix.SetLineData(pVal.Row);
            int gix = Grouppays.FindIndex(x => x.Linenum == grpDt.Rows.Offset);
            if (gix>-1)
            {
                var grp = Grouppays[gix];
                grp.PayAcc = payonacctot;
                grp.SelJ = grpDt.GetValue("gSelJ", grpDt.Rows.Offset).ToString();
                grp.SelP = grpDt.GetValue("gSelD", grpDt.Rows.Offset).ToString();
                // grp.Poption = "JE";
                Grouppays[gix] = grp;
            }
           


        }
        private void PayOption(ItemEvent pVal)
        {
            string cardCode = ((EditText)ogrpMatrix.GetCellSpecific("cCodex", pVal.Row)).Value;
            string Poption = "N";
            try
            {
                if (((CheckBox)ogrpMatrix.GetCellSpecific("SelJ", pVal.Row)).Checked)
                    Poption = "Y";
            }
            catch (Exception e)
            {
                Poption = "N";
            }
            double payonacctot = Convert.ToDouble(((EditText)ogrpMatrix.GetCellSpecific("cPayaccx", pVal.Row)).Value, CultureInfo.InvariantCulture);

            grpDt.Rows.Offset = Convert.ToInt32(((EditText)ogrpMatrix.GetCellSpecific("Rownum", pVal.Row)).Value);
            grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, payonacctot);
            int gix = Grouppays.FindIndex(x => x.Linenum == grpDt.Rows.Offset);
            if (gix > -1)
            {
                var grp = Grouppays[gix];
                
                grp.SelJ = Poption;
                Grouppays[gix] = grp;
             
            }

          
        }
        private void JEBtnClick(ItemEvent pVal)
        {
            string cardCode = ((EditText)ogrpMatrix.GetCellSpecific("cCodex", pVal.Row)).Value;
           // string Poption = "N";
          
          /*  try
            {
                if (((CheckBox)ogrpMatrix.GetCellSpecific("SelJ", pVal.Row)).Checked)
                    Poption = "Y";
            }
            catch (Exception e)
            {
                Poption = "N";
            }*/
            double payonacctot = Convert.ToDouble(((EditText)ogrpMatrix.GetCellSpecific("cPayaccx", pVal.Row)).Value, CultureInfo.InvariantCulture);
          //  if (Poption == "Y")
            {
                int Jref = 0;
                try
                {
                   Jref = Convert.ToInt32( grpDt.GetValue("JDocentry", grpDt.Rows.Offset).ToString());
                } catch (Exception e)
                {
                    Jref = 0;
                }
                bool editable = true;
                if (Jref > 0)
                    editable = false;
                grpDt.Rows.Offset = pVal.Row - 1;
                int ftop = ogrpMatrix.Item.Top + (pVal.Row * 20);
                var jkey = buildJE(cardCode, payonacctot,pVal.Row-1);
                JEPopup jepopup = new JEPopup();
                jepopup.AddControls(JES[jkey], jkey, ftop, this,editable);
            }
        }
        private JesKey buildJE(string cardCode,double payonacctot,int Linenum)
        {
            JesKey jkey = new JesKey();
            jkey.CardCode = cardCode;
            jkey.DocEntry = grpDt.Rows.Offset;
            List<JE> ljes = new List<JE>();
            JE je = new JE();
            int ix = -1;
 
            if (JES.ContainsKey(jkey))
            {
                ljes = JES[jkey];
                if (ljes.Count == 0)
                {
                    ljes = newJEList(cardCode, payonacctot, Linenum);
                }
                
                
                    ix = 0; // ljes.FindIndex(x => x.Account == cardCode);
                    je = ljes[ix];
                    je.Debit = payonacctot;
                    ljes[ix] = je;
                    JES[jkey] = ljes;
                
            }
            else
            {
                ljes = newJEList(cardCode, payonacctot, Linenum);
                JES.Add(jkey, ljes);
            }
            return jkey;
           
            
        }
        private List<JE> newJEList(string cardCode, double payonacctot, int Linenum)
        {
            JesKey jkey = new JesKey();
            jkey.CardCode = cardCode;
            jkey.DocEntry = grpDt.Rows.Offset;
            List<JE> ljes = new List<JE>();
            JE je = new JE();
            /*SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string qry = $"select c.DebPayAcct, a.Segment_0 + '-' + a.Segment_1 + '-' + a.Segment_2 as GLAcct from ocrd c inner join oact a on a.acctcode = c.DebPayAcct  where cardcode = '{cardCode}'";
            orec.DoQuery(qry);
            string glacc = orec.Fields.Item(1).Value.ToString();*/
            je.Account = "10121-00-200";
            je.Debit = payonacctot;
            je.LineMemo = "";
            je.Credit = 0;
            je.DocType = "";
            je.Docnum = 0;
            je.LineNum = 0;
            je.PostDate = DateTime.Today;
            je.CardCode = "";
            je.SumLinenum = Linenum;
            ljes.Add(je);
            je = new JE();
            je.Account = cardCode;
            je.CardCode = cardCode;
            je.Credit = payonacctot;
            je.PostDate = DateTime.Today;
            je.LineMemo = "";
            je.Debit = 0;
            je.DocType = "";
            je.Docnum = 0;
            je.SumLinenum = Linenum;
            je.LineNum = 1;
            ljes.Add(je);

            return ljes;
        }
        public void oMatrix_LostFocusAfter(ItemEvent pVal)
        {
            string data = "";
            string fname = "";
            if (pVal.Row < 1)
                return;
            try
            {
                if (oMatrix.Columns.Item(pVal.ColUID).Type == BoFormItemTypes.it_CHECK_BOX)
                {

                    if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Checked == true)
                    {
                        try
                        {
                            string v = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).ValOn;
                            v = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).ValOff;
                        }
                        catch (Exception e)
                        {

                        }
                        UpdateList(pVal, false);

                    }
                    else
                    {

                        UpdateList(pVal, true);

                    }

                    //  Matrixdt.SetValue(fname, Matrixdt.Rows.Offset, data);

                }


                SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(pVal.ColUID);
                fname = oColumn.DataBind.Alias;
                double paytot = Convert.ToDouble(((EditText)oMatrix.GetCellSpecific("cTotalPay", pVal.Row)).Value, CultureInfo.InvariantCulture);
                double disctot = Convert.ToDouble(((EditText)oMatrix.GetCellSpecific("cTotalDisc", pVal.Row)).Value, CultureInfo.InvariantCulture);
                if ((paytot == 0 && disctot == 0) &&
                    (pVal.ColUID.CompareTo("cTotalPay") == 0 || pVal.ColUID.CompareTo("cTotalDisc") == 0))
                {
                    UpdateList(pVal, true);
                }
                else
                if (pVal.ColUID.CompareTo("cTotalPay") == 0 || pVal.ColUID.CompareTo("cTotalDisc") == 0 &&
                    (paytot != 0 || disctot != 0))
                {
                    UpdateList(pVal, false);
                }

            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + pVal.ColUID + "-//fldname=" + fname));
            }

        }
        private void UpdateList(ItemEvent pVal, bool delete)
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(pVal.ColUID);

            string fname = oColumn.DataBind.Alias;
            string data = "N";
            Matrixdt.Rows.Offset = Convert.ToInt32(((EditText)oMatrix.GetCellSpecific("cRownum", pVal.Row)).Value);

            int color = 255 | (255 << 8) | (255 << 16);
            if (!delete)
            {
                data = "Y";
                color = 252 | (221 << 8) | (130 << 16);
                oMatrix.CommonSetting.SetRowBackColor(pVal.Row, color);
            }
            else
            {
                for (int ix = 1; ix < oMatrix.Columns.Count; ix++)
                {
                    oMatrix.CommonSetting.SetCellBackColor(pVal.Row, ix, cellcolor[ix - 1]);
                }

            }

            double discTotal = 0;
            double payTotal = 0;
            double docTotal = 0;
            double balDue = 0;
            string cardcode = "";
            string docType = "";
            string DocNum = "";
            int Linenum = 0;
            try
            {
                balDue = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cBalDue").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }
            try
            {

                Linenum = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cLinenum").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }

            try
            {
                discTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotalDisc").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }
            try
            {
                payTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotalPay").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }
            try
            {
                docTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocTotal").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }
            try
            {
                cardcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cCardCode").Cells.Item(pVal.Row).Specific).Value;
            }
            catch (Exception e)
            {


            }
            try
            {
                docType = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocType").Cells.Item(pVal.Row).Specific).Value;
            }
            catch (Exception e)
            {

            }
            try
            {
                DocNum = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocNum").Cells.Item(pVal.Row).Specific).Value;
            }
            catch (Exception e)
            {

            }

            if (pVal.ColUID == "cTotalDisc")
            {
                payTotal = balDue - discTotal;

            }
            else
                if (pVal.ColUID == "cTotalPay")
            {

                discTotal = balDue - payTotal;
            }

            try
            {

                int pix = allpayments.FindIndex(pym => pym.CardCode == cardcode && pym.Linenum == Linenum);
                Matrixdt.SetValue("TotalDisc", Matrixdt.Rows.Offset, discTotal);
                Matrixdt.SetValue("TotalPay", Matrixdt.Rows.Offset, payTotal);
                Matrixdt.SetValue("Sel", Matrixdt.Rows.Offset, data);
                if (allpayments[pix].Sel == "Y" && data == "N")
                {
                    int gix = Grouppays.FindIndex(x => x.CardCode == allpayments[pix].CardCode && x.PostedJDocEntry == allpayments[pix].PostedDocEntry);
                    if (gix > -1)
                    {
                        Grouppay grpy = Grouppays[gix];
                        int pdix = grpy.paymentdetail.FindIndex(x => x.DocDocEntry == allpayments[pix].DocEntry && x.DocNum == allpayments[pix].DocNum);
                        if (pdix>-1)
                        {
                            grpy.paymentdetail.RemoveAt(pdix);
                            
                        }
                        if (grpy.paymentdetail.Count ==0)
                        {
                            grpy.TotalDisc = 0;
                            grpy.TotalPay = 0;
                            grpy.PayAcc = 0;
                            grpy.DocNum = "";
                            grpy.DocTotal = 0;
                            grpDt.Rows.Offset = grpy.Linenum;
                            grpDt.SetValue("gTotal", grpDt.Rows.Offset, 0.00);
                            grpDt.SetValue("gDiscount", grpDt.Rows.Offset, 0.00);
                            grpDt.SetValue("gPayment", grpDt.Rows.Offset, 0.00);
                            grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, 0.00);

                        }
                    }
                }
                
                allpayments[pix].Sel = data;
                allpayments[pix].TotalDisc = discTotal;
                allpayments[pix].TotalPay = payTotal;
                if (oForm.Mode != BoFormMode.fm_ADD_MODE)
                {
                    oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                }

                if (pVal.EventType == BoEventTypes.et_LOST_FOCUS)
                    oMatrix.SetLineData(pVal.Row);
               this.Listgrup();
            }
            catch (Exception e)
            {
                Logger.Log(e);
            }
        }
        public void oMatrix_GotFocusAfter(ItemEvent pVal)
        {
            string fname = "";
            string data = "N";
            try
            {
                if (pVal.Row < 0)
                {
                    return;
                }

                if (oMatrix.Columns.Item(pVal.ColUID).Type == BoFormItemTypes.it_CHECK_BOX)
                {
                    Matrixdt.Rows.Offset = Convert.ToInt32(((EditText)oMatrix.GetCellSpecific("cRownum", pVal.Row)).Value);
                    data = Matrixdt.GetValue("Sel", Matrixdt.Rows.Offset).ToString();
                    int colorN = 255 | (255 << 8) | (255 << 16);
                    int colorY = 252 | (221 << 8) | (130 << 16);



                    if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Checked == true)
                    {
                        // oMatrix.CommonSetting.SetRowBackColor(pVal.Row, colorN);
                        UpdateList(pVal, false);

                    }
                    else
                    if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Checked == false)
                    {
                        // oMatrix.CommonSetting.SetRowBackColor(pVal.Row, colorY);
                        UpdateList(pVal, true);
                    }

                    //  Matrixdt.SetValue(fname, Matrixdt.Rows.Offset, data);

                }





            }
            catch (Exception er)
            {
                Logger.Log(new Exception(er.Message + pVal.ColUID + "--fldname=" + fname));
            }

        }
       
        public void Listgrup()
        {
            grpDt = oForm.DataSources.DataTables.Item("GrpPay");


            var groupedCustomerList = allpayments.Where(x => x.Sel == "Y" ).GroupBy(u => new { u.CardCode, u.PostedDocEntry })
                                    .Select(grp => new { Code = grp.Key, Payment = grp.ToList(), DocTotal = grp.Sum(s => s.DocTotal), DiscTotal = grp.Sum(s => s.TotalDisc), PaymentTotal = grp.Sum(s => s.TotalPay) })
                                    .ToList();

            
            ogrpMatrix.Clear();

            
            SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            Grouppay gpay = new Grouppay();
            for (int ix = 0; ix < groupedCustomerList.Count; ix++)
            {
                ogrpMatrix.Item.Enabled = true;
                var payment = groupedCustomerList[ix];
               // if (payment.Code.PostedDocEntry > 0)
                 //   continue;
                int gix  = Grouppays.FindIndex(x => x.CardCode == payment.Code.CardCode && x.PostedJDocEntry == payment.Code.PostedDocEntry);
                
                if (gix<0)
                {
                    
                    gpay = new Grouppay();
                    grpDt.Rows.Add();
                    grpDt.Rows.Offset = grpDt.Rows.Count-1;
                    grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, 0.00);
                    gpay.PayAcc = 0;
                    gpay.PostedJDocEntry = -1;
                    gpay.Linenum = grpDt.Rows.Offset;
                    grpDt.SetValue("Rownum", grpDt.Rows.Offset, ix);
                    gpay.Row = ix;
                    grpDt.SetValue("gCardCode", grpDt.Rows.Offset, payment.Code.CardCode);
                    grpDt.SetValue("Linenum", grpDt.Rows.Offset, grpDt.Rows.Offset);
                    gpay.CardCode = payment.Code.CardCode;
                   
                    string qry = $"Select \"CardName\" from \"OCRD\" Where \"CardCode\" = '{payment.Code.CardCode}'";
                    orec.DoQuery(qry);
                    try
                    {
                        grpDt.SetValue("gCardName", grpDt.Rows.Offset, orec.Fields.Item(0).Value);
                    }
                    catch (Exception e)
                    {

                    }
                    try
                    {
                        grpDt.SetValue("gTotal", grpDt.Rows.Offset, Convert.ToDouble(payment.DocTotal));
                    } catch (Exception e)
                    {
                        grpDt.SetValue("gTotal", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gDiscount", grpDt.Rows.Offset, Convert.ToDouble(payment.DiscTotal));
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gDiscount", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gPayment", grpDt.Rows.Offset, Convert.ToDouble(payment.PaymentTotal));
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gPayment", grpDt.Rows.Offset, 0.00);
                    }
                    gpay.SelJ = "N";
                    try
                    {
                        grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "N");
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gSelJ", grpDt.Rows.Offset,"N");
                    }
                    gpay.SelP = "N";
                    try
                    {
                        grpDt.SetValue("gSelD", grpDt.Rows.Offset, "N");
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gSelD", grpDt.Rows.Offset, "N");
                    }
                    gpay.DocTotal = payment.DocTotal;
                    gpay.TotalDisc = payment.DiscTotal;
                    gpay.TotalPay = payment.PaymentTotal;
                    gpay.paymentdetail = new List<grpPaymentdetail>();
                    foreach (var item in payment.Payment)
                    {
                        grpPaymentdetail gdetail = new grpPaymentdetail();
                        gdetail.DocDocEntry = item.DocEntry;
                        gdetail.DocNum = item.DocNum;
                        gdetail.DocType = item.DocType;
                        gdetail.BalDue = item.BalDue;
                        gdetail.TotalDisc = item.TotalDisc;
                        gdetail.TotalPay = item.TotalPay;
                       
                        gpay.paymentdetail.Add(gdetail);
                    }
                    Grouppays.Add(gpay);
                }
                else
                {
                    gpay = Grouppays[gix];
                    grpDt.Rows.Offset = gpay.Linenum;
                    try
                    { 
                    grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, gpay.PayAcc);
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gPaymonAcc", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gTotal", grpDt.Rows.Offset, Convert.ToDouble(payment.DocTotal));
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gTotal", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gDiscount", grpDt.Rows.Offset, Convert.ToDouble(payment.DiscTotal));
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gDiscount", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gPayment", grpDt.Rows.Offset, Convert.ToDouble(payment.PaymentTotal));
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gPayment", grpDt.Rows.Offset, 0.00);
                    }
                    try
                    {
                        grpDt.SetValue("gSelJ", grpDt.Rows.Offset, gpay.SelJ);
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gSelJ", grpDt.Rows.Offset, "N");
                    }
                    try
                    {
                        grpDt.SetValue("gSelD", grpDt.Rows.Offset, gpay.SelP);
                    }
                    catch (Exception e)
                    {
                        grpDt.SetValue("gSelD", grpDt.Rows.Offset, "N");
                    }

                    gpay.TotalDisc = payment.DiscTotal;
                    gpay.TotalPay = payment.PaymentTotal;
                    gpay.paymentdetail = new List<grpPaymentdetail>();
                    foreach (var item in payment.Payment)
                    {
                        grpPaymentdetail gdetail = new grpPaymentdetail();
                        gdetail.DocDocEntry = item.DocEntry;
                        gdetail.DocNum = item.DocNum;
                        gdetail.DocType = item.DocType;
                        gdetail.TotalPay = item.TotalPay;
                        gdetail.TotalDisc = item.TotalDisc;
                        gdetail.BalDue = item.BalDue;
                       
                        gpay.paymentdetail.Add(gdetail);
                    }
                    if (gpay.paymentdetail.Count == 0)
                    {
                        grpDt.Rows.Offset = gpay.Linenum;
                        grpDt.Rows.Remove(gpay.Linenum);
                    }
                    Grouppays[gix] = gpay;
                }
              

            }
            ogrpMatrix.LoadFromDataSourceEx();
            int fx = 0;
            foreach (var item in Grouppays)
            {
                fx++;
                if (item.paymentdetail.Count == 0)
                {
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 7, false);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 8, false);
                        ogrpMatrix.CommonSetting.SetCellEditable(fx, 9, false);
                        ogrpMatrix.CommonSetting.SetCellEditable(fx, 10, false);
                   
                } else
                {
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 7, true);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 8, true);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 9, true);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 10, true);

                }
                if (item.PostedJDocEntry>0)
                {
                    ogrpMatrix.CommonSetting.SetCellEditable(fx , 8, false);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 9, false);
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 10, false);
                }
                if (item.PostedDDocEntry > 0)
                {
                    
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 7, false);
                }
                if (item.PayAcc == 0)
                {
                  
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 9, false);
                   
                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 10, false); 
                } else
                    if (item.PayAcc == 0)
                {

                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 9, true);

                    ogrpMatrix.CommonSetting.SetCellEditable(fx, 10, true);
                }
                /*  if (item.PostedDocEntry ==-1)
                 {
                     int fix = groupedCustomerList.FindIndex(x => x.Code.CardCode == item.CardCode);
                     if (fix<0)
                     {
                         grpDt.Rows.Remove(item.Linenum);
                         Grouppays.Remove(item);
                     }
                 }*/
            }
          /*  for (int ix = 0;ix<grpDt.Rows.Count;ix++)
            {
                if (Convert.ToDouble(grpDt.GetValue("gPaymonAcc", ix)) == 0)
                {
                    ogrpMatrix.CommonSetting.SetCellEditable(ix + 1, 8, false);
                }
            }
          */
            System.Runtime.InteropServices.Marshal.ReleaseComObject(orec);

            GC.Collect();
        }
        private void LinkedPress(ItemEvent pVal)
        {
            try
            {
                oForm.Freeze(true);
                Column oColumn = oMatrix.Columns.Item(pVal.ColUID);
                LinkedButton oLink = (LinkedButton)oColumn.ExtendedObject;
                int rw = Convert.ToInt32(((EditText)oMatrix.GetCellSpecific("cRownum", pVal.Row)).Value);
                Matrixdt.Rows.Offset = Convert.ToInt32(((EditText)oMatrix.GetCellSpecific("cRownum", pVal.Row)).Value);
                string data = Matrixdt.GetValue("DocType", Matrixdt.Rows.Offset).ToString();
                string docnum = Matrixdt.GetValue("DocNum", Matrixdt.Rows.Offset).ToString();
                string docentry = Matrixdt.GetValue("DocEntry", Matrixdt.Rows.Offset).ToString();
                ((EditText)oMatrix.GetCellSpecific("cDocNum", pVal.Row)).Value = docentry;
                // Matrixdt.SetValue("DocNum", Matrixdt.Rows.Offset, docentry);
                if (data == "IN")
                {

                    oLink.LinkedObjectType = "13";
                }
                else
                    if (data == "CM")
                {

                    oLink.LinkedObjectType = "14";
                }
                else
                    if (data == "JE")
                {
                    oLink.LinkedObjectType = "30";

                }
                Linked = true;
                crow = pVal.Row;
            }
            catch (Exception e)
            {

            }
        }
        private void Filter(ItemEvent pVal)
        {
            string chk = "";
            string cardcode = "";
            string docnum = "";
            try
            {
                List<Payment> filterlst = new List<Payment>(allpayments);
                
                    CheckBox fchk = (CheckBox)oForm.Items.Item("fselect").Specific;

                    if (fchk.Checked == true)
                    {
                        chk = "Y";
                        filterlst = filterlst.FindAll(x => x.Sel == "Y");
                    }
                if (!((CheckBox)oForm.Items.Item("AllRec").Specific).Checked)
                {
                    if (((EditText)oForm.Items.Item("fcardcode").Specific).Value.Length > 0)
                    {
                        cardcode = ((EditText)oForm.Items.Item("fcardcode").Specific).Value;
                        filterlst = filterlst.FindAll(x => x.CardCode == cardcode);
                    }
                    if (((EditText)oForm.Items.Item("fdocnum").Specific).Value.Length > 0)
                    {
                        docnum = ((EditText)oForm.Items.Item("fdocnum").Specific).Value;
                        string[] sdocnum = docnum.Split(',');
                        filterlst = filterlst.Where(x => sdocnum.Contains(x.DocNum)).ToList(); 
                           // .FindAll(x => x.DocNum == docnum);
                    }
                }
                oForm.Freeze(true);
                Matrixdt.Rows.Clear();
                oMatrix.Clear();
                for (int ix = 0; ix < filterlst.Count; ix++)
                {

                    Matrixdt.Rows.Add();

                    Matrixdt.Rows.Offset = ix;


                    setField(Matrixdt, ix, "Sel", filterlst[ix].Sel);
                    setField(Matrixdt, ix, "CardCode", filterlst[ix].CardCode);
                    setField(Matrixdt, ix, "DocNum", filterlst[ix].DocNum);
                    setField(Matrixdt, ix, "DocEntry", filterlst[ix].DocEntry);
                    setField(Matrixdt, ix, "DocType", filterlst[ix].DocType);
                    setField(Matrixdt, ix, "DocDate", filterlst[ix].DocDate);
                    setField(Matrixdt, ix, "DueDate", filterlst[ix].DueDate);
                    setField(Matrixdt, ix, "DPastDue", filterlst[ix].DPastDue);
                    setField(Matrixdt, ix, "DocTotal", filterlst[ix].DocTotal);
                    setField(Matrixdt, ix, "BalDue", filterlst[ix].BalDue);
                    setField(Matrixdt, ix, "TotalDisc", filterlst[ix].TotalDisc);
                    setField(Matrixdt, ix, "TotalPay", filterlst[ix].TotalPay);
                    setField(Matrixdt, ix, "Linenum", filterlst[ix].Linenum);


                    try
                    {
                        setField(Matrixdt, ix, "PostedDocEntry", filterlst[ix].PostedDocEntry);

                    }
                    catch (Exception e)
                    {
                        setField(Matrixdt, ix, "PostedDocEntry", -1);
                    }
                    try
                    {
                        setField(Matrixdt, ix, "Processed", "N");
                        if (Convert.ToInt32(filterlst[ix].PostedDocEntry) > 0)
                            setField(Matrixdt, ix, "Processed", "Y");


                    }
                    catch (Exception e)
                    {

                    }

                    setField(Matrixdt, ix, "Rownum", ix);

                }
                oMatrix.LoadFromDataSourceEx();
                oForm.Freeze(false);
            }
            catch (Exception e)
            {
                Logger.Log(e);
            }

        }
        public void EventAll(ItemEvent pVal)
        {
            /*if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && Linked)
            {
                oMatrix.SetLineData(crow);
                Linked = false;
               // oMatrix.LoadFromDataSourceEx();
            }*/
            if (pVal.BeforeAction == true)
            {
                if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ColUID == "cDocNum")
                {
                    LinkedPress(pVal);
                }
                else
                     if (pVal.ItemUID == "txtDocNum" && pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.CharPressed == 13
                          && oForm.Mode == BoFormMode.fm_FIND_MODE)
                {
                    try
                    {

                        string docnum = ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocNum").Specific).Value;
                        if (docnum.Length > 0)
                            this.Find(docnum);

                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }
                }
            }
            else
            if (pVal.BeforeAction == false)
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "fselect")
                {

                    Filter(pVal);
                }
                else
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "AllRec")
                {
                    Filter(pVal);
                }
                else
                if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ColUID == "cDocNum")
                {
                    oMatrix.SetLineData(crow);
                    Linked = false;
                    oForm.Freeze(false);
                    // oMatrix.LoadFromDataSourceEx();
                }
                else
                if (pVal.ItemUID == "btnlist")
                {
                    clickbtnlist(pVal);
                }
                else
                 if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && (pVal.ItemUID == "txtCCode" || pVal.ItemUID == "txtCName"))
                {
                    try
                    {
                        this.EditText_ChooseFromListAfter(pVal);
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }
                else
               if (pVal.ItemUID == "fbtn" && pVal.EventType == BoEventTypes.et_CLICK)
                {
                    ((CheckBox)oForm.Items.Item("AllRec").Specific).Checked = false;
                    Filter(pVal);
                }
                else
             if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "loadbt")
                {
                    try
                    {

                        loadbt = (Button)oForm.Items.Item("loadbt").Specific;

                        if (loadbt.Item.Enabled)
                        {
                            if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                string code = ((SAPbouiCOM.EditText)oForm.Items.Item("txtCCode").Specific).Value;
                                this.Bind(code);
                            }
                            else
                            {
                                string docnum = ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocNum").Specific).Value;
                                if (docnum.Length > 0)
                                    this.Find(docnum);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }
                else
             if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" &&
                       (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE ))
                {
                    try
                    {
                        this.Save();
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }
                else
                 if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1")
                {
                    oForm.Close();
                }
                else

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "fldGrpTot")
                {
                    try
                    {
                        this.Listgrup();
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }
                else


                if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK || pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN) && pVal.ColUID == "cSel")

                {
                    try
                    {
                        this.oMatrix_GotFocusAfter(pVal);
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }

                else
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == "mtpreMtx")
                {
                    try
                    {
                        this.oMatrix_LostFocusAfter(pVal);
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }

                else
                if (pVal.ItemUID == "mtgrpMtx")
                {
                   
                    
                    if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ColUID == "cbtn")
                    {
                        JEBtnClick(pVal);
                    }
                    else
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
                    {
                        try
                        {
                            this.grpMatrix_DblClick(pVal);
                        }
                        catch (Exception e)
                        {
                            Logger.Log(e);
                        }
                    }
                    else

                    if (pVal.EventType == BoEventTypes.et_VALIDATE )
                    {
                        try
                        {
                            this.grpMatrix_LostFocusAfter(pVal);
                        }
                        catch (Exception e)
                        {
                            Logger.Log(e);
                        }

                    }
                }

            }
        }

        public void Menu(ref MenuEvent pVal)
        {
            if (pVal.BeforeAction == false && (pVal.MenuUID == "1282" || pVal.MenuUID == "1281" || pVal.MenuUID == "1288"
                || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291"))
            {
                this.NavRec(pVal.MenuUID);
            }
        }
    }

}
