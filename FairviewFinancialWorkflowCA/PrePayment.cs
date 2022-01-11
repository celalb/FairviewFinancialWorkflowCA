using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialWorkflowCA
{
    public struct Payment
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string DocType { get; set; }
        public string DocNum { get; set; }
        public Double DocTotal { get; set; }
        public Double DiscTotal { get; set; }
        public Double PaymentTotal { get; set; }
        public int Row { get; set; }

        
    }
    public struct JE
    {

        
         public string CardCode { get; set; }
         public DateTime RefDate { get; set; }
         public int LineNum { get; set; }
         public string Account { get; set; }
        public string DocType { get; set; }
        public int Docnum { get; set; }
        public string LineMemo { get; set; }
        public Decimal Debit { get; set; }
        public Decimal Credit { get; set; }
       
    }
    public class PrePayment : IForm
    {
        protected SAPbouiCOM.Form oForm = null;
      
        SAPbouiCOM.Matrix ogrpMatrix, oMatrix;
        DataTable Matrixdt, grpDt;
        double totdiscTotal;
        double totDocTotal;
        
        double totPActotal,totpayment;
        UserDataSource UD_Total, UD_TotalPay, UD_PaTotal, UD_DcTotal,UD_8,UD_6;

        Dictionary<int, List<JE>> JES;
        List<Payment> payments = new List<Payment>();
        public void AddJES(List<JE> jelst)
        {
/*            if (JES == null)
            {
                JES = new Dictionary<int, List<JE>>();
            }
            JES.Add()*/
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
                   
                    B1Starter.Forms.Add(oForm.UniqueID,this);
                    oForm.AutoManaged = true;
                    oForm.Visible = true;
                    ogrpMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("mtgrpMtx").Specific);
                    oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("mtpreMtx").Specific);
                    Matrixdt = oForm.DataSources.DataTables.Add("PrePay");
                    Matrixdt.Columns.Add("Sel", BoFieldsType.ft_AlphaNumeric, 1);
                    Matrixdt.Columns.Add("CardCode", BoFieldsType.ft_AlphaNumeric, 15);
                    Matrixdt.Columns.Add("DocNum", BoFieldsType.ft_Integer, 11);
                    Matrixdt.Columns.Add("DocType", BoFieldsType.ft_AlphaNumeric, 1);
                    Matrixdt.Columns.Add("DocDate", BoFieldsType.ft_Date);
                    Matrixdt.Columns.Add("DueDate", BoFieldsType.ft_Date);
                    Matrixdt.Columns.Add("DPastDue", BoFieldsType.ft_AlphaNumeric, 4);                   
                    Matrixdt.Columns.Add("DocTotal", BoFieldsType.ft_Price);
                    Matrixdt.Columns.Add("BalDue", BoFieldsType.ft_Price);
                    Matrixdt.Columns.Add("TotDisc", BoFieldsType.ft_Price);
                    Matrixdt.Columns.Add("TotPay", BoFieldsType.ft_Price);
                    Matrixdt.Columns.Add("Rownum", BoFieldsType.ft_Integer);
                    grpDt = oForm.DataSources.DataTables.Add("GrpPay");
                    // ChooseFromList Blist =   oForm.ChooseFromLists.Item("BList");
                    // Conditions cons = Blist.GetConditions();

                    Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                 //   DateTime buildDate = new DateTime(2000, 1, 1)
                   //                         .AddDays(version.Build).AddSeconds(version.Revision * 2);
                    string displayableVersion = $"({version})";
                    oForm.Title = $"Pre Payments {displayableVersion}";
                   
                    oForm.ActiveItem = "txtCCode";
                    UD_Total = oForm.DataSources.UserDataSources.Item("UD_Total");
                    UD_DcTotal = oForm.DataSources.UserDataSources.Item("UD_DcTotal");
                    UD_TotalPay = oForm.DataSources.UserDataSources.Item("UD_TotPay");
                    UD_PaTotal = oForm.DataSources.UserDataSources.Item("UD_PaTotal");
                    UD_8 = oForm.DataSources.UserDataSources.Item("UD_8");
                    UD_8.ValueEx = DateTime.Now.ToString("yyyyMMdd");
                    UD_6 = oForm.DataSources.UserDataSources.Item("UD_6");
                    UD_6.ValueEx = DateTime.Now.ToString("yyyyMMdd");

                }


            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }

        }
        public void EditText_ChooseFromListAfter(ItemEvent pVal)
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

                itemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));
                   SAPbouiCOM.EditText EditText = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                EditText.Value = itemCode;
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

                int docEntry = dbutil.InsertHeader(cardCode, docDate, dueDate, docRef, user);

                int docNum = 0;
                string cardName = "";
                string docType = "";
                string daysPastDue = "";
                Double docTotal = 0;
                Double balDue = 0;
                Double discTotal = 0;
                Double payTotal = 0;
                string selected = "N";
                
                for (int ix = 1;ix<= oMatrix.RowCount ; ix++)
                {
                    selected = "N";
                    if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("cSel").Cells.Item(ix).Specific).Checked)
                        selected = "Y";
                    docNum = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocNum").Cells.Item(ix).Specific).Value);
                    cardCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cCardCode").Cells.Item(ix).Specific).Value;
                    docType = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocType").Cells.Item(ix).Specific).Value;
                    val = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocDate").Cells.Item(ix).Specific).Value;
                    docDate = DateTime.ParseExact(val, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);
                    val = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDueDate").Cells.Item(ix).Specific).Value;
                    dueDate = DateTime.ParseExact(val, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);
                    daysPastDue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDPastDue").Cells.Item(ix).Specific).Value;
                    docTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocTotal").Cells.Item(ix).Specific).Value,CultureInfo.InvariantCulture);
                    balDue = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cBalDue").Cells.Item(ix).Specific).Value, CultureInfo.InvariantCulture);
                    discTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotDisc").Cells.Item(ix).Specific).Value, CultureInfo.InvariantCulture);
                    payTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotPay").Cells.Item(ix).Specific).Value, CultureInfo.InvariantCulture);
                    dbutil.InsertTrans(docEntry, docNum, cardCode, cardName, docType, docDate,
                                   dueDate, daysPastDue, docTotal, balDue, discTotal, payTotal, selected);


                }
            } catch (Exception e)
            {
                Logger.Log(e);
            }

        }
        private void setBindCols()
        {
            string fldname = "";
            for (int ix =1; ix<oMatrix.Columns.Count;ix++)
            {
                try
                {
                    fldname = oMatrix.Columns.Item(ix).UniqueID;
                    fldname = fldname.Substring(1, fldname.Length - 1);
                    
                    oMatrix.Columns.Item(ix).DataBind.Bind(Matrixdt.UniqueID, fldname);
                } catch (Exception e)
                {
                    Logger.Log(e);
                }
            }
        }
        private void setField(DataTable dt, int row, string fldname, object value)
        {
            try
            {
                dt.SetValue(fldname,row,value);
            } catch (Exception e)
            {
                Logger.Log(e);
            }
        }
        public void Bind(string ConsCode)
        {
            //string qry = $"EXEC rediSP_BPOpenTrans '{ConsCode}'";
            oForm.Freeze(true);
            string qry = "";
            if (ProgData.B1Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                 qry = "select 'N' as \"Sel\",\"DocNum\", \"CardCode\", \"CardName\", 'IN' as \"DocType\", \"DocDate\", \"DocDueDate\" , ";
                qry += "CASE when  DAYS_BETWEEN(CURRENT_DATE,\"DocDueDate\" ) < 0 then '*' else  CAST(DAYS_BETWEEN(CURRENT_DATE,\"DocDueDate\") as VARCHAR) ";
                qry += "end as \"DaysPastDue\", ";

                qry += "\"DocTotal\", \"DocTotal\"-\"PaidToDate\" as \"BalDue\" ";

                qry += "from \"OINV\" ";
            }
            else
                qry = $"EXEC rediSP_BPOpenTrans {ConsCode}";
             Recordset oRecSet = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecSet.DoQuery(qry);
            /*
            Matrixdt.Columns.Add("Sel", BoFieldsType.ft_AlphaNumeric, 1);
            Matrixdt.Columns.Add("CardCode", BoFieldsType.ft_AlphaNumeric, 15);
            Matrixdt.Columns.Add("DocNum", BoFieldsType.ft_Integer, 11);
            Matrixdt.Columns.Add("DocType", BoFieldsType.ft_AlphaNumeric, 1);
            Matrixdt.Columns.Add("DocDate", BoFieldsType.ft_Date);
            Matrixdt.Columns.Add("DueDate", BoFieldsType.ft_Date);
            Matrixdt.Columns.Add("DPastDue", BoFieldsType.ft_AlphaNumeric, 4);
            Matrixdt.Columns.Add("DocTotal", BoFieldsType.ft_Price);
            Matrixdt.Columns.Add("BalDue", BoFieldsType.ft_Price);
            Matrixdt.Columns.Add("TotDisc", BoFieldsType.ft_Price);
            Matrixdt.Columns.Add("TotPay", BoFieldsType.ft_Price);
            */
            //var task = Task.Run(async () => await FillMatrixdt(oRecSet));
            FillMatrixdt(oRecSet);
            oMatrix.LoadFromDataSourceEx();
            setBindCols();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            totdiscTotal=0;
            totDocTotal=0;
            totPActotal = 0;
            totpayment = 0;
            oForm.Freeze(false);

        }
        private void FillMatrixdt(Recordset oRecSet)
        {
            int ix = -1;

            try
            {
                
                    int yx = 0;
                Matrixdt.Rows.Clear();

                while (!oRecSet.EoF)
                {
                    ix++;
                    Matrixdt.Rows.Add();
                    Matrixdt.Rows.Offset = ix;
                    setField(Matrixdt, ix, "Sel", "N");
                    setField(Matrixdt, ix, "CardCode", oRecSet.Fields.Item("CardCode").Value.ToString());
                    setField(Matrixdt, ix, "DocNum", oRecSet.Fields.Item("DocNum").Value);
                    setField(Matrixdt, ix, "DocType", oRecSet.Fields.Item("DocType").Value);
                    setField(Matrixdt, ix, "DocDate", oRecSet.Fields.Item("DocDate").Value);
                    setField(Matrixdt, ix, "DueDate", oRecSet.Fields.Item("DocDueDate").Value);
                    setField(Matrixdt, ix, "DPastDue", oRecSet.Fields.Item("DaysPastDue").Value);
                    setField(Matrixdt, ix, "DocTotal", oRecSet.Fields.Item("DocTotal").Value);
                    setField(Matrixdt, ix, "TotDisc", 0);
                    setField(Matrixdt, ix, "TotPay", 0);
                    yx++;
                    if (yx>50)
                    {

                        break;
                        
                    }
                    oRecSet.MoveNext();
                }

            }
            catch (Exception e)
            {
                Logger.Log(e);
            }

            oMatrix.LoadFromDataSourceEx();
            //Task t = Task.Delay(500);
            //await t;
        }
        public void grpMatrix_LostFocusAfter(ItemEvent pVal)
        {
            string cardCode = ((EditText)ogrpMatrix.GetCellSpecific("cCodex", pVal.Row)).Value;
            double payonacctot = Convert.ToDouble(((EditText)ogrpMatrix.GetCellSpecific("cPayaccx", pVal.Row)).Value, CultureInfo.InvariantCulture);
            if (payonacctot != 0 && pVal.ColUID.CompareTo("cPayaccx") == 0)
            {
                JEPopup jepopup = new JEPopup();
                List<Payment> JEList = payments.Where(w => w.CardCode == cardCode).ToList();
                double doctotal = JEList.Sum(c => c.DocTotal);

                double rate = payonacctot / doctotal;
                int ftop = ogrpMatrix.Item.Top+(pVal.Row*20);
                jepopup.AddControls(JEList, cardCode, payonacctot, rate,ftop);
            }
        }
        public void oMatrix_LostFocusAfter(ItemEvent pVal)
        {
            string data = "";
            string fname = "";
            if (pVal.Row < 1)
                return;
            try
            {

                SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(pVal.ColUID);
                fname = oColumn.DataBind.Alias;
                double paytot = Convert.ToDouble(((EditText)oMatrix.GetCellSpecific("cTotPay", pVal.Row)).Value, CultureInfo.InvariantCulture);
                double disctot = Convert.ToDouble(((EditText)oMatrix.GetCellSpecific("cTotDisc", pVal.Row)).Value, CultureInfo.InvariantCulture);
                if ((paytot ==0 && disctot == 0) &&
                    (pVal.ColUID.CompareTo("cTotPay") == 0 || pVal.ColUID.CompareTo("cTotDisc") == 0))
                {
                   UpdateList(pVal,true);
                } else
                if (pVal.ColUID.CompareTo("cTotPay") == 0 || pVal.ColUID.CompareTo("cTotDisc") == 0 && 
                    (paytot != 0 || disctot !=0   ))
                {
                    UpdateList(pVal,false);
                }

            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + pVal.ColUID + "-//fldname=" + fname));
            }

        }
        private void UpdateList(ItemEvent pVal,bool delete)
        {
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(pVal.ColUID);
            string fname = oColumn.DataBind.Alias;
            string data = "Y";
            int color = 255 | (255 << 8) | (255 << 16);
            if (!delete)
            {
                data = "N";
                color = 252 | (221 << 8) | (130 << 16);
               
            }
            
            double discTotal = 0;
            double payTotal = 0;
            double docTotal = 0;
            string cardcode = "";
            string docType = "";
            string DocNum = "";
           
           
            try
            {
                discTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotDisc").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {

            }
            try
            {
                payTotal = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("cTotPay").Cells.Item(pVal.Row).Specific).Value, CultureInfo.InvariantCulture);
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
                DocNum= ((SAPbouiCOM.EditText)oMatrix.Columns.Item("cDocNum").Cells.Item(pVal.Row).Specific).Value;
            }
            catch (Exception e)
            {

            }
            Matrixdt.Rows.Offset = pVal.Row - 1;
            Matrixdt.SetValue("TotDisc", pVal.Row - 1, discTotal);
            Matrixdt.SetValue("TotPay", pVal.Row - 1, payTotal);
            Matrixdt.SetValue("Sel", pVal.Row - 1, data);
            try
            {
                oMatrix.CommonSetting.SetRowBackColor(pVal.Row, color);
                Payment payment = payments.Find(pym => pym.CardCode == cardcode && pym.Row == pVal.Row);

                if (delete && payments.Contains(payment))
                {
                    discTotal = payment.DiscTotal;
                    docTotal = payment.DocTotal;
                    payTotal = payment.PaymentTotal;
                    payments.Remove(payment);
                    totdiscTotal -= discTotal;
                    totDocTotal -= docTotal;
                    totPActotal -= payTotal;
                    totpayment -= payTotal;
                }
                else
                    if (!delete && payments.Contains(payment))
                {
                    double _discTotal = payment.DiscTotal;
                    double _docTotal = payment.DocTotal;
                    double _payTotal = payment.PaymentTotal;
                    totdiscTotal -= _discTotal;
                    totDocTotal -= _docTotal;
                    totPActotal -= _payTotal;
                    totpayment -= _payTotal;
                    payments.Remove(payment);
                    payments.Add(new Payment() { CardCode = cardcode, CardName = "", DocType = docType,DocNum=DocNum  ,DocTotal = docTotal, DiscTotal = discTotal, PaymentTotal = payTotal, Row = pVal.Row });

                    totdiscTotal += discTotal;
                    totDocTotal += docTotal;
                    totPActotal += payTotal;
                    totpayment += payTotal;
                }
                else
                if (!delete)
                {
                    payments.Add(new Payment() { CardCode = cardcode, CardName = "", DocType = docType, DocNum = DocNum, DocTotal = docTotal, DiscTotal = discTotal, PaymentTotal = payTotal, Row = pVal.Row });

                    totdiscTotal += discTotal;
                    totDocTotal += docTotal;
                    totPActotal += payTotal;
                    totpayment += payTotal;
                }
                UD_PaTotal.ValueEx = Convert.ToString(totPActotal, CultureInfo.InvariantCulture);
                UD_Total.ValueEx = Convert.ToString(totDocTotal, CultureInfo.InvariantCulture);
                UD_DcTotal.ValueEx = Convert.ToString(totdiscTotal, CultureInfo.InvariantCulture);
                UD_TotalPay.ValueEx = Convert.ToString(totpayment, CultureInfo.InvariantCulture);
                oMatrix.SetLineData(pVal.Row);
                oMatrix.SetLineData(pVal.Row);
                this.Listgrup();
            } catch (Exception e)
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
                   
                        if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Checked)
                    {
                        data = "N";
                        UpdateList(pVal, true);
                        
                    } else
                    {
                        data = "Y";
                        UpdateList(pVal, false);

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
                
           var groupedCustomerList = payments.GroupBy(u => u.CardCode)
                                   .Select(grp => new { Code = grp.Key , DocTotal = grp.Sum(s=>s.DocTotal), DiscTotal = grp.Sum(s=>s.DiscTotal) , PaymentTotal = grp.Sum(s => s.PaymentTotal) })
                                   .ToList();
            
            grpDt.Rows.Clear();
            ogrpMatrix.Clear();
            
            if (grpDt.Columns.Count == 0)
            {

                grpDt.Columns.Add("gCardCode", BoFieldsType.ft_AlphaNumeric, 20);
                grpDt.Columns.Add("gCardName", BoFieldsType.ft_AlphaNumeric, 60);
                grpDt.Columns.Add("gDocNum", BoFieldsType.ft_AlphaNumeric, 20);
                grpDt.Columns.Add("gTotal", BoFieldsType.ft_Price);
                grpDt.Columns.Add("gDiscount", BoFieldsType.ft_Price);
                grpDt.Columns.Add("gPayment", BoFieldsType.ft_Price);                
                grpDt.Columns.Add("gPaymonAcc", BoFieldsType.ft_Price);
                ogrpMatrix.Columns.Item("cCodex").DataBind.Bind(grpDt.UniqueID, "gCardCode");
                ogrpMatrix.Columns.Item("cNamex").DataBind.Bind(grpDt.UniqueID, "gCardName");
                ogrpMatrix.Columns.Item("cPDocNox").DataBind.Bind(grpDt.UniqueID, "gDocNum");
                ogrpMatrix.Columns.Item("cTotalx").DataBind.Bind(grpDt.UniqueID, "gTotal");
                ogrpMatrix.Columns.Item("cTotDiscx").DataBind.Bind(grpDt.UniqueID, "gDiscount");
                ogrpMatrix.Columns.Item("cPaytotx").DataBind.Bind(grpDt.UniqueID, "gPayment");
                ogrpMatrix.Columns.Item("cPayaccx").DataBind.Bind(grpDt.UniqueID, "gPaymonAcc");

            }
            SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ProgData.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
           
            for (int ix = 0;ix < groupedCustomerList.Count;ix++)
            {
                var payment =  groupedCustomerList[ix];
                grpDt.Rows.Add();
                grpDt.Rows.Offset = ix;
                grpDt.SetValue("gCardCode", ix, payment.Code);
                string qry = $"Select \"CardName\" from \"OCRD\" Where \"CardCode\" = '{payment.Code}'";
                orec.DoQuery(qry);
                try
                {
                    grpDt.SetValue("gCardName", ix, orec.Fields.Item(0).Value);
                } catch (Exception e)
                {

                }
                grpDt.SetValue("gTotal", ix, Convert.ToDouble(payment.DocTotal));
                grpDt.SetValue("gDiscount", ix, Convert.ToDouble(payment.DiscTotal));
                grpDt.SetValue("gPayment", ix, Convert.ToDouble(payment.PaymentTotal));
                grpDt.SetValue("gPaymonAcc", ix, 0);

            }
            ogrpMatrix.LoadFromDataSourceEx();
            
        }

        public void EventAll(ItemEvent pVal)
        {

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == "txtCCode" )
            {
                try
                {
                    string code = ((SAPbouiCOM.EditText)oForm.Items.Item("txtCCode").Specific).Value;
                    this.Bind(code);
                }
                catch (Exception e)
                {
                    Logger.Log(e);
                }

            }
            else

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" )
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

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.ItemUID == "fldGrpTot" )
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

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.ItemUID == "mtpreMtx" )
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
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == "mtpreMtx" )
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
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == "mtgrpMtx" )
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
