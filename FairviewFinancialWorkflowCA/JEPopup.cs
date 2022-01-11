using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialWorkflowCA
{
    public class JEPopup:IForm
    {
        protected SAPbouiCOM.Form oForm = null;
        SAPbouiCOM.Matrix Matrix0;
        DataTable Matrixdt;
        double DebTot = 0;
        double CrdTot = 0;
        public PrePayment caller;
        UserDataSource tuds1, tuds2;
        int formtop;
        public  void CreatenewForm()
        {
            try
            {
                SAPbouiCOM.FormCreationParams oCP = null;
                oCP = ((SAPbouiCOM.FormCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
               
                oCP.UniqueID = "JEF" + Guid.NewGuid().ToString().Substring(0, 6);

                oCP.FormType = "JEF";
                oCP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

                oForm = ProgData.B1Application.Forms.AddEx(oCP);
                B1Starter.Forms.Add(oForm.UniqueID, this);
                oForm.Title = "Template of Journal Entry";
                oForm.Top = formtop;
                oForm.Left = 800;
                oForm.Width = 670;
                oForm.Height = 250;
                oForm.AutoManaged = true;
                Item oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 6;
                oItem.Width = 65;
                oItem.Top = 175;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Save";
                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = 175;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Cancel";

                oItem = oForm.Items.Add("Jebtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 165;
                oItem.Top = 175;
                oItem.Height = 19;
                oItem.Enabled = false;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Add Journal Entry";


            }
            catch (Exception err)
            {
                Logger.Log(err);
            }


        }
        public void Save_Click(ItemEvent pVal)
        {
            if ( DebTot - CrdTot ==  0)
            {

                ProgData.B1Application.StatusBar.SetText("In this version it doesn't work", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                string cardcode = ((SAPbouiCOM.EditText)oForm.Items.Add("ccodex", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific).Value;
                
                string val = ((SAPbouiCOM.EditText)oForm.Items.Item("pdatex").Specific).Value;
                DateTime refDate = DateTime.ParseExact(val, "yyyyMMdd",
                                 CultureInfo.InvariantCulture);
                for (int ix = 0; ix < Matrixdt.Rows.Count; ix++)
                {
                    JE je = new JE();
                    
                    je.Debit = Convert.ToDecimal(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture);
                    je.Credit = Convert.ToDecimal(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
                    je.Account = Matrixdt.GetValue("Acct", ix).ToString();
                    je.DocType = Matrixdt.GetValue("DocType", ix).ToString();
                    je.Docnum = Convert.ToInt32( Matrixdt.GetValue("DocNum", ix));
                    je.LineMemo = Matrixdt.GetValue("LMemo", ix).ToString();
                    je.LineNum = ix+1;
                    je.CardCode = cardcode;
                   
                }
            }
            
        }
        public void JeBtn_Click(ItemEvent pVal)
        {

        }
        public  string AddControls(List<Payment> payments,string CardCode,double PayonAcc,double rate,int ftop)
        {
            formtop = ftop;
            int top = 25;

            try
            {
                if (oForm == null)
                    CreatenewForm();
                oForm.Freeze(true);


                SAPbouiCOM.StaticText label = (SAPbouiCOM.StaticText)oForm.Items.Add("Label01", SAPbouiCOM.BoFormItemTypes.it_STATIC).Specific;
                label.Item.Left = 2;
                label.Item.Top = top + 0;
                label.Item.Height = 14;
                label.Item.Width = 80;
                label.Caption = "CardCode";
                UserDataSource uds1 = oForm.DataSources.UserDataSources.Add("uds1", BoDataType.dt_SHORT_TEXT, 20);
                SAPbouiCOM.EditText crdEdt1 = (SAPbouiCOM.EditText)oForm.Items.Add("ccodex", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific;
                crdEdt1.Item.Left = 84;
                crdEdt1.Item.Top = top;
                crdEdt1.Item.Width = 100;
                crdEdt1.DataBind.SetBound(true, "", "uds1");
                crdEdt1.Item.Enabled = false;
                uds1.ValueEx = CardCode;
                label = (SAPbouiCOM.StaticText)oForm.Items.Add("Label02", SAPbouiCOM.BoFormItemTypes.it_STATIC).Specific;
                label.Item.Left = 300;
                label.Item.Top = top + 0;
                label.Item.Height = 14;
                label.Item.Width = 80;
                label.Caption = "JE Post Date";
                UserDataSource uds2 = oForm.DataSources.UserDataSources.Add("uds2", BoDataType.dt_DATE,10);
                SAPbouiCOM.EditText pdEdt1 = (SAPbouiCOM.EditText)oForm.Items.Add("pdatex", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific;
                pdEdt1.Item.Left = 384;
                pdEdt1.Item.Top = top;
                pdEdt1.Item.Width = 100;
                pdEdt1.DataBind.SetBound(true, "", "uds2");
                uds2.ValueEx = DateTime.Today.ToString("yyyyMMdd");
                Matrix0 = (SAPbouiCOM.Matrix)(oForm.Items.Add("MxJE01", SAPbouiCOM.BoFormItemTypes.it_MATRIX).Specific);

                Matrix0.Item.Width = 650;
                Matrix0.Item.Height = 100;
                Matrix0.Item.Top = top+25;
                Matrix0.Item.Left = 16;
                // oDBDataSource = oForm.DataSources.DBDataSources.Add("@LSBNKLOG");
                Matrixdt = oForm.DataSources.DataTables.Add("JEPre");

                Matrix0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                Matrixdt.Columns.Add("Acct", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("DocType", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("DocNum", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("LMemo", BoFieldsType.ft_AlphaNumeric, 100);
                Matrixdt.Columns.Add("Debit", BoFieldsType.ft_Float);
                Matrixdt.Columns.Add("Credit", BoFieldsType.ft_Float);
                this.AddChooseFromList();
                Column col = Matrix0.Columns.Add("00", BoFormItemTypes.it_EDIT);
              
                col.TitleObject.Caption = "#";
                col.Editable = false;
                col.Width = 10;
                col.DisplayDesc = false;
                 col = Matrix0.Columns.Add("Acct", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Acct");
                col.TitleObject.Caption = "Account";
                col.Editable = false;
                col.Width = 100; 
                col.DisplayDesc = false;
                col.ChooseFromListUID = "CFL1";
                col.ChooseFromListAlias = "AcctCode";
                //   (col as EditTextColumn).LinkedObjectType = "1";
                col = Matrix0.Columns.Add("DocType", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "DocType");
                col.TitleObject.Caption = "Doc.Type";
                col.Editable = false;
                col.Width = 80;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("DocNum", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "DocNum");
                col.TitleObject.Caption = "Doc.Num";
                col.Editable = false;
                col.Width = 100;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("LMemo", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "LMemo");
                col.TitleObject.Caption = "Reference";
                col.Editable = true;
                col.Width = 100;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("Debit", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Debit");
                col.TitleObject.Caption = "Debit Amount";
                col.Editable = true;
                col.Width = 120;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("Credit", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Credit");
                col.TitleObject.Caption = "Credit Amount";
                col.Editable = true;
                col.Width = 120;
                col.DisplayDesc = false;
                int ix = -1;
                double topbalance=0.0;
                double calc = 0.0;
                tuds1 = oForm.DataSources.UserDataSources.Add("tuds1", BoDataType.dt_PRICE);
                SAPbouiCOM.EditText topEdt1 = (SAPbouiCOM.EditText)oForm.Items.Add("topEdt1", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific;
                topEdt1.Item.Left = 406;
                topEdt1.Item.Top = top+130;
                topEdt1.Item.Width = 120;
                topEdt1.DataBind.SetBound(true, "", "tuds1");
                topEdt1.Item.Enabled = false;
                tuds2 = oForm.DataSources.UserDataSources.Add("tuds2", BoDataType.dt_PRICE);
                topEdt1 = (SAPbouiCOM.EditText)oForm.Items.Add("topEdt2", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific;
                topEdt1.Item.Left = 530;
                topEdt1.Item.Top = top + 130;
                topEdt1.Item.Width = 120;
                topEdt1.DataBind.SetBound(true, "", "tuds2");
                topEdt1.Item.Enabled = false;

                foreach (var item in payments)
                {
                    ix++;
                    Matrixdt.Rows.Add();
                    Matrixdt.Rows.Offset = ix;
                    Matrixdt.SetValue("Acct", ix, item.CardCode);
                    Matrixdt.SetValue("DocType", ix, item.DocType);
                    Matrixdt.SetValue("DocNum", ix, item.DocNum);
                    Matrixdt.SetValue("Debit", ix, 0);
                    calc = item.DocTotal * rate;
                    Matrixdt.SetValue("Credit", ix, calc);
                    topbalance += calc;
                }
                Matrixdt.Rows.Add();
                ix++;
                Matrixdt.Rows.Offset = ix;
                Matrixdt.SetValue("Acct", ix, "10110-15-200");
                Matrixdt.SetValue("DocType", ix, "");
                Matrixdt.SetValue("DocNum", ix,"");
                Matrixdt.SetValue("Debit", ix, topbalance);
                Matrixdt.SetValue("Credit", ix, 0);
                Matrix0.LoadFromDataSource();
              
                Matrix0.CommonSetting.SetCellEditable(ix+1, 1, true);
            }
            catch (Exception err)
            {

                Logger.Log(err);
            }
            
            
            oForm.PaneLevel = 1;
            oForm.Freeze(false);
            oForm.Visible = true;
            return oForm.UniqueID;
        }
        private void Matrix0_ChooseFromListAfter(ItemEvent pVal)
        {
            try
            {

                SAPbouiCOM.Column oColumn = Matrix0.Columns.Item(pVal.ColUID);
                string fname = pVal.ColUID;
                SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
               
                SAPbouiCOM.DataTable oDataTable = null;
                oDataTable = oCFLEvento.SelectedObjects;
                string itemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));

              
                Matrix0.GetLineData(pVal.Row);
                Matrixdt.SetValue(fname, Matrixdt.Rows.Offset, itemCode);
                Matrix0.SetLineData(pVal.Row);

            }
            catch (Exception er)
            {

            }


        }
        private void Matrix0_GotFocusAfter(ItemEvent pVal)
        {
            
            if (Matrix0.RowCount == Matrixdt.Rows.Count && pVal.Row == Matrix0.RowCount)
            {
                Matrixdt.Rows.Add();
            }
        }
         private void Matrix0_LostFocusAfter(ItemEvent pVal)
        {
            string data = "";
            string fname = "";
            if (pVal.Row < 1)
                return;
            try
            {

                SAPbouiCOM.Column oColumn = Matrix0.Columns.Item(pVal.ColUID);
                fname = oColumn.UniqueID;
                if (fname == "Debit" || fname == "Credit")
                {
                    double tutar = 0;
                    
                    tutar =Convert.ToDouble( ((EditText) Matrix0.GetCellSpecific(fname, pVal.Row)).Value, CultureInfo.InvariantCulture);
                    Matrixdt.SetValue(fname, pVal.Row - 1,tutar);
                    DebTot = 0;
                    CrdTot = 0;
                    for (int ix =0;ix<Matrixdt.Rows.Count;ix++)
                    {
                        DebTot += Convert.ToDouble(Matrixdt.GetValue("Debit", ix),CultureInfo.InvariantCulture);
                        CrdTot += Convert.ToDouble(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
                    }
                    tuds1.ValueEx = Convert.ToString(DebTot);
                    tuds2.ValueEx = Convert.ToString(CrdTot);
                }
            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + pVal.ColUID + "-//fldname=" + fname));
            }

        }

        private void AddChooseFromList()
        {
            try
            {

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "1";
                oCFLCreationParams.UniqueID = "CFL1";

                oCFL = oCFLs.Add(oCFLCreationParams);

                //  Adding Conditions to CFL1

                oCons = oCFL.GetConditions();

                oCon = oCons.Add();
                oCon.Alias = "Postable";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCFL.SetConditions(oCons);

              

            }
            catch
            {
                
            }
        }

        public void EventAll(ItemEvent pVal)
        {

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1")
            {
                try
                {
                    this.Save_Click(pVal);
                }
                catch (Exception e)
                {
                    Logger.Log(e);
                }

            }
            else
               if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "JeBtn")
            {
                try
                {
                    this.JeBtn_Click(pVal);
                }
                catch (Exception e)
                {
                    Logger.Log(e);
                }

            }
            else
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.ItemUID == "MxJE01")
            {
                try
                {
                    this.Matrix0_GotFocusAfter(pVal);
                }
                catch (Exception e)
                {
                    Logger.Log(e);
                }

            }
            else
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    this.Matrix0_ChooseFromListAfter(pVal);
                }
             else
             if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.ItemUID == "MxJE01")
             {
                 try
                 {
                     this.Matrix0_LostFocusAfter(pVal);
                 }
                 catch (Exception e)
                 {
                     Logger.Log(e);
                 }

             }

        }
    }
}
