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
    public class JEPopup : Shared.IForm
    {
        protected SAPbouiCOM.Form oForm = null;
        Button JE_btn;
        int selectedRownum;
        int specrow = -1;
        SAPbouiCOM.Matrix Matrix0;
        DataTable Matrixdt;
        double DebTot = 0;
        double CrdTot = 0;
        public PrePayment caller;
        List<JE> Jelist;
        JesKey jkey;
        UserDataSource tuds1, tuds2;
        int formtop;
        bool saved,editable;
        string Cardcode;
        public void CreatenewForm()
        {
            try
            {
                SAPbouiCOM.FormCreationParams oCP = null;
                oCP = ((SAPbouiCOM.FormCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                oCP.UniqueID = "JEF" + Guid.NewGuid().ToString().Substring(0, 6);

                oCP.FormType = "JEF";
                oCP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                oCP.Modality = BoFormModality.fm_Modal;
                oForm = ProgData.B1Application.Forms.AddEx(oCP);
                ProgData.Forms.Add(oForm.UniqueID, this);
                oForm.Title = "Template of Journal Entry";
                oForm.Mode = BoFormMode.fm_OK_MODE;
                oForm.Top = formtop;
                oForm.Left = 800;
                oForm.Width = 670;
                oForm.Height = 350;
                oForm.AutoManaged = true;
                Item oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 6;
                oItem.Width = 65;
                oItem.Top = 255;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Ok";
                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = 255;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Cancel";

               /* oItem = oForm.Items.Add("Jebtn", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 145;
                oItem.Width = 165;
                oItem.Top = 255;
                oItem.Height = 19;
                oItem.Enabled = false;
                JE_btn = (SAPbouiCOM.Button)oItem.Specific;
                JE_btn.Caption = "Add Journal Entry";
               */
            }
            catch (Exception err)
            {
                Logger.Log(err);
            }


        }
        public void Save_Click(ItemEvent pVal)
        {
            // ProgData.B1Application.StatusBar.SetText("In this version it doesn't work", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            string cardcode = ((SAPbouiCOM.EditText)oForm.Items.Item("ccodex").Specific).Value;

            string val = ((SAPbouiCOM.EditText)oForm.Items.Item("pdatex").Specific).Value;
            DateTime postDate = DateTime.ParseExact(val, "yyyyMMdd",  CultureInfo.InvariantCulture);
            double Debit = 0;
                double Credit = 0;
            Jelist = new List<JE>();
            string acct = "";
            string code = "";
            int dnum = 0;
            string dtyp = "";
            string lmemo = "";
            double debit = 0;
            double credit = 0;
            Matrix0.SetCellFocus(Matrix0.RowCount, 2);
           
             for (int ix = 0;ix< Matrixdt.Rows.Count;ix++)
            {
                acct = Matrixdt.GetValue("Acct", ix).ToString();
                code = Matrixdt.GetValue("CardCode", ix).ToString();
                try
                {
                    dnum = Convert.ToInt32(Matrixdt.GetValue("DocNum", ix));
                } catch (Exception e)
                {
                    dnum = 0;
                }
                dtyp = Matrixdt.GetValue("DocType", ix).ToString();
                try
                {
                    debit = Convert.ToDouble(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture);
                }
                catch (Exception e)
                {
                    debit = 0;
                }
                try
                {
                    credit = Convert.ToDouble(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
                }
                catch (Exception e)
                {
                    credit = 0;
                }
                try
                {
                    lmemo = Matrixdt.GetValue("LMemo", ix).ToString();
                }
                catch (Exception e)
                {
                    lmemo = "";
                }
            if (String.IsNullOrEmpty(acct) ||    (debit == 0 && credit == 0))
                    continue;
                JE je = new JE();
                je.LineMemo = lmemo;
                je.LineNum = ix;
                je.PostDate = postDate;
                je.CardCode = code;
                je.Account = acct;
                je.DocType = dtyp;
                je.Docnum = dnum;
                je.Debit = debit;
                je.Credit = credit;
                je.Posted = "A";
                Debit += je.Debit;
                Credit += je.Credit;
                Jelist.Add(je);
            }
            saved = false;
            if (Debit - Credit == 0 )
                saved = true;
            else
                saved = false;
            caller.JES[jkey] = Jelist;
            if (saved)
               oForm.Mode =BoFormMode.fm_OK_MODE;
            /*      
                  DbUtility dbutil = new DbUtility();
                  Decimal Debit = 0, Credit = 0;
                  for (int ix = 0; ix < Matrixdt.Rows.Count; ix++)
                  {
                      dbutil.InsertPayAcct(0, Convert.ToInt32(Matrixdt.GetValue("DocNum", ix)), cardcode, Matrixdt.GetValue("DocType", ix).ToString(), postDate,
                             Matrixdt.GetValue("LMemo", ix).ToString(), Convert.ToDecimal(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture),
                             Convert.ToDecimal(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture), ProgData.B1Company.UserName);
                      Debit += Convert.ToDecimal(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture);
                      Credit += Convert.ToDecimal(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
                  }
                  saved = true;
                  if (Debit - Credit == 0 && saved)
                      JE_btn.Item.Enabled = true;

                  /*  for (int ix = 0; ix < Matrixdt.Rows.Count; ix++)
                    {
                        JE je = new JE();

                        je.Debit = Convert.ToDecimal(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture);
                        je.Credit = Convert.ToDecimal(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
                        je.Account = Matrixdt.GetValue("Acct", ix).ToString();
                        je.DocType = Matrixdt.GetValue("DocType", ix).ToString();
                        je.Docnum = Convert.ToInt32(Matrixdt.GetValue("DocNum", ix));
                        je.LineMemo = Matrixdt.GetValue("LMemo", ix).ToString();
                        je.LineNum = ix + 1;
                        je.CardCode = cardcode;

                    }*/


        }
        public void JeBtn_Click(ItemEvent pVal)
        {
            Jelist = new List<JE>();
            Jelist = caller.JES[jkey] ;
            for (int i = 0; i < Jelist.Count; i++)
            {
                var je = Jelist[i];
                je.Posted= "A";
                Jelist[i] = je;
            }
            caller.JES[jkey] = Jelist;
            oForm.Close();
        }
        public string AddControls(List<JE> Jes, JesKey _jkey ,int ftop,PrePayment _caller,bool _editable)
        {
            formtop = ftop;
            editable = _editable;
            int top = 25;
            caller = _caller;
            jkey = _jkey;
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
                uds1.ValueEx = jkey.CardCode;
                Cardcode = jkey.CardCode;
                label = (SAPbouiCOM.StaticText)oForm.Items.Add("Label02", SAPbouiCOM.BoFormItemTypes.it_STATIC).Specific;
                label.Item.Left = 300;
                label.Item.Top = top + 0;
                label.Item.Height = 14;
                label.Item.Width = 80;
                label.Caption = "JE Post Date";
                UserDataSource uds2 = oForm.DataSources.UserDataSources.Add("uds2", BoDataType.dt_DATE, 10);
                SAPbouiCOM.EditText pdEdt1 = (SAPbouiCOM.EditText)oForm.Items.Add("pdatex", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific;
                pdEdt1.Item.Left = 384;
                pdEdt1.Item.Top = top;
                pdEdt1.Item.Width = 100;
                pdEdt1.DataBind.SetBound(true, "", "uds2");
                pdEdt1.Item.Enabled = editable;
                uds2.ValueEx = Jes[0].PostDate.ToString("yyyyMMdd");  //  DateTime.Today.ToString("yyyyMMdd");
                Matrix0 = (SAPbouiCOM.Matrix)(oForm.Items.Add("MxJE01", SAPbouiCOM.BoFormItemTypes.it_MATRIX).Specific);

                Matrix0.Item.Width = 650;
                Matrix0.Item.Height = 200;
                Matrix0.Item.Top = top + 25;
                Matrix0.Item.Left = 16;
                Matrix0.Item.Enabled = editable;
                // oDBDataSource = oForm.DataSources.DBDataSources.Add("@LSBNKLOG");
                Matrixdt = oForm.DataSources.DataTables.Add("JEPre");

                Matrix0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                Matrixdt.Columns.Add("Acct", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("DocType", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("DocNum", BoFieldsType.ft_AlphaNumeric, 20);
                Matrixdt.Columns.Add("LMemo", BoFieldsType.ft_AlphaNumeric, 100);
                Matrixdt.Columns.Add("Debit", BoFieldsType.ft_Sum);
                Matrixdt.Columns.Add("Credit", BoFieldsType.ft_Sum);
                Matrixdt.Columns.Add("CardCode", BoFieldsType.ft_AlphaNumeric, 20);
                this.AddChooseFromList("1","CFL1", "Postable", "Y");
                this.AddChooseFromList("2", "CFL2", "CardType", "C");
                Column col = Matrix0.Columns.Add("00", BoFormItemTypes.it_EDIT);

                col.TitleObject.Caption = "#";
                col.Editable = false;
                col.Width = 10;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("Acct", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Acct");
                col.TitleObject.Caption = "Account";
                col.Editable = true;
                col.Width = 100;
                col.DisplayDesc = false;
                col.Editable = editable;
                col.ChooseFromListUID = "CFL1";
                col.ChooseFromListAlias = "AcctCode";
                //   (col as EditTextColumn).LinkedObjectType = "1";
                col = Matrix0.Columns.Add("DocType", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "DocType");
                col.TitleObject.Caption = "Doc.Type";
                
                col.Width = 80;
                col.Editable = editable;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("DocNum", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "DocNum");
                col.TitleObject.Caption = "Doc.Num";
                col.Editable = editable;
                col.Width = 100;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("LMemo", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "LMemo");
                col.TitleObject.Caption = "Reference";
                col.Editable = editable;
                col.Width = 100;
                col.DisplayDesc = false;
                col = Matrix0.Columns.Add("Debit", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Debit");
                col.TitleObject.Caption = "Debit Amount";
                col.Editable = editable;
                col.RightJustified = true;
                col.Width = 120;
                col.DisplayDesc = false;
                col.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                col = Matrix0.Columns.Add("Credit", BoFormItemTypes.it_EDIT);
                col.DataBind.Bind("JEPre", "Credit");
                col.TitleObject.Caption = "Credit Amount";
                col.RightJustified = true;
                col.Editable = editable;
                col.Width = 120;
                col.DisplayDesc = false;
                col.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                int ix = 0;
                DebTot = 0;
                CrdTot = 0;
                for (ix = 0; ix < Jes.Count; ix++)
                {
                    Matrixdt.Rows.Add();
                    Matrixdt.Rows.Offset = ix;
                    
                    Matrixdt.SetValue("Acct", ix, Jes[ix].Account);
                    Matrixdt.SetValue("DocType", ix, Jes[ix].DocType);
                    Matrixdt.SetValue("DocNum", ix, Jes[ix].Docnum);
                    Matrixdt.SetValue("LMemo", ix, Jes[ix].LineMemo);
                    Matrixdt.SetValue("Debit", ix, Jes[ix].Debit);
                    Matrixdt.SetValue("Credit", ix, Jes[ix].Credit);
                    Matrixdt.SetValue("CardCode", ix, Jes[ix].CardCode);
                    if (Jes[ix].Account == Jes[ix].CardCode)
                        specrow = ix;
                    DebTot += Jes[ix].Debit;
                    CrdTot += Jes[ix].Debit;
                }
               
                addRow();
                Matrix0.LoadFromDataSource();
                if (Matrixdt.GetValue("Acct",0).ToString() =="")
                    Matrix0.CommonSetting.SetCellEditable(1, 1, true); else
                Matrix0.CommonSetting.SetCellEditable(1, 1, false);
                Matrix0.CommonSetting.SetCellEditable(1, 2, true);
                Matrix0.CommonSetting.SetCellEditable(1, 3, true);
                Matrix0.CommonSetting.SetCellEditable(1, 4, true);
                Matrix0.CommonSetting.SetCellEditable(1, 5, false);
                Matrix0.CommonSetting.SetCellEditable(1, 6, false);
                Matrix0.CommonSetting.SetCellEditable(specrow + 1, 1, false);

                //  Matrix0.CommonSetting.SetCellEditable(ix + 1, 1, true);
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
        private void addRow(int col = 6)
        {
            int ix = Matrixdt.Rows.Count - 1;
            int lix = ix;
            lix++;
            double debit = 0;
            double credit = 0;
            if (ix>-1)
            {
                debit = Convert.ToDouble(Matrixdt.GetValue("Debit", ix), CultureInfo.InvariantCulture);
                credit = Convert.ToDouble(Matrixdt.GetValue("Credit", ix), CultureInfo.InvariantCulture);
            }
            if ((DebTot-CrdTot)!=0 && (debit - credit) !=0 )
            {
                Matrixdt.Rows.Add();

                ix  = Matrixdt.Rows.Count -1;
                Matrixdt.SetValue("Acct", ix, "");
                Matrixdt.SetValue("DocType", ix, "");
                Matrixdt.SetValue("DocNum", ix, "");
                Matrixdt.SetValue("Debit", ix, 0);
                Matrixdt.SetValue("Credit", ix, 0);
                Matrixdt.SetValue("CardCode", ix, "");
                Matrix0.LoadFromDataSource();
                col++;
                if (col > 6)
                {
                    lix++;
                    col = 1;
                }
                Matrix0.SetCellFocus(lix , col);
                Matrix0.SetCellFocus(lix, col);
            }
            
        }
        private void EditText_ChooseFromListAfter(ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;


            try
            {
                oForm.Select();
                oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                SAPbouiCOM.EditText edt = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                string sCFL_ID = null;
                sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.DataTable oDataTable = null;
                oDataTable = oCFLEvento.SelectedObjects;
                string itemCode = null;

                itemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));
                //   SAPbouiCOM.EditText EditText = (SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific;
                edt.Value = itemCode;
            }
            catch (Exception e)
            {
                Logger.Log(e);

            }
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
                string acccode = "";
                try
                { 
                 acccode = oDataTable.GetValue("Segment_0", 0).ToString() + '-' + oDataTable.GetValue("Segment_1", 0).ToString() + '-' + oDataTable.GetValue("Segment_2", 0).ToString();
                } catch (Exception e)
                {
                    acccode = "";
                }
                string itemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));

                Matrixdt.Rows.Offset = pVal.Row - 1;
                if (String.IsNullOrEmpty(acccode))
                {
                    Matrixdt.SetValue(fname, Matrixdt.Rows.Offset, acccode);
                } else
                Matrixdt.SetValue(fname, Matrixdt.Rows.Offset, itemCode);
                Matrix0.SetLineData(pVal.Row);

            }
            catch (Exception er)
            {

            }


        }
        private void Matrix0_GotFocusAfter(ItemEvent pVal)
        {


            selectedRownum = pVal.Row;  

             if (pVal.ColUID == "Acct")
            {
                oForm.Freeze(true);
                Matrixdt.Rows.Offset = pVal.Row - 1;
                if (String.IsNullOrEmpty(Matrixdt.GetValue("Acct", Matrixdt.Rows.Offset).ToString()))
                {
                   
                    Matrixdt.SetValue("Acct", Matrixdt.Rows.Offset, Cardcode);
                }
               
                if (!String.IsNullOrEmpty(Matrixdt.GetValue("CardCode",Matrixdt.Rows.Offset).ToString()))
                {
                    Column col = Matrix0.Columns.Item(pVal.ColUID);
                    col.ChooseFromListUID = "CFL2";
                    col.ChooseFromListAlias = "CardCode";

                }
                else
                {
                    Column col = Matrix0.Columns.Item(pVal.ColUID);
                    col.ChooseFromListUID = "CFL1";
                    col.ChooseFromListAlias = "AcctCode";
                }
                Matrix0.LoadFromDataSourceEx();
                oForm.Freeze(false);

            }

        }
        private void Matrix0_LostFocusAfter(ItemEvent pVal)
        {

            string fname = "";
            if (pVal.Row < 1)
                return;
            if (editable)
                oForm.Mode = BoFormMode.fm_UPDATE_MODE;
            else
                return;
            try
            {

                SAPbouiCOM.Column oColumn = Matrix0.Columns.Item(pVal.ColUID);
                fname = oColumn.UniqueID;
                Column col = Matrix0.Columns.Item(pVal.ColUID);
                if (fname == "Acct" && oColumn.ChooseFromListUID == "CFL2")
                {
                    string ccode = ((EditText)col.Cells.Item(pVal.Row).Specific).Value;

                    Matrixdt.SetValue("CardCode", pVal.Row - 1, ccode);
                    Matrixdt.SetValue(fname, pVal.Row - 1, ((EditText)col.Cells.Item(pVal.Row).Specific).Value);
                }
                else
                if (fname == "Acct")
                {
                    Matrixdt.SetValue("CardCode", pVal.Row - 1, "");
                    Matrixdt.SetValue(fname, pVal.Row - 1, ((EditText)col.Cells.Item(pVal.Row).Specific).Value);
                }
                else
                 if (fname == "LMemo" || fname == "DocNum" || fname == "DocType")
                {
                    Matrixdt.SetValue(fname, pVal.Row - 1, ((EditText)col.Cells.Item(pVal.Row).Specific).Value);
                } else 
                 if (fname == "Debit" || fname == "Credit")
                {
                    double tutar = 0;
                    double tutarx = 0;
                    tutarx = Convert.ToDouble(Matrixdt.GetValue(fname, pVal.Row - 1), CultureInfo.InvariantCulture);
                    tutar = Convert.ToDouble(((EditText)Matrix0.GetCellSpecific(fname, pVal.Row)).Value, CultureInfo.InvariantCulture);
                    if (tutar - tutarx != 0)
                        saved = false;
                    Matrixdt.SetValue(fname, pVal.Row - 1, tutar);
                    if (fname == "Debit" && tutar != 0)
                    {
                        Matrixdt.SetValue("Credit", pVal.Row - 1, 0);

                    }
                    else
                         if (fname == "Credit" && tutar != 0)
                    {
                        Matrixdt.SetValue("Debit", pVal.Row - 1, 0);

                    }
                
                  //  Matrix0.LoadFromDataSourceEx();
                   
                    DebTot = Convert.ToDouble(Matrix0.Columns.Item("Debit").ColumnSetting.SumValue, CultureInfo.InvariantCulture);
                    CrdTot = Convert.ToDouble(Matrix0.Columns.Item("Credit").ColumnSetting.SumValue, CultureInfo.InvariantCulture);

                    if (DebTot - CrdTot == 0 && saved && !String.IsNullOrEmpty(((EditText)Matrix0.GetCellSpecific(fname, pVal.Row)).Value))
                    {
                        //  JE_btn.Item.Enabled = true;
                    }
                    else
                    {
//                        JE_btn.Item.Enabled = false;
                        int cl = 5;
                        if (fname == "Debit")
                            cl = 5;
                        else
                            cl = 6;
                        addRow(cl);
                    }
                    //    Matrix0.SetCellFocus(pVal.Row + 1, 1);
                } 
                /*else
                {
                    Matrixdt.SetValue(fname, pVal.Row - 1, ((EditText)col.Cells.Item(pVal.Row).Specific).Value);
                }*/
                
                 Matrix0.LoadFromDataSourceEx();
                Matrixdt.Rows.Offset = pVal.Row - 1;
                oForm.Freeze(false);

            }
            catch (Exception e)
            {
                Logger.Log(new Exception(e.Message + pVal.ColUID + "-//fldname=" + fname));
            }

        }

        private void AddChooseFromList(string objtyp,string uID,string alias,string value)
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
                oCFLCreationParams.ObjectType = objtyp;
                oCFLCreationParams.UniqueID = uID;

                oCFL = oCFLs.Add(oCFLCreationParams);

                //  Adding Conditions to CFL1

                oCons = oCFL.GetConditions();

                oCon = oCons.Add();
                oCon.Alias = alias;
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = value;
                oCFL.SetConditions(oCons);



            }
            catch
            {

            }
        }

        public void EventAll(ItemEvent pVal)
        {
            if (pVal.BeforeAction == true)
            {
                if (pVal.EventType == BoEventTypes.et_KEY_DOWN)
                {
                    if (pVal.ColUID == "Acct" && pVal.Modifiers == BoModifiersEnum.mt_CTRL && pVal.CharPressed == 9)
                    {
                        Column col = Matrix0.Columns.Item(pVal.ColUID);
                        col.ChooseFromListUID = "CFL2";
                        col.ChooseFromListAlias = "CardCode";

                    }
                    else
                    if (pVal.ColUID == "Acct")
                    {
                        Column col = Matrix0.Columns.Item(pVal.ColUID);
                        col.ChooseFromListUID = "CFL1";
                        col.ChooseFromListAlias = "AcctCode";
                    }

                }

            } else
            if (pVal.BeforeAction == false)
            {
                
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_UPDATE_MODE)
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
                 if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_OK_MODE)
                {
                    try
                    {
                        oForm.Close();
                    }
                    catch (Exception e)
                    {
                        Logger.Log(e);
                    }

                }
                else
               if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "Jebtn" && oForm.Items.Item("Jebtn").Enabled)
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
                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "MxJE01")
                {
                    this.Matrix0_ChooseFromListAfter(pVal);
                }
                else
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "MxJE01")
                {
                    this.Matrix0_ChooseFromListAfter(pVal);
                   
                }
                else

                 if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS || pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE) && pVal.ItemUID == "MxJE01")
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

        public void Menu(ref MenuEvent pVal)
        {
            if (pVal.BeforeAction == false)
            {
                if (pVal.MenuUID == "linedel")
                {

                    if (Matrix0.RowCount != 1)
                    {
                        if (specrow + 1 >= selectedRownum)
                        {
                            Matrix0.CommonSetting.SetCellEditable(specrow + 1, 1, true);
                            specrow--;
                            if (specrow < 0)
                                specrow = 0;
                            else
                                Matrix0.CommonSetting.SetCellEditable(specrow + 1, 1, false);
                        }
                        Matrix0.DeleteRow(selectedRownum);
                        Matrixdt.Rows.Remove(selectedRownum - 1);
                        Matrix0.FlushToDataSource();



                    }
                }
            }
        }
    }

}
