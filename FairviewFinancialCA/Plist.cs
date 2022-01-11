using SAPbouiCOM;
using Shared;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialCA
{
    internal class Plist : Shared.IForm
    {
        protected SAPbouiCOM.Form oForm = null;
        Grid Grid0;
        DataTable dt;
        public string retvalue;
        public event EventHandler<DataTable> ListClosed;
        public void EventAll(ItemEvent pVal)
        {
            if (pVal.ItemUID == "grd0" && pVal.EventType == BoEventTypes.et_LOST_FOCUS)
            {
                Grid0_Lost(pVal);
            } else
            if (pVal.ItemUID == "grd0" && pVal.EventType == BoEventTypes.et_DOUBLE_CLICK)
            {
                Grid0_DoubleClickAfter(pVal);
            }
            else
             if (pVal.ItemUID == "s" && pVal.EventType == BoEventTypes.et_CLICK)
            {
                OButton_ClickAfter(pVal);
            }
        }

        public void Menu(ref MenuEvent pVal)
        {
           
        }
        private void OButton_ClickAfter(ItemEvent pVal)
        {
            try
            {
               
                OnListClosed(Grid0.DataTable);

                oForm.Close();
            }
            catch (Exception er)
            {

            }
        }
        protected virtual void OnListClosed(DataTable dt)
        {
            ListClosed?.Invoke(this, dt);
        }
        public void CreatenewForm(int left,int tp)
        {
            try
            {
                SAPbouiCOM.FormCreationParams oCP = null;
                oCP = ((SAPbouiCOM.FormCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                oCP.UniqueID = "pl" + Guid.NewGuid().ToString().Substring(0, 6);

                oCP.FormType = "pl";
                oCP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                oCP.Modality = BoFormModality.fm_Modal;
                oForm = ProgData.B1Application.Forms.AddEx(oCP);
                ProgData.Forms.Add(oForm.UniqueID, this);  
                oForm.Title = "List of Process Payments";
                oForm.Top = tp;
                oForm.Left = left;
                oForm.Width = 370;
                oForm.Height = 340;
                oForm.AutoManaged = true;
                
                oForm.Freeze(true);
                /*Item oItem = oForm.Items.Add("s", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 6;
                oItem.Width = 65;
                oItem.Top = 280;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Select";*/
                Item oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 6;
                oItem.Width = 65;
                oItem.Top = 280;
                oItem.Height = 19;
                ((SAPbouiCOM.Button)oItem.Specific).Caption = "Cancel";
                oForm.Freeze(true);
                int top = 15;
                Grid0 = (SAPbouiCOM.Grid)(oForm.Items.Add("grd0", SAPbouiCOM.BoFormItemTypes.it_GRID).Specific);
                Grid0.Item.FromPane = 0;
                Grid0.Item.ToPane = 0;
                Grid0.Item.Width = 350;
                Grid0.Item.Height = 260;

                Grid0.Item.Top = top;
                Grid0.Item.Left = 1;


                Grid0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                dt = oForm.DataSources.DataTables.Add("MTable");
                string qry  = "Select  T0.CardCode ,T0.DocDate ,T0.DocDueDate as PostDate ,"; ;
                qry += "T0.DocNum ,T0.DocRef ,T0.DocEntry ";
                qry += " from redi_ConsBP_Header T0 ";
                qry += " Order By T0.DocEntry ";
                dt.ExecuteQuery(qry);
                Grid0.DataTable = oForm.DataSources.DataTables.Item("MTable");
                SAPbouiCOM.GridColumn oColumn = Grid0.Columns.Item(0);
                // oColumn.Visible = false;
                oColumn.Editable = false;
                oColumn = Grid0.Columns.Item(1);
                oColumn.Editable = false;
                oColumn.TitleObject.Sortable = true;
                oColumn.Editable = false;
                oColumn = Grid0.Columns.Item(2);
                oColumn.Editable = false;
                oColumn.TitleObject.Sortable = true;
                oColumn = Grid0.Columns.Item(3);
                oColumn.Editable = false;
                oColumn.TitleObject.Sortable = true;
                oColumn = Grid0.Columns.Item(4);
                oColumn.Editable = false;
                oColumn.Width = 100;
                oColumn.TitleObject.Sortable = true;
                oColumn = Grid0.Columns.Item(5);
                oColumn.Editable = false;
                oColumn.Visible = false;
                oForm.PaneLevel = 1;
                oForm.Freeze(false);
                oForm.Visible = true;
            }
            catch (Exception err)
            {
                Logger.Log(err);
            }


        }
        private void Grid0_Lost(ItemEvent pVal)
        {
            Grid0.DataTable.Rows.Offset = Grid0.GetDataTableRowIndex(pVal.Row);
        }
        private void Grid0_DoubleClickAfter(ItemEvent pVal)
        {
            try
            {
                Grid0.DataTable.Rows.Offset = Grid0.GetDataTableRowIndex(pVal.Row);
                OnListClosed(Grid0.DataTable);

                oForm.Close();
            }
            catch (Exception er)
            {

            }
        }
    }
}
