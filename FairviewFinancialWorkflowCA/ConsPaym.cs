using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FairviewFinancialWorkflowCA
{
    public class ConsPaym
    {
        protected SAPbouiCOM.Form oForm = null;
        string FormUID;
        SAPbouiCOM.Matrix ogrpMatrix, oMatrix;
        DataTable Matrixdt, grpDt;
        public ConsPaym()
        {
            try
            {
                SAPbouiCOM.FormCreationParams oCP = null;
                oCP = ((SAPbouiCOM.FormCreationParams)(ProgData.B1Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
               
                oCP.UniqueID = "Cp" + Guid.NewGuid().ToString().Substring(0, 6); ;

                oCP.FormType = "Cp";
                oCP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;

                oForm = ProgData.B1Application.Forms.AddEx(oCP);

                oForm.Title = "Consolidates Payments";
                oForm.Top = 30;
                oForm.Left = 400;
                oForm.Width = 961;
                oForm.Height = 512;
                oForm.AutoManaged = true;


                Item oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = 420;
                oItem.Height = 19;

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 80;
                oItem.Width = 85;
                oItem.Top = 420;
                oItem.Height = 19;

                
            }
            catch (Exception err)
            {
                Logger.Log(err);
            }

        }
    }
}
