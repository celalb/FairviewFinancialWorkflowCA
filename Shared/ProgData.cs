using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public class ProgData
    {
        public static SAPbobsCOM.Company B1Company = null;
        public static SAPbouiCOM.Application B1Application = null;
        public static string sqlConnectionString;
        public static SAPbouiCOM.Application oAppl { get; set; }

        public static Dictionary<String, IForm> Forms;



    }

}
