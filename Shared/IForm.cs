﻿using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared
{
    public interface IForm
    {
        void EventAll(ItemEvent pVal);
        void Menu(ref SAPbouiCOM.MenuEvent pVal);
        
    }
}