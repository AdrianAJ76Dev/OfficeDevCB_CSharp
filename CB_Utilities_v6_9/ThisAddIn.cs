﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace CB_Utilities_v6_9
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Find the global add-in and load for AutoText
            /* Work
            string templatefullname = 
                @"\\nyodska01\cbwide\RAS Contracts Management\Training Documents\CM Utilities v61.dotm";
            */

            /* Home
            string templatefullname =
                @"D:\Dev Projects\MS Office Development\CM Utilities v61.dotm";
            */

            /* Ran this once and after running, it seems to autoload the add-in.
             * Will need this to have access to AutoText
            */
            string templatefullname =  @"\\nyodska01\cbwide\RAS Contracts Management\Training Documents\CM Utilities v62.dotx";
            Globals.ThisAddIn.Application.AddIns[templatefullname].Installed = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CM_Utilities_Ribbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
