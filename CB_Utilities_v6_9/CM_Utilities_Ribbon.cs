﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new CM_Utilities_Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace CB_Utilities_v6_9
{
    [ComVisible(true)]
    public class CM_Utilities_Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public CM_Utilities_Ribbon()
        {

        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CB_Utilities_v6_9.CM_Utilities_Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        /* Create callback methods here. For more information about adding callback methods, 
         * visit https://go.microsoft.com/fwlink/?LinkID=271226
         */

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void Clean_Up_Riders_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.RemoveUnnecssaryRiders();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,"Remove Unnecessary Riders",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        public void CreateSoleSourceLetter_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Create Sole Source Letter", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void MakeHEDAmendment_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.MakeHEDAmendment();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Make HED Amendment", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatPrice_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.FormatPrice();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Format Price", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatDateSpellOutMonth_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.SpellOutMonth();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Spell Out Month", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatPhoneNumber_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.FormatPhoneNumber();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Format Phone Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatCommonwealth_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.FormatCommonwealth();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Format Commonwealth", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InterfaceForSpellNumber_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                    CleanUpUtilities.SpellOutNumber();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Spell Out Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveTerDAtesFromFeeSchedule_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            try
            {
                CleanUpUtilities.RemoveTerDatesFromFeeSchedule();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Remove Term Dates From Fee Schedule", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool GetEnabled(Office.IRibbonControl rbnCtrl)
        {
            return true;
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
