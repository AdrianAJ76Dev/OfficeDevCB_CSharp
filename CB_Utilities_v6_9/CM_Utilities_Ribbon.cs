using System;
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
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

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
                throw e;
            }
        }
        public void CreateSoleSourceLetter_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            // I think I'm going to try and do the "Web Version of the Sole Source Letter for this one.
        }
        public void MakeHEDAmendment_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            DialogResult lngResult;
            long lngPageNumberSignaturePage;
            Word.Selection sel;
            Word.AutoTextEntry hedaddendum;
            Word.Template tmpl;

            const string strAUTOTEXT_AMENDMENT = "HSA - HED Standard Addendum";
            const int SIGNATURE_PAGE_AMENDMENT = 2;

            try
            {
                string msg = "This deletes pages in the main part of the agreement\n" +
                    "up to the signature page and then replaces those removed pages\n"+
                    "with the standard Higher Education Amendment Page.";

                string caption = "Make HED Amendment";

                lngResult = MessageBox.Show(msg,caption,MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (lngResult==DialogResult.Yes)
                {
                    lngPageNumberSignaturePage = CleanUpUtilities.FindSignaturePage();

                    sel = Globals.ThisAddIn.Application.Selection;
                    sel.HomeKey(Word.WdUnits.wdStory);
                    sel.Extend();
                    sel.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToNext, Word.WdGoToDirection.wdGoToAbsolute, lngPageNumberSignaturePage);
                    sel.Delete();

                    tmpl = Globals.ThisAddIn.Application.Templates[@"\\nyodska01\cbwide\RAS Contracts Management\Training Documents\CM Utilities v61.dotm"];
                    hedaddendum = tmpl.AutoTextEntries[strAUTOTEXT_AMENDMENT];
                    hedaddendum.Insert(sel.Range);

                    // Remove paragraph page before and consolidate signature page and Amendment page.
                    sel.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToNext, Word.WdGoToDirection.wdGoToAbsolute, SIGNATURE_PAGE_AMENDMENT);
                    sel.Range.ParagraphFormat.PageBreakBefore=-1;
                    sel.HomeKey(Word.WdUnits.wdStory);
                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        public void RefreshShortcuts_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Refresh Shortcuts",MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public void DeleteMyRoad_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Not Coming Soon. You should delete this", "DeleteMyRoad", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void FormatPrice_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Format Price", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void FormatDateSpellOutMonth_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Format Date Spell Out Month", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void FormatPhoneNumber_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Format Phone Number", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void InterfaceForSpellNumber_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            MessageBox.Show("Coming Soon", "Spell Number", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
