using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace CB_Utilities_v6_9
{
    static class CleanUpUtilities
    {
        /* Date Created:    06/19/2017
         * Author:          Adrian Jones
         * Purpose:         Separate code from the ribbon so it's not "so coupled together"
         * Updates:         06/19/2017 - None
         */
        public static void RemoveUnnecssaryRiders()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document currentDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraph fldPara;
            Word.Range fldRider;

            int intRidersTotal = 0;
            int intUnnecessaryRiders = 0;
            int intNecessaryRiders = 0;

            // const string strRIDER_NAME_TOKEN = "-";
            const string strRIDER_HEADER = "Schedule to College Board";

            try
            {
                if (!currentDoc.TrackRevisions)
                    currentDoc.TrackRevisions = true;

                if (currentDoc.Fields.Count != 0)
                {
                    foreach (Word.Field fld in currentDoc.Fields)
                    {
                        /* Each Rider is a separate paragraph with a field at the end
                         * The purpose of the code is to determine if that field is expanded
                         * to display the ENTIRE RIDER or if it is just a field at the end of
                         * a lone paragraph.  If lone paragraph, then delete.
                        */
                        if (fld.Type == Word.WdFieldType.wdFieldIf)
                        {
                            /* This is a much "simplier" and straight forward deletion of the unncessary riders
                             * This is MORE related to the architecture i.e. Merge Field is either "True" Or "False"
                             * So if the field code is "False = True", then delete it, that field
                             * Tested an it works.
                             * Look at old VBA code to see how complicated I made the selection.
                             */
                            fldPara = fld.Result.Paragraphs[1];
                            fldRider = fld.Result;
                            if (fld.Code.Text.Contains("\"False\" = \"True"))
                            {
                                intUnnecessaryRiders++;
                            }
                            else
                            {
                                fld.Unlink();
                                intNecessaryRiders++;
                                fldRider.Find.Execute(strRIDER_HEADER);
                                if (fldRider.Find.Found)
                                    fldRider.ParagraphFormat.PageBreakBefore = -1;
                            }
                            fldPara.Range.Select();
                            fldPara.Range.Delete();
                            intRidersTotal++;
                        }
                        else
                        {
                            /* Get rid of the highlighted Rider Names here
                             * During testing on 02/11/2016 noticed this ALONE works to clean up riders
                             */
                        }
                    }

                    if (intRidersTotal == 0)
                    {
                        MessageBox.Show("No Riders exist in this document:\n");
                    }
                    else
                        MessageBox.Show("Number of Unnecessary Riders Found: " + intUnnecessaryRiders + "\n"
                            + "Number of Necessary Riders Found " + intNecessaryRiders + "\n"
                            + "Number of Total Riders Found " + intRidersTotal);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static int FindSignaturePage()
        {
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            Word.Range tmprange = Globals.ThisAddIn.Application.ActiveDocument.StoryRanges[Word.WdStoryType.wdMainTextStory];
            const string strFIND_SIGNATURE_PAGE_TEXT = "Signature";
            do
            {
                tmprange.Find.Execute(strFIND_SIGNATURE_PAGE_TEXT, Word.WdFindWrap.wdFindContinue);
                tmprange.Select();
            } while (tmprange.Find.Found == true && tmprange.Information[Word.WdInformation.wdWithInTable]==false);

            return sel.Information[Word.WdInformation.wdActiveEndPageNumber];
        }

        private static void RemoveSurroundingTables()
        {
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            do
            {
                sel.Rows.ConvertToText(Word.WdTableFieldSeparator.wdSeparateByParagraphs, false);
            } while (sel.Information[Word.WdInformation.wdWithInTable]);

            //sel.ParagraphFormat.SpaceAfter = 0.0;
        }
    }
}