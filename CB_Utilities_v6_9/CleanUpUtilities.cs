using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;

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

        public static void FormatPrice()
        {
            /* 08-1-2017 - The idea is to allow user to place *cursor* on number/price
             * to be formatted. Instead of looking at ALL the characters in the number,
             * the code "zooms" out to the next level of selection (price looks like a 
             * collection of "words" instead of the price being "1 word". So, zoom out
             * to a sentence and FIND the matching text
             * 
             * When the selection is larger than just a *cursor* Need to accomodate that.
             */
            // This is what the malformed price looks like.
            const string regexpattern = @"[$]\s?\d+\S\d{2}";

            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            Word.Range startrange = sel.Range;
            Word.Range searchrange = sel.Range;
            Word.Range tmprange = sel.Range;

            Regex regex = new Regex(regexpattern, RegexOptions.IgnoreCase);

            if (sel.Type==Word.WdSelectionType.wdSelectionIP) // Only a cursor
            {
                //The Smallest Search Context: a sentence (see comment above) WHEN multiple paragraphs is NOT selected.
                searchrange = sel.Sentences.First;
            }
            else if (startrange.Paragraphs.Count > 1)
            {
                searchrange = sel.Range;
            }

            MatchCollection ms = regex.Matches(searchrange.Text);
            Debug.WriteLine("Matches Found " + ms.Count);
            Debug.WriteLine("******************************************");
            
            foreach (Match singlematch in ms)
            {
                Debug.WriteLine("Match Value: " + singlematch.Value);
                Debug.WriteLine("Index/Start Position: " + singlematch.Index);
                Debug.WriteLine("Length: " + singlematch.Length);

                // Is startrange "position" within 1st Match, 2nd, etc.?
                if (startrange.InRange(tmprange))
                {
                    // Have to reevaluate this logic because it does not work in table Cells!!
                    tmprange.SetRange((searchrange.Start + singlematch.Index),
                            (searchrange.Start + (singlematch.Index + singlematch.Length)));
                    tmprange.Select();
                }
                else if (searchrange.Information[Word.WdInformation.wdWithInTable])
                {
                    //foreach (Word.Cell rngTableCell in sel.Cells) { }
                    // tmprange.Select();
                }
            }
            Debug.WriteLine("******************************************");


            /*
            string tmpPrice=String.Empty;
            if (sel.Information[Word.WdInformation.wdWithInTable])
            {
                foreach (Word.Cell rngTableCell in sel.Cells)
                {
                    tmpPrice = rngTableCell.Range.Text;
                    if (tmpPrice.Contains("$"))
                    {
                        //tmpPrice = tmpPrice.Replace("$", string.Empty);
                        tmpPrice = tmpPrice.Replace("\r\a",string.Empty);
                        //tmpPrice = tmpPrice.Trim();
                        tmpPrice = string.Format("{0:C}", float.Parse(tmpPrice));
                        rngTableCell.Range.Text = tmpPrice;
                    }
                }
            }
            */
            // else
            {
                /* Different behavior if user selects the entire number to format or
                 * if user only drops cursor somewhere in the number
                */
            }
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