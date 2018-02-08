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

// For Hashtable
using System.Collections;

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

            if (!currentDoc.TrackRevisions)
            {
                currentDoc.TrackRevisions = true;
                currentDoc.ShowRevisions = false;
            }

            app.ScreenRefresh();

            try
            {
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
                            fldPara.Range.Select();
                            if (fld.Code.Text.Contains("\"False\" = \"True\""))
                            {
                                intUnnecessaryRiders++;
                                if (app.Selection.Characters.Count == 1)
                                {
                                    app.Selection.Paragraphs[1].Range.Delete();
                                }
                            }
                            else
                            {
                                fld.Unlink();
                                fldRider.Find.Execute(strRIDER_HEADER);
                                if (fldRider.Find.Found)
                                {
                                    fldRider.ParagraphFormat.PageBreakBefore = -1;
                                    intNecessaryRiders++;
                                }
                            }
                            fldPara.Range.Delete();
                            currentDoc.AcceptAllRevisions();
                            app.ScreenRefresh();
                        }
                        else
                        {
                            /* Get rid of the highlighted Rider Names here
                             * During testing on 02/11/2016 noticed this ALONE works to clean up riders
                             */
                        }
                    }

                    intRidersTotal = intNecessaryRiders + intUnnecessaryRiders;
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
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Remove Unnecessary Riders", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                if (!currentDoc.ShowRevisions)
                    currentDoc.ShowRevisions = true;

                app.ScreenUpdating = true;
                app.ScreenRefresh();
            }
        }

        public static void MakeHEDAmendment()
        {
            /* 07/27/2017 Create another class that handles all of this code or put it
             * in The CleanUpUtilites class and change its name to just Utilities
            */
            DialogResult lngResult;
            long lngPageNumberSignaturePage;
            Word.Selection sel;
            // Word.AutoTextEntry atxhedaddendum;
            Word.BuildingBlock hedaddendum;
            Word.Template tmpl;

            const string strAUTOTEXT_AMENDMENT = "HSA - HED Standard Addendum";
            const string strAGREEMENT_K12 = "COLLEGE READINESS";
            const string strAGREEMENT_HED = "ENROLLMENT AGREEMENT";
            const int SIGNATURE_PAGE_AMENDMENT = 2;

            bool foundHEDAgreement;

            try
            {
                string msg = "This deletes pages in the main part of the agreement\n" +
                    "up to the signature page and then replaces those removed pages\n" +
                    "with the standard Higher Education Amendment Page.";

                string msg2 = "Currently, Amendments can only be made from Higher Ed contracts\nComing Soon for K12";

                string caption = "Make HED Amendment";

                sel = Globals.ThisAddIn.Application.Selection;
                foundHEDAgreement = sel.Find.Execute(strAGREEMENT_HED);

                if (foundHEDAgreement)
                {
                    lngResult = MessageBox.Show(msg, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (lngResult == DialogResult.Yes)
                    {
                        CleanUpUtilities.TurnOffOnTrackChangesDisplay(false);
                        lngPageNumberSignaturePage = CleanUpUtilities.FindSignaturePage();

                        sel = Globals.ThisAddIn.Application.Selection;
                        sel.HomeKey(Word.WdUnits.wdStory);
                        sel.Extend();
                        sel.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToNext, Word.WdGoToDirection.wdGoToAbsolute, lngPageNumberSignaturePage);
                        sel.Delete();

                        // 07/26/2017 This template is not found even though I have a copy here now.
                        string templatefullname = @"\\nyodska01\cbwide\RAS Contracts Management\Training Documents\CM Utilities v62.dotx";

                        /* 07/27/2017 Of course the templates collection is a collection of all loaded add-ins
                         * so I may have to load the add-in here because it's no longer in Startup
                        */
                        tmpl = Globals.ThisAddIn.Application.Templates[templatefullname];
                        hedaddendum = tmpl.BuildingBlockEntries.Item(strAUTOTEXT_AMENDMENT);
                        hedaddendum.Insert(sel.Range, true);

                        lngPageNumberSignaturePage = CleanUpUtilities.FindSignaturePage();
                        // Remove paragraph page before and consolidate signature page and Amendment page.
                        sel.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToNext, Word.WdGoToDirection.wdGoToAbsolute, lngPageNumberSignaturePage);
                        Globals.ThisAddIn.Application.ScreenRefresh();
                        sel.Range.ParagraphFormat.PageBreakBefore = 0;
                        CleanUpUtilities.TurnOffOnTrackChangesDisplay(true);
                    }
                }
                else
                    MessageBox.Show(msg2, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
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

            TurnOffOnTrackChangesDisplay(false);
            do
            {
                tmprange.Find.Execute(strFIND_SIGNATURE_PAGE_TEXT, Word.WdFindWrap.wdFindContinue);
                tmprange.Select();
            } while (tmprange.Find.Found == true && tmprange.Information[Word.WdInformation.wdWithInTable] == false);

            TurnOffOnTrackChangesDisplay(true);
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
            const string regexpatternword = "[$]?[0-9]{4,}[.][0-9]{2}";
            char[] stripchars = { '$' };

            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            Word.Range searchrange = sel.Range;
            Word.Range endrange = sel.Range;

            TurnOffOnTrackChangesDisplay(false);
            Regex regex = new Regex(regexpattern, RegexOptions.IgnoreCase);
            sel.Find.Text = regexpatternword;
            sel.Find.MatchWildcards = true;

            if (sel.Type == Word.WdSelectionType.wdSelectionIP)
            {
                searchrange = sel.Sentences[1];
                searchrange.Select();
                searchrange.MoveEnd(Word.WdUnits.wdCharacter, -1);
                searchrange.Select();
            }

            MatchCollection selprices = regex.Matches(searchrange.Text);
            if (selprices.Count == 1)
            {
                sel.Text = double.Parse(sel.Text.Trim(stripchars)).ToString("C2");
            }
            else
            {
                sel.Find.Execute();
                while (sel.Find.Found && sel.InRange(searchrange))
                {
                    if (sel.Find.Found)
                    {
                        sel.Select();
                        sel.Text = double.Parse(sel.Text.Trim(stripchars)).ToString("C2");
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        endrange = sel.Range;
                    }
                    sel.Find.Execute();
                }
                endrange.Select();
            }
            TurnOffOnTrackChangesDisplay(true);
        }

        public static void SpellOutNumber()
        {
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            string strJustNumbers = String.Empty;
            string strNumber = String.Empty;
            string strDecimals = String.Empty;
            string strSpelledOutNumber = String.Empty;
            char[] stripchars = {'$', '/', ' ', ',', '\r', '\n'};

            strJustNumbers = sel.Text.Trim(stripchars);
            strNumber = strJustNumbers.Split('.')[0];
            strDecimals = strJustNumbers.Split('.')[1];

            TurnOffOnTrackChangesDisplay(false);
            if (int.TryParse(strNumber, out int resultNumber))
            {
                strSpelledOutNumber += Spell(resultNumber);
            }

            if (int.TryParse(strDecimals, out int resultDecimal))
            {
                strSpelledOutNumber += Spell(resultDecimal);
            }

            if (sel.Text.Contains('\r') || sel.Text.Contains('\n'))
            {
                sel.MoveEnd(Word.WdUnits.wdCharacter, -1);
            }
            sel.Text = strSpelledOutNumber;
            TurnOffOnTrackChangesDisplay(true);
        }

        public static void SpellOutMonth()
        {
            const string DATE_SPELL_OUT_MONTH_FORMAT = "MMMM d, yyyy";
            const char CHAR_WORD_PARAGRAPH = '\r';
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            string DateNumberFormat;

            TurnOffOnTrackChangesDisplay(false);
            if (sel.Text.Contains(CHAR_WORD_PARAGRAPH))
            {
                sel.MoveEnd(Word.WdUnits.wdCharacter, -1);
            }

            DateNumberFormat = sel.Text;
            if (DateTime.TryParse(DateNumberFormat, out DateTime result))
            {
                sel.Text = result.ToString(DATE_SPELL_OUT_MONTH_FORMAT);
            }
            else
                MessageBox.Show("A date doesn't appear to be selected\n", "Spell Out Date",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            TurnOffOnTrackChangesDisplay(true);
        }

        public static void FormatPhoneNumber()
        {
            const string PHONE_FORMAT = "(###) ###-####";
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            string strPhoneNumberDigits;

            strPhoneNumberDigits = sel.Words[1].Text;
            strPhoneNumberDigits = strPhoneNumberDigits.Trim();

            TurnOffOnTrackChangesDisplay(false);

            if (strPhoneNumberDigits.Length == 10 && long.TryParse(strPhoneNumberDigits, out long result))
            {
                strPhoneNumberDigits = result.ToString(PHONE_FORMAT);
                sel.Words[1].Text = strPhoneNumberDigits;
            }
            else
                MessageBox.Show("Your selection does not solely consist of numbers\n"
                    + "or consists of more than or less than 10 digits - Number Count: " + strPhoneNumberDigits.Length,
                    "Format Phone #", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            TurnOffOnTrackChangesDisplay(true);
        }

        public static void FormatCommonwealth()
        {
            Word.Document activedocument = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            Word.Range wrdState;
            Word.Range rngSelectedStateOfPhrase;

            string strCommonwealth=String.Empty;
            const int PARA_COMMONWEALTH = 3;
            const int POSITION_STATE_OF = 76;
            const int POSITION_STATE = 78;

            string state = String.Empty;
            bool IsCommonwealth = false;
            //string[] aresult;
            string[] acommonwealths = { "Kentucky", "Massachusetts", "Pennsylvania", "Virginia" };

            TurnOffOnTrackChangesDisplay(false);

            sel.HomeKey(Word.WdUnits.wdStory);
            wrdState = activedocument.Paragraphs[PARA_COMMONWEALTH].Range.Words[POSITION_STATE];
            rngSelectedStateOfPhrase = activedocument.Range(
                activedocument.Paragraphs[PARA_COMMONWEALTH].Range.Words[POSITION_STATE_OF].Start, 
                activedocument.Paragraphs[PARA_COMMONWEALTH].Range.Words[POSITION_STATE].End
                );

            rngSelectedStateOfPhrase.Select();

            // Search for match in array.
            IsCommonwealth = acommonwealths.Contains(wrdState.Text);

            //Replace "State" with "Commonwealth"
            if (IsCommonwealth)
            {
                activedocument.Paragraphs[PARA_COMMONWEALTH].Range.Words[POSITION_STATE_OF].Text = "Commonwealth ";
            }

            rngSelectedStateOfPhrase.Select();

            TurnOffOnTrackChangesDisplay(true);
        }

        public static void RemoveTerDatesFromFeeSchedule()
        {
            Word.Table tbl;
            Word.Table tblNewFromSplit;
            int rowHeaderRowColumnCount;
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            Word.Application app = Globals.ThisAddIn.Application;

            const string BMK_NAME_PARAGRAPH_SPLIT = "SplitTableParagarph";
            const string BMK_NAME_SPLIT_TABLE2 = "SplitTable2";

            try
            {
                if (sel.Information[Word.WdInformation.wdWithInTable])
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    app.ScreenUpdating = false;

                    tbl = sel.Tables[1];
                    tbl.AllowAutoFit = false;

                    rowHeaderRowColumnCount = tbl.Rows[1].Cells.Count;
                    /* Take the 1st rows count of columns --- that's the header --- and cycle through the rows until reaching a row that
                     * has a DIFFERENT column count
                     */
                    foreach (Word.Row currRow in tbl.Rows)
                    {
                        if (currRow.Cells.Count != rowHeaderRowColumnCount)
                        {
                            currRow.Select();
                            // Get reference to bottom half of table because of split by adding bookmark
                            sel.Bookmarks.Add(BMK_NAME_SPLIT_TABLE2, sel.Range);
                            sel.SplitTable();
                            sel.Bookmarks.Add(BMK_NAME_PARAGRAPH_SPLIT, sel.Range);

                            tblNewFromSplit = app.ActiveDocument.Bookmarks[BMK_NAME_SPLIT_TABLE2].Range.Tables[1];

                            //The original table selected, before the split, is the "top half" of the table
                            tbl.Columns[2].Select();
                            tbl.Columns[2].Delete();
                            tbl.Columns[2].Select();
                            tbl.Columns[2].Delete();

                            foreach (Word.Cell c in tbl.Columns[2].Cells)
                            {
                                c.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            app.ActiveDocument.Bookmarks[BMK_NAME_SPLIT_TABLE2].Delete();

                            /* New method of resizing table parts
                             * Resize Table Top - Table 1 of split, the size of window
                             */
                            tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                            tblNewFromSplit.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

                            tbl.AllowAutoFit = false;
                            tblNewFromSplit.AllowAutoFit = false;
                            app.ActiveDocument.Bookmarks[BMK_NAME_PARAGRAPH_SPLIT].Range.Delete();
                            app.ActiveDocument.Bookmarks[BMK_NAME_PARAGRAPH_SPLIT].Delete();

                            //Reinforce resize of entier table
                            tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                            break;
                        }
                    }
                }
                else
                    MessageBox.Show("You have to 1st put the cursor into a table", "Remove Term Date Columns in Fee Schedule", MessageBoxButtons.OK, MessageBoxIcon.Information);

                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
                app.ScreenUpdating = true;
            }
            catch (Exception e)
            {
                throw;
            }

        }

        private static void RemoveSurroundingTables()
        {
            TurnOffOnTrackChangesDisplay(false);

            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            do
            {
                sel.Rows.ConvertToText(Word.WdTableFieldSeparator.wdSeparateByParagraphs, false);
            } while (sel.Information[Word.WdInformation.wdWithInTable]);

            TurnOffOnTrackChangesDisplay(true);
            //sel.ParagraphFormat.SpaceAfter = 0.0;
        }

        private static string Spell(int number)
        {
            var baseWords = new Hashtable
            {
                { 0, "" },
                { 1, "one" },
                { 2, "two" },
                { 3, "three" },
                { 4, "four" },
                { 5, "five" },
                { 6, "six" },
                { 7, "seven" },
                { 8, "eight" },
                { 9, "nine" },
                { 10, "ten" },
                { 11, "eleven" },
                { 12, "twelve" },
                { 13, "thirteen" },
                { 14, "fourteen" },
                { 15, "fifteen" },
                { 16, "sixteen" },
                { 17, "seventeen" },
                { 18, "eighteen" },
                { 19, "nineteen" },
                { 20, "twenty" },
                { 30, "thirty" },
                { 40, "forty" },
                { 50, "fifty" },
                { 60, "sixty" },
                { 70, "seventy" },
                { 80, "eighty" },
                { 90, "ninety" }
            };

            TurnOffOnTrackChangesDisplay(false);

            if (number >= 1000)
            {
                return Spell(number / 1000) + " thousand " + Spell(number % 1000);
            }
            if (number >= 100)
            {
                return Spell(number / 100) + " hundred " + Spell(number % 100);
            }
            if (number >= 21)
            {
                return baseWords[number / 10 * 10] + " " + baseWords[number % 10];
            }

            TurnOffOnTrackChangesDisplay(true);

            return baseWords[number].ToString();
        }

        public static void TurnOffOnTrackChangesDisplay(bool switchOnOff)
        {
            Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions = switchOnOff;
            Globals.ThisAddIn.Application.ScreenRefresh();
        }
    }
}