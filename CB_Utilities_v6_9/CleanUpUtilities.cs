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
            } while (tmprange.Find.Found == true && tmprange.Information[Word.WdInformation.wdWithInTable] == false);

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

        }

        public static void SpellOutMonth()
        {
            const string DATE_SPELL_OUT_MONTH_FORMAT = "MMMM d, yyyy";
            const char CHAR_WORD_PARAGRAPH = '\r';
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            string DateNumberFormat;

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
        }

        public static void FormatPhoneNumber()
        {
            const string PHONE_FORMAT = "(###) ###-####";
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            string strPhoneNumberDigits;

            strPhoneNumberDigits = sel.Words[1].Text;
            strPhoneNumberDigits = strPhoneNumberDigits.Trim();

            if (strPhoneNumberDigits.Length == 10 && int.TryParse(strPhoneNumberDigits, out int result))
            {
                strPhoneNumberDigits = result.ToString(PHONE_FORMAT);
                sel.Words[1].Text = strPhoneNumberDigits;
            }
            else
                MessageBox.Show("Your selection does not solely consist of numbers\n"
                    + "or consists of more than or less than 10 digits - Number Count: " + strPhoneNumberDigits.Length,
                    "Format Phone #", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private static string Spell(int number)
        {
            var baseWords = new Hashtable();
            baseWords.Add(0, "");
            baseWords.Add(1, "one");
            baseWords.Add(2, "two");
            baseWords.Add(3, "three");
            baseWords.Add(4, "four");
            baseWords.Add(5, "five");
            baseWords.Add(6, "six");
            baseWords.Add(7, "seven");
            baseWords.Add(8, "eight");
            baseWords.Add(9, "nine");
            baseWords.Add(10, "ten");
            baseWords.Add(11, "eleven");
            baseWords.Add(12, "twelve");
            baseWords.Add(13, "thirteen");
            baseWords.Add(14, "fourteen");
            baseWords.Add(15, "fifteen");
            baseWords.Add(16, "sixteen");
            baseWords.Add(17, "seventeen");
            baseWords.Add(18, "eighteen");
            baseWords.Add(19, "nineteen");
            baseWords.Add(20, "twenty");
            baseWords.Add(30, "thirty");
            baseWords.Add(40, "forty");
            baseWords.Add(50, "fifty");
            baseWords.Add(60, "sixty");
            baseWords.Add(70, "seventy");
            baseWords.Add(80, "eighty");
            baseWords.Add(90, "ninety");

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
            return baseWords[number].ToString();
        }
    }
}