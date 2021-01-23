using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TranscriptV4
{
    public partial class Transcript_main : Form
    {

        private void lsc1_Click(object sender, EventArgs e)
        {
            DialogResult result = this.scs1.ShowDialog();
            if (result == DialogResult.OK)
            {
                sc1.Text = scs1.FileName;
            }
        }

        private void lsc2_Click(object sender, EventArgs e)
        {
            DialogResult result = this.scs2.ShowDialog();
            if (result == DialogResult.OK)
            {
                sc2.Text = scs2.FileName;
            }
        }

        private void lsc3_Click(object sender, EventArgs e)
        {
            DialogResult result = this.scs3.ShowDialog();
            if (result == DialogResult.OK)
            {
                sc3.Text = scs3.FileName;
            }
        }

        private void lsc4_Click(object sender, EventArgs e)
        {
            DialogResult result = this.scs4.ShowDialog();
            if (result == DialogResult.OK)
            {
                sc4.Text = scs4.FileName;
            }
        }

        private void word_in_bton_Click(object sender, EventArgs e)
        {
            DialogResult result = this.read_word.ShowDialog();
            if (result == DialogResult.OK)
            {
                word_in_loc.Text = read_word.FileName;
            }
        }

        private void op_file_load_Click(object sender, EventArgs e)
        {
            DialogResult result = this.opf.ShowDialog();
            if (result == DialogResult.OK)
            {
                op_loc.Text = opf.SelectedPath;
            }
        }

        public void logit(string log_stuff)
        {
            logs.AppendText(log_stuff + Environment.NewLine);
        }
        private void btrp1_Click(object sender, EventArgs e)
        {

            switch (trp1.Text)
            {
                case "忽略":

                    int four_two1 = 0;
                    four_two1 += (trp1.Text == "目標") ? 1 : 0;
                    four_two1 += (trp2.Text == "目標") ? 1 : 0;
                    four_two1 += (trp3.Text == "目標") ? 1 : 0;
                    four_two1 += (trp4.Text == "目標") ? 1 : 0;
                    if (four_two1 == 0)
                    {
                        trp1.Text = "目標";
                    }
                    else
                    {
                        int four_two2 = 0;
                        four_two2 += (trp1.Text == "參考") ? 1 : 0;
                        four_two2 += (trp2.Text == "參考") ? 1 : 0;
                        four_two2 += (trp3.Text == "參考") ? 1 : 0;
                        four_two2 += (trp4.Text == "參考") ? 1 : 0;
                        if (four_two2 == 0)
                        {
                            trp1.Text = "參考";
                        }
                        else
                        {
                            //no move
                        }
                    }
                    break;
                case "目標":

                    trp1.Text = "忽略";

                    break;
                case "參考":
                    trp1.Text = "忽略";
                    break;
                default:
                    break;
            }
        }

        private void btrp2_Click(object sender, EventArgs e)
        {
            switch (trp2.Text)
            {
                case "忽略":

                    int four_two1 = 0;
                    four_two1 += (trp1.Text == "目標") ? 1 : 0;
                    four_two1 += (trp2.Text == "目標") ? 1 : 0;
                    four_two1 += (trp3.Text == "目標") ? 1 : 0;
                    four_two1 += (trp4.Text == "目標") ? 1 : 0;
                    if (four_two1 == 0)
                    {
                        trp2.Text = "目標";
                    }
                    else
                    {
                        int four_two2 = 0;
                        four_two2 += (trp1.Text == "參考") ? 1 : 0;
                        four_two2 += (trp2.Text == "參考") ? 1 : 0;
                        four_two2 += (trp3.Text == "參考") ? 1 : 0;
                        four_two2 += (trp4.Text == "參考") ? 1 : 0;
                        if (four_two2 == 0)
                        {
                            trp2.Text = "參考";
                        }
                        else
                        {
                            //no move
                        }
                    }
                    break;
                case "目標":

                    trp2.Text = "忽略";

                    break;
                case "參考":
                    trp2.Text = "忽略";
                    break;
                default:
                    break;
            }
        }

        private void btrp3_Click(object sender, EventArgs e)
        {
            switch (trp3.Text)
            {
                case "忽略":

                    int four_two1 = 0;
                    four_two1 += (trp1.Text == "目標") ? 1 : 0;
                    four_two1 += (trp2.Text == "目標") ? 1 : 0;
                    four_two1 += (trp3.Text == "目標") ? 1 : 0;
                    four_two1 += (trp4.Text == "目標") ? 1 : 0;
                    if (four_two1 == 0)
                    {
                        trp3.Text = "目標";
                    }
                    else
                    {
                        int four_two2 = 0;
                        four_two2 += (trp1.Text == "參考") ? 1 : 0;
                        four_two2 += (trp2.Text == "參考") ? 1 : 0;
                        four_two2 += (trp3.Text == "參考") ? 1 : 0;
                        four_two2 += (trp4.Text == "參考") ? 1 : 0;
                        if (four_two2 == 0)
                        {
                            trp3.Text = "參考";
                        }
                        else
                        {
                            //no move
                        }
                    }
                    break;
                case "目標":

                    trp3.Text = "忽略";

                    break;
                case "參考":
                    trp3.Text = "忽略";
                    break;
                default:
                    break;
            }
        }

        private void btrp4_Click(object sender, EventArgs e)
        {
            switch (trp4.Text)
            {
                case "忽略":

                    int four_two1 = 0;
                    four_two1 += (trp1.Text == "目標") ? 1 : 0;
                    four_two1 += (trp2.Text == "目標") ? 1 : 0;
                    four_two1 += (trp3.Text == "目標") ? 1 : 0;
                    four_two1 += (trp4.Text == "目標") ? 1 : 0;
                    if (four_two1 == 0)
                    {
                        trp4.Text = "目標";
                    }
                    else
                    {
                        int four_two2 = 0;
                        four_two2 += (trp1.Text == "參考") ? 1 : 0;
                        four_two2 += (trp2.Text == "參考") ? 1 : 0;
                        four_two2 += (trp3.Text == "參考") ? 1 : 0;
                        four_two2 += (trp4.Text == "參考") ? 1 : 0;
                        if (four_two2 == 0)
                        {
                            trp4.Text = "參考";
                        }
                        else
                        {
                            //no move
                        }
                    }
                    break;
                case "目標":

                    trp4.Text = "忽略";

                    break;
                case "參考":
                    trp4.Text = "忽略";
                    break;
                default:
                    break;
            }
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell, int maxDecPoint_in)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            //浮點數讀取的問題出在這裡，value變數，不管了，暴力解

            //Console.WriteLine("value:"+ value+ "\nif:"+ (cell.DataType != null && cell.DataType.Value == CellValues.SharedString).ToString()+"\nyes:"+ stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText);

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                //Console.WriteLine("error check:" + value);
                // Console.WriteLine("error check:"+stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText);

                string floaterror = value;
                bool gotcha = false;
                int countOnMe = 0;
                int thisIsCountOfIndex = 0;
                Regex isnumber1 = new Regex(@"^[0-9]$");
                foreach (var charItem in floaterror)
                {
                    if (isnumber1.IsMatch(charItem.ToString()) || charItem == '.')
                    {
                        //nomove
                    }
                    else
                    {
                        //it is not nubmer or 小數
                        break;
                    }
                    if (gotcha)
                    {
                        if (countOnMe == maxDecPoint_in)
                        {

                            floaterror = floaterror.Remove(thisIsCountOfIndex + 1);
                            floaterror = Math.Round(Decimal.Parse(floaterror), maxDecPoint_in).ToString();

                            break;
                        }
                        countOnMe++;
                    }
                    else
                    {
                        if (charItem == '.')
                        {
                            gotcha = true;
                        }
                    }
                    thisIsCountOfIndex++;
                }
                return floaterror;
            }
        }

        private static String HexConverter(System.Drawing.Color c)
        {
            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        public DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues HighlightColorValuesOfUser()
        {
            switch (highlightColorPicker.SelectedIndex)
            {
                case 0:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Black;
                    break;
                case 1:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Blue;
                    break;
                case 2:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Cyan;
                    break;
                case 3:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Green;
                    break;
                case 4:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Magenta;
                    break;
                case 5:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Red;
                    break;
                case 6:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Yellow;
                    break;
                case 7:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.White;
                    break;
                case 8:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkBlue;
                    break;
                case 9:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkCyan;
                    break;
                case 10:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkGreen;
                    break;
                case 11:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkMagenta;
                    break;
                case 12:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkRed;
                    break;
                case 13:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkYellow;
                    break;
                case 14:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.DarkGray;
                    break;
                case 15:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.LightGray;
                    break;
                case 16:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.None;
                    break;
                default:
                    return DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.None;
                    break;
            }

        }
        public void change_cell_text(TableCell cin, string stuff_to_change)
        {
            cin.RemoveAllChildren();

            Paragraph new_p = new Paragraph();
            //DocumentFormat.OpenXml.Wordprocessing.Run new_r = new DocumentFormat.OpenXml.Wordprocessing.Run();
            //DocumentFormat.OpenXml.Wordprocessing.Text new_t = new DocumentFormat.OpenXml.Wordprocessing.Text();
            //new_t.Text = stuff_to_change;
            //new_r.AppendChild(new_t);

            ParagraphProperties UserHeadingParagPro = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
            DocumentFormat.OpenXml.Wordprocessing.Run formattedRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
            DocumentFormat.OpenXml.Wordprocessing.RunProperties runPro = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            RunFonts runFont = new RunFonts() { Ascii = chooseFonts.SelectedItem.ToString(), HighAnsi = chooseFonts.SelectedItem.ToString() };
            DocumentFormat.OpenXml.Wordprocessing.Bold bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
            DocumentFormat.OpenXml.Wordprocessing.Text text = new DocumentFormat.OpenXml.Wordprocessing.Text(stuff_to_change);
            DocumentFormat.OpenXml.Wordprocessing.Color color;
            DocumentFormat.OpenXml.Wordprocessing.Highlight highlight;
            bool color_need_change = false;

            if (ifColor.Checked)
            {
                if (stuff_to_change.Contains("進"))
                {
                    if (ifciI.Checked)
                    {
                        color_need_change = true;
                    }
                }
                else if (stuff_to_change.Contains("退"))
                {
                    if (ifciB.Checked)
                    {
                        color_need_change = true;
                    }
                }
                else if (decimal.TryParse(stuff_to_change, out decimal numValue))
                {
                    decimal tmp_numValue= numValue;
                    decimal tmpSix = 60;
                    if (decimal.TryParse(failScore.Text,out decimal six))
                    {
                        tmpSix = six;
                    }
                    if (failUpDown.Checked)//>
                    {
                        if (tmp_numValue> tmpSix)
                        {
                            color_need_change = true;
                        }
                    }
                    else//小於
                    {
                        if (tmp_numValue< tmpSix)
                        {
                            color_need_change = true;
                        }
                    }

                }
                else
                {
                    //Int32.TryParse could not parse '{inputString}' to an int.");

                }

            }
            else
            {
                color_need_change = false;
            }

            if (color_need_change)
            {
                color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = HexConverter(COLORing.BackColor) };
                highlight = new DocumentFormat.OpenXml.Wordprocessing.Highlight() { Val = HighlightColorValuesOfUser() };
            }
            else
            {
                color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "000000" };
                highlight = new DocumentFormat.OpenXml.Wordprocessing.Highlight() { Val = HighlightColorValues.None };

            }

            Justification CenterHeading = new Justification { Val = user_justification() };
            UserHeadingParagPro.Append(CenterHeading);

            runPro.Append(runFont);
            if (blodAll.Checked) { runPro.Append(bold); } else { }
            runPro.Append(color);
            runPro.Append(highlight);
            runPro.Append(text);
            formattedRun.Append(runPro);

            new_p.Append(UserHeadingParagPro);
            new_p.AppendChild(formattedRun);
            cin.AppendChild(new_p);

            if (printAllData.Checked) { logit(stuff_to_change); }
        }
        public JustificationValues user_justification()
        {
            switch (JustificationValuesEnum.SelectedIndex)
            {
                case 0:
                    return JustificationValues.Left;
                    break;
                case 1:
                    return JustificationValues.Start;
                    break;
                case 2:
                    return JustificationValues.Center;
                    break;
                case 3:
                    return JustificationValues.Right;
                    break;
                case 4:
                    return JustificationValues.End;
                    break;
                case 5:
                    return JustificationValues.Both;
                    break;
                case 6:
                    return JustificationValues.MediumKashida;
                    break;
                case 7:
                    return JustificationValues.Distribute;
                    break;
                case 8:
                    return JustificationValues.NumTab;
                    break;
                case 9:
                    return JustificationValues.HighKashida;
                    break;
                case 10:
                    return JustificationValues.LowKashida;
                    break;
                case 11:
                    return JustificationValues.ThaiDistribute;
                    break;
                default:
                    return JustificationValues.Center;
                    break;
            }
        }
        public int row_count_er(string file_in, int sd_page)
        {
            int stuff_to_return = 0;
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(file_in, true))
            {
                //Get workbookpart
                WorkbookPart workbookPart = myDoc.WorkbookPart;

                //then access to the worksheet part
                IEnumerable<WorksheetPart> worksheetPart = workbookPart.WorksheetParts;

                foreach (WorksheetPart WSP in worksheetPart)
                {
                    //find sheet data
                    IEnumerable<SheetData> sheetData = WSP.Worksheet.Elements<SheetData>();
                    // Iterate through every sheet inside Excel sheet
                    int count_sd = 1;
                    foreach (SheetData SD in sheetData)
                    {
                        // Get the row IEnumerator
                        IEnumerable<Row> row = SD.Elements<Row>();

                        if (count_sd == sd_page)
                        {
                            stuff_to_return = row.Count();
                        }
                        count_sd++;
                    }
                }
            }
            return stuff_to_return;

        }
        private bool WordReplace(string pTemplatePath, Dictionary<string, string> pReplaceDic)
        {
            bool tResultbool = true;

            try
            {
                string tPath = pTemplatePath;

                if (File.Exists(tPath) == true)
                {

                    using (WordprocessingDocument tWordDocument = WordprocessingDocument.Open(tPath, true))
                    {
                        Body tBody = tWordDocument.MainDocumentPart.Document.Body;

                        foreach (KeyValuePair<string, string> tKeyVP in pReplaceDic)
                        {
                            string tSource = tKeyVP.Value;
                            string tTarget = tKeyVP.Key;
                            char[] tTargetArray = tTarget.ToCharArray();

                            foreach (Paragraph tParagraph in tBody.Descendants<Paragraph>())
                            {
                                //若尋找目標存在於此段落
                                if (tParagraph.InnerText.Trim().Contains(tTarget) == true)
                                {
                                    int tIndex = 0;  //Target Index
                                    int tRunIndex = 0; // 目前 Run 的 Index
                                    int tStartRun = 0; //啟始的 Run
                                    string tStartString = string.Empty;
                                    int tEndRun = 0; //最後的 Run
                                    string tEndString = string.Empty;
                                    bool tIsFindStart = false;
                                    foreach (DocumentFormat.OpenXml.Wordprocessing.Run tRun in tParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                                    {
                                        char[] tSourceArray = tRun.InnerText.Trim().ToCharArray();

                                        bool tIsFind = FindWord(tTargetArray, tSourceArray, ref tIndex);

                                        if (tIsFind == true)
                                        {
                                            if (tIsFindStart == false)
                                            {
                                                tIsFindStart = true; //記錄目前有找到Run
                                                tStartRun = tRunIndex; //記錄第一個找到的Run
                                                tStartString = tTarget.Substring(0, tIndex + 1);
                                                tTargetArray = null;
                                                tTargetArray = tTarget.Substring(tIndex + 1, tTarget.Length - (tIndex + 1)).ToCharArray();
                                                tIndex = 0;

                                                if (tTargetArray.Length == 0)
                                                {
                                                    tEndRun = tRunIndex;
                                                    tEndString = tStartString;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                string tTempString = new string(tTargetArray);
                                                tTargetArray = null;
                                                tTargetArray = tTempString.Substring(tIndex + 1, tTempString.Length - (tIndex + 1)).ToCharArray();
                                                tIndex = 0;

                                                if (tTargetArray.Length == 0)
                                                {
                                                    tEndRun = tRunIndex;
                                                    tEndString = tTempString;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (tIsFindStart == true)
                                            {
                                                tStartRun = -1;
                                                tIsFindStart = false;
                                                tTargetArray = null;
                                                tTargetArray = tTarget.ToCharArray();
                                                tIndex = 0;
                                                tStartString = string.Empty;
                                                tEndString = string.Empty;
                                            }
                                        }

                                        tRunIndex++;
                                    }

                                    tRunIndex = 0;
                                    foreach (DocumentFormat.OpenXml.Wordprocessing.Run tRun in tParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                                    {
                                        if (tRunIndex >= tStartRun && tRunIndex <= tEndRun)
                                        {
                                            foreach (DocumentFormat.OpenXml.Wordprocessing.Text tText in tRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                                            {
                                                if (tRunIndex == tStartRun || tStartRun == tEndRun)
                                                {
                                                    tText.Text = tText.Text.Replace(tStartString, tSource);
                                                }
                                                else if (tRunIndex == tEndRun)
                                                {
                                                    tText.Text = tText.Text.Replace(tEndString, "");
                                                }
                                                else
                                                {
                                                    tText.Text = "";
                                                }
                                            }
                                        }
                                        tRunIndex++;
                                    }

                                    //重新設定一次要尋找的目標
                                    tTargetArray = tTarget.ToCharArray();
                                }
                            }
                        }

                        tWordDocument.MainDocumentPart.Document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                tResultbool = false;
            }

            return tResultbool;
        }

        private bool FindWord(char[] pTargetArray, char[] pSourceArray, ref int pTargetIndex)
        {
            bool tResultBool = true;

            if (pSourceArray.Length > 0)
            {
                for (int i = 0; i < pSourceArray.Length; i++)
                {
                    if (pSourceArray[i] == pTargetArray[pTargetIndex])
                    {
                        if ((pTargetIndex + 1) == pTargetArray.Length || (pTargetIndex + 1) == pSourceArray.Length)
                        {
                            break;
                        }
                        pTargetIndex++;
                    }
                    else
                    {
                        if ((pTargetIndex + 1) == pTargetArray.Length && (pSourceArray.Length - (pTargetIndex + 1)) == 0)
                        {
                            tResultBool = false;
                            pTargetIndex = 0;
                            break;
                        }
                        else
                        {
                            //重置 Source Array
                            string tTempString = new string(pSourceArray);
                            //一次右移一碼
                            char[] tSourceArray = tTempString.Substring(1, pSourceArray.Length - 1).ToCharArray();
                            //重置目標 index
                            pTargetIndex = 0;

                            tResultBool = FindWord(pTargetArray, tSourceArray, ref pTargetIndex);
                            break;
                        }
                    }
                }
            }
            else
            {
                tResultBool = false;
                pTargetIndex = 0;
            }

            return tResultBool;
        }


    }
}