using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TranscriptV4
{
    public partial class Transcript_main : Form
    {
        public Transcript_main()
        {
            InitializeComponent();
        }

        List<string> stuff_to_remove = new List<string>();
        string last_name = "";

        private void go_Click(object sender, EventArgs e)
        {
            int copy_esc_01_count = 0;
            int copy_esc_02_count = 0;
            int copy_esc_03_count = 0;
            int copy_esc_04_count = 0;
            if (eb1.Checked)
            {
                string copy_esc_01 = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(sc1.Text.Trim()), @"temp_file_1"), sc1.Text.Trim().Split('.').Last<string>());
                System.IO.File.Copy(sc1.Text.Trim(), copy_esc_01, true);
                stuff_to_remove.Add(copy_esc_01);
                logit("成績輸入1:" + copy_esc_01);
                copy_esc_01_count = row_count_er(copy_esc_01, 1) - Convert.ToInt32(off1.Value);
                logit("學生數目" + copy_esc_01_count.ToString());
            }
            if (eb2.Checked)
            {
                string copy_esc_02 = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(sc2.Text.Trim()), @"temp_file_2"), sc2.Text.Trim().Split('.').Last<string>());
                System.IO.File.Copy(sc2.Text.Trim(), copy_esc_02, true);
                stuff_to_remove.Add(copy_esc_02);
                logit("成績輸入2:" + copy_esc_02);
                copy_esc_02_count = row_count_er(copy_esc_02, 1) - Convert.ToInt32(off2.Value);
                logit("學生數目" + copy_esc_02_count.ToString());
            }
            if (eb3.Checked)
            {
                string copy_esc_03 = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(sc3.Text.Trim()), @"temp_file_3"), sc3.Text.Trim().Split('.').Last<string>());
                System.IO.File.Copy(sc3.Text.Trim(), copy_esc_03, true);
                stuff_to_remove.Add(copy_esc_03);
                logit("成績輸入3:" + copy_esc_03);
                copy_esc_03_count = row_count_er(copy_esc_03, 1) - Convert.ToInt32(off3.Value);
                logit("學生數目" + copy_esc_03_count.ToString());
            }
            if (eb4.Checked)
            {
                string copy_esc_04 = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(sc4.Text.Trim()), @"temp_file_4"), sc4.Text.Trim().Split('.').Last<string>());
                System.IO.File.Copy(sc4.Text.Trim(), copy_esc_04, true);
                stuff_to_remove.Add(copy_esc_04);
                logit("成績輸入1:" + copy_esc_04);
                copy_esc_04_count = row_count_er(copy_esc_04, 1) - Convert.ToInt32(off4.Value);
                logit("學生數目" + copy_esc_04_count.ToString());
            }

            int check_if_same1 = copy_esc_01_count;
            int check_if_same2 = (eb2.Checked) ? copy_esc_02_count : check_if_same1;
            int check_if_same3 = (eb3.Checked) ? copy_esc_03_count : check_if_same1;
            int check_if_same4 = (eb4.Checked) ? copy_esc_04_count : check_if_same1;

            string copy_word_root = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(word_in_loc.Text.Trim()), @"temp_file_0"), word_in_loc.Text.Trim().Split('.').Last<string>());
            System.IO.File.Copy(word_in_loc.Text.Trim(), copy_word_root, true);
            stuff_to_remove.Add(copy_word_root);
            logit("模板輸入:" + copy_word_root);

            if ((check_if_same1 == check_if_same2) && (check_if_same3 == check_if_same4) && (check_if_same2 == check_if_same3))
            {
                pgb.Maximum = check_if_same1;
                for (int i = 1; i < copy_esc_01_count + 1; i++)
                {
                    pgb.Value++;
                    string copy_word_tmp = Path.ChangeExtension(Path.Combine(op_loc.Text.Trim(), @"座號" + i.ToString()), word_in_loc.Text.Trim().Split('.').Last<string>());
                    System.IO.File.Copy(word_in_loc.Text.Trim(), copy_word_tmp, true);
                    logit(i.ToString());
                    change_score(copy_word_tmp, i);
                    //string tPath = txtBrowse.Text.Trim();
                    //https://dotblogs.com.tw/rtomosaka/2017/09/28/csharp_openxml
                    //string tOutPath = string.Format(@"C:\Users\bear\Desktop\{0}.docx", DateTime.Now.ToString("yyyyMMddHHmmss"));
                    Dictionary<string, string> tReplaceDic = new Dictionary<string, string>();
                    string new_lastnamme = excel_numb_name(i+Convert.ToInt32(off1.Text),virb_name.Text);
                    string new_numb = excel_numb_name(i + Convert.ToInt32(off1.Text), virb_numb.Text);
                    tReplaceDic.Add("§§name§§", new_lastnamme);
                    tReplaceDic.Add("§§numb§§", new_numb);


                    bool tResult = WordReplace(copy_word_tmp, tReplaceDic);
                }
            }
            else
            {
                MessageBox.Show("輸入成績表單學生人數不相符");
                System.Windows.Forms.Application.Exit();
            }
            flush();
        }



        private void Transcript_main_Load(object sender, EventArgs e)
        {

        }

        public void change_score(string file_loc, int bypass_stu)
        {
            string copy_word_tmp = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(file_loc), @"tmp1" + bypass_stu.ToString()), file_loc.Split('.').Last<string>());
            System.IO.File.Copy(file_loc, copy_word_tmp, false);
            stuff_to_remove.Add(copy_word_tmp);
            // Open word document for read
            using (var doc = WordprocessingDocument.Open(file_loc, true))
            {
                int rowCount = 0;

                // Find the first table in the document. 
                DocumentFormat.OpenXml.Wordprocessing.Table table = doc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().First();

                // To get all rows from table
                IEnumerable<TableRow> rows = table.Elements<TableRow>();
                int testcounter = 0;
                // To read data from rows and to add records to the temporary table
                foreach (TableRow row in rows)
                {
                    if (rowCount == 0)
                    {
                        rowCount += 1;
                    }
                    else
                    {
                        int i = 0;
                        foreach (TableCell cell in row.Descendants<TableCell>())
                        {
                            if (i < off0.Value)
                            {//no move
                            }
                            else
                            {
                                string my_head = "";
                                string my_score = "";
                                string compA = "";
                                string compB = "";
                                int compAi = 0;
                                int compBi = 0;
                                int Cdiff = 0;
                                string diff_out = "";
                                switch (eb_checker())
                                {
                                    case 1:
                                        logit("流程1");
                                        my_head = get_head(copy_word_tmp, i);
                                        my_score = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);
                                        if (my_head == "姓名")
                                        {
                                            last_name = my_score;
                                        }
                                        Console.WriteLine(my_score);
                                        change_cell_text(cell, my_score);
                                        break;
                                    case 2:
                                        logit("流程2");
                                        switch (testcounter)
                                        {
                                            case 1:
                                                logit("流程21");

                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 2:
                                                logit("流程22");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 3:
                                                logit("流程-比較1");
                                                my_head = get_head(copy_word_tmp, i);
                                                if (trp1.Text == "目標")
                                                {
                                                    compA = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "目標")
                                                {
                                                    compA = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "目標")
                                                {
                                                    compA = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "目標")
                                                {
                                                    compA = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                if (trp1.Text == "參考")
                                                {
                                                    compB = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "參考")
                                                {
                                                    compB = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "參考")
                                                {
                                                    compB = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "參考")
                                                {
                                                    compB = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                Console.WriteLine("*****成績差:" + compA + compB);
                                                if (DS.Items.Contains(my_head) == false)
                                                {
                                                    if (Int32.TryParse(compA, out compAi))
                                                    {
                                                        Console.WriteLine("*******A比較過");
                                                        if (Int32.TryParse(compB, out compBi))
                                                        {
                                                            Console.WriteLine("*******B比較過");
                                                            Cdiff = compAi - compBi;
                                                            if (fan.Items.Contains(my_head))
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                change_cell_text(cell, diff_out);
                                                Console.WriteLine("差動輸出" + diff_out);
                                                break;
                                            default:
                                                break;
                                        }
                                        break;
                                    case 3:
                                        logit("流程3");
                                        switch (testcounter)
                                        {
                                            case 1:
                                                logit("流程31");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 2:
                                                logit("流程32");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 3:
                                                logit("流程33");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 4:
                                                logit("流程-比較1");
                                                my_head = get_head(copy_word_tmp, i);
                                                if (trp1.Text == "目標")
                                                {
                                                    compA = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "目標")
                                                {
                                                    compA = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "目標")
                                                {
                                                    compA = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "目標")
                                                {
                                                    compA = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                if (trp1.Text == "參考")
                                                {
                                                    compB = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "參考")
                                                {
                                                    compB = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "參考")
                                                {
                                                    compB = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "參考")
                                                {
                                                    compB = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                if (DS.Items.Contains(my_head) == false)
                                                {
                                                    if (Int32.TryParse(compA, out compAi))
                                                    {
                                                        Console.WriteLine("*******A比較過");
                                                        if (Int32.TryParse(compB, out compBi))
                                                        {
                                                            Console.WriteLine("*******B比較過");
                                                            Cdiff = compAi - compBi;
                                                            if (fan.Items.Contains(my_head))
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                change_cell_text(cell, diff_out);
                                                Console.WriteLine("差動輸出" + diff_out);
                                                break;
                                            default:
                                                break;
                                        }
                                        break;
                                    case 4:
                                        logit("流程4");
                                        switch (testcounter)
                                        {
                                            case 1:
                                                logit("流程41");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 2:
                                                logit("流程42");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;

                                            case 3:
                                                logit("流程43");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 4:
                                                logit("流程44");
                                                my_head = get_head(copy_word_tmp, i);
                                                my_score = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                if (my_head == "姓名")
                                                {
                                                    last_name = my_score;
                                                }
                                                Console.WriteLine(my_score);
                                                change_cell_text(cell, my_score);
                                                break;
                                            case 5:
                                                logit("流程-比較1");
                                                my_head = get_head(copy_word_tmp, i);
                                                if (trp1.Text == "目標")
                                                {
                                                    compA = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "目標")
                                                {
                                                    compA = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "目標")
                                                {
                                                    compA = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "目標")
                                                {
                                                    compA = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                if (trp1.Text == "參考")
                                                {
                                                    compB = excel_dt(sc1.Text.Trim(), bypass_stu + Convert.ToInt32(off1.Value), my_head);

                                                }
                                                else if (trp2.Text == "參考")
                                                {
                                                    compB = excel_dt(sc2.Text.Trim(), bypass_stu + Convert.ToInt32(off2.Value), my_head);
                                                }
                                                else if (trp3.Text == "參考")
                                                {
                                                    compB = excel_dt(sc3.Text.Trim(), bypass_stu + Convert.ToInt32(off3.Value), my_head);
                                                }
                                                else if (trp4.Text == "參考")
                                                {
                                                    compB = excel_dt(sc4.Text.Trim(), bypass_stu + Convert.ToInt32(off4.Value), my_head);
                                                }
                                                else
                                                {

                                                }
                                                if (DS.Items.Contains(my_head) == false)
                                                {
                                                    if (Int32.TryParse(compA, out compAi))
                                                    {
                                                        Console.WriteLine("*******A比較過");
                                                        if (Int32.TryParse(compB, out compBi))
                                                        {
                                                            Console.WriteLine("*******B比較過");
                                                            Cdiff = compAi - compBi;
                                                            if (fan.Items.Contains(my_head))
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (Cdiff < 0)
                                                                {
                                                                    diff_out = bmpT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else if (Cdiff > 0)
                                                                {
                                                                    diff_out = impT.Text + Math.Abs(Cdiff).ToString();
                                                                }
                                                                else
                                                                {
                                                                    diff_out = nmpT.Text;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                change_cell_text(cell, diff_out);
                                                Console.WriteLine("差動輸出" + diff_out);
                                                break;
                                            default:
                                                break;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                            i++;
                        }
                    }
                    testcounter++;
                }

            }
        }

        public string get_head(string file_loc, int in_i)
        {
            string stuff_to_return = "";

            // Open word document for read
            using (var doc = WordprocessingDocument.Open(file_loc, false))
            {
                // To create a temporary table 

                int rowCount = 0;

                // Find the first table in the document. 
                DocumentFormat.OpenXml.Wordprocessing.Table table = doc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().First();

                // To get all rows from table
                IEnumerable<TableRow> rows = table.Elements<TableRow>();

                // To read data from rows and to add records to the temporary table
                foreach (TableRow row in rows)
                {
                    if (rowCount == 0)
                    {
                        int i_tmp = 0;
                        foreach (TableCell cell in row.Descendants<TableCell>())
                        {
                            stuff_to_return = (i_tmp == in_i) ? cell.InnerText : stuff_to_return;
                            //logit(cell.InnerText);
                            i_tmp++;
                        }
                        rowCount += 1;
                    }
                    else
                    {
                    }
                }

            }
            return stuff_to_return;
        }
        private string excel_dt(string excel_file_where, int stu, string head_to_find)
        {

            string copy_exc_tmp = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(excel_file_where), @"tmp1" + stu.ToString()), excel_file_where.Split('.').Last<string>());
            System.IO.File.Copy(excel_file_where, copy_exc_tmp, true);
            stuff_to_remove.Add(copy_exc_tmp);
            string stuff_to_return = "";
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(excel_file_where, false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                int stunum = 1;
                foreach (Row row in rows) //this will also include your header row...
                {
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        string my_word = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        if (stunum == stu && excel_gethead(copy_exc_tmp, i) == head_to_find)
                        {
                            stuff_to_return = my_word;
                        }
                    }
                    stunum++;
                }
            }
            return stuff_to_return;
        }

        private string excel_gethead(string excel_file_where, int iin)
        {
            string stuff_to_return = "";
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(excel_file_where, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();
                int i_tmp = 0;
                foreach (Cell cell in rows.ElementAt(0))
                {
                    if (i_tmp == iin)
                    {
                        stuff_to_return = GetCellValue(spreadSheetDocument, cell);
                    }
                    i_tmp++;
                }
            }

            return stuff_to_return;
        }

        private int eb_checker()
        {
            int stuff_to_return = 0;
            stuff_to_return += (eb1.Checked == true) ? 1 : 0;
            stuff_to_return += (eb2.Checked == true) ? 1 : 0;
            stuff_to_return += (eb3.Checked == true) ? 1 : 0;
            stuff_to_return += (eb4.Checked == true) ? 1 : 0;
            return stuff_to_return;
        }
        public void flush()
        {
            foreach (var file_t_dl_obj in stuff_to_remove)
            {
                // Delete a file by using File class static method...
                if (System.IO.File.Exists(file_t_dl_obj))
                {
                    // Use a try block to catch IOExceptions, to
                    // handle the case of the file already being
                    // opened by another process.
                    try
                    {
                        System.IO.File.Delete(file_t_dl_obj);
                    }
                    catch (System.IO.IOException e)
                    {
                        Console.WriteLine(e.Message);
                        return;
                    }
                }

            }
            logit("完成數據移置作業");
        }

        private void fanrm_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(fanip.Text) == false)
            {
                if (fan.Items.Contains(fanip.Text)) // case sensitive is not important
                    fan.Items.Remove(fanip.Text);
            }
        }

        private void fanad_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(fanip.Text) == false)
            {
                if (!fan.Items.Contains(fanip.Text)) // case sensitive is not important
                    fan.Items.Add(fanip.Text);
            }
        }

        private void DSa_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(DSi.Text) == false)
            {
                if (!DS.Items.Contains(DSi.Text)) // case sensitive is not important
                    DS.Items.Add(DSi.Text);
            }
        }

        private void DSr_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(DSi.Text) == false)
            {
                if (DS.Items.Contains(DSi.Text)) // case sensitive is not important
                    DS.Items.Remove(DSi.Text);
            }
        }

        private string excel_numb_name(int stu, string head_to_find)
        {

            string copy_exc_tmp = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(sc1.Text.Trim()), @"name_numb_tmp1" + stu.ToString()), sc1.Text.Trim().Split('.').Last<string>());
            System.IO.File.Copy(sc1.Text.Trim(), copy_exc_tmp, true);
            stuff_to_remove.Add(copy_exc_tmp);
            string stuff_to_return = "";
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(sc1.Text.Trim(), false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                int stunum = 1;
                foreach (Row row in rows) //this will also include your header row...
                {
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        string my_word = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        if (stunum == stu && excel_gethead(copy_exc_tmp, i) == head_to_find)
                        {
                            stuff_to_return = my_word;
                        }
                    }
                    stunum++;
                }
            }
            return stuff_to_return;
        }

        
    }
}
