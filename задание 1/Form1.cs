using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Access = Microsoft.Office.Interop.Access;
using ADODB;

namespace задание_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            openFileDialog1.ShowDialog();
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            saveFileDialog1.ShowDialog();
            StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
            string line;
            string[] w;

            while (!sr.EndOfStream)
            {
                line = sr.ReadLine();
                w = line.Split(' ');

                st1.Surname = w[0];
                st1.Mat = Convert.ToInt32(w[1]);
                st1.Fiz = Convert.ToInt32(w[2]);
                st1.Him = Convert.ToInt32(w[3]);
                st1.Inf = Convert.ToInt32(w[4]);
                st1.Citizenship = w[5];

                double stip = st1.Stipa(1000);

                string s = st1.Surname + " " + stip.ToString() + " " + "рублей";
                sw.WriteLine(s);
            }
            sr.Close();
            sw.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Excel.Application appE = new Excel.Application();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            Excel.Workbook wb1 = appE.Workbooks.Open(openFileDialog1.FileName);
            appE.Visible = true;
            Excel.Workbook wb2 = appE.Workbooks.Add();
            wb2.Worksheets[1].cells[1, 1].Value = "Фамилия";
            wb2.Worksheets[1].cells[1, 2].Value = "Cтипенлия";
            int k = 1;
            while (wb1.Worksheets[1].cells[k + 1, 1].Value != null)
            {
                st1.Surname = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                st1.Mat = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 2].Value);
                st1.Fiz = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 3].Value);
                st1.Him = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 4].Value);
                st1.Inf = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 5].Value);
                st1.Citizenship = wb1.Worksheets[1].Cells[k + 1, 6].Value;
                double stip = st1.Stipa(1000);
                wb2.Worksheets[1].Cells[k + 1, 1].Value = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                wb2.Worksheets[1].Cells[k + 1, 2].Value = stip.ToString();
                k++;
            }
            wb1.Close();
            wb2.SaveAs(saveFileDialog1.FileName);
            wb2.Close();
            appE.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Excel.Application appE = new Excel.Application();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            openFileDialog1.ShowDialog();
            Excel.Workbook wb1 = appE.Workbooks.Open(openFileDialog1.FileName);
            saveFileDialog1.ShowDialog();
            Excel.Workbook wb2 = appE.Workbooks.Add();
            wb2.Worksheets[1].cells[1, 1].Value = "Фамилия";
            wb2.Worksheets[1].cells[1, 2].Value = "Cтипенлия";

            int k = 1;
            while (wb1.Worksheets[1].cells[k + 1, 1].Value != null)
            {
                st1.Surname = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                st1.Mat = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 2].Value);
                st1.Fiz = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 3].Value);
                st1.Him = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 4].Value);
                st1.Inf = Convert.ToInt32(wb1.Worksheets[1].Cells[k + 1, 5].Value);
                st1.Citizenship = wb1.Worksheets[1].Cells[k + 1, 6].Value;
                double stip = st1.Stipa(1000);
                wb2.Worksheets[1].Cells[k + 1, 1].Value = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                wb2.Worksheets[1].Cells[k + 1, 2].Value = stip.ToString();
                k++;
            }
            wb2.Worksheets[1].Range("a1:b11").Select();
            wb2.Worksheets[1].Shapes.AddChart(Excel.XlChartType.xlColumnStacked, 100, 10, 500, 400);


            Excel.Chart ch = appE.Charts.Add();
            string s = "b2:e" + k.ToString();
            Excel.Range r = wb1.Worksheets[1].Range(s);
            ch.SetSourceData(r, Type.Missing);
            ch.ChartType = Excel.XlChartType.xl3DColumn;
            ch.HasDataTable = true;
            ch.DataTable.Font.Size = 9;
            ch.HasTitle = true;
            ch.ChartTitle.Text = "Стипендии";
            ch.ChartTitle.Font.Size = 24;
            ch.ChartTitle.Font.Color = 100;
            Excel.Axis ox = ch.Axes(Excel.XlAxisType.xlCategory);
            ox.HasTitle = false;
            Excel.Axis oy = ch.Axes(Excel.XlAxisType.xlSeriesAxis);
            oy.HasTitle = true;
            oy.AxisTitle.Text = "Предмет";
            Excel.Axis oz = ch.Axes(Excel.XlAxisType.xlValue);
            oz.HasTitle = true;
            oz.AxisTitle.Text = "Оценки";
            ox.HasMajorGridlines = true;
            oy.HasMajorGridlines = true;
            oz.MajorUnit = 1;
            Excel.SeriesCollection ser = ch.SeriesCollection();
            ser.Item(1).Name = "Математика";
            ser.Item(2).Name = "Физика";
            ser.Item(3).Name = "Химия";
            ser.Item(4).Name = "Информатика";
            string v = "={\"";
            for (int j = 2; j < k; j++)
                v += wb1.Worksheets[1].Cells[j, 1].Value + "\";\"";
            v += wb1.Worksheets[1].Cells[k, 1].Value + "\"}";
            ser.Item(1).XValues = v;
            wb1.Close();
            wb2.SaveAs(saveFileDialog1.FileName);
            wb2.Close();
            appE.Quit();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Word.Application appW = new Word.Application();
            Excel.Application appE = new Excel.Application();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            Excel.Workbook wb1 = appE.Workbooks.Open(openFileDialog1.FileName);
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            Word.Document d1 = appW.Documents.Add();
            appW.Visible = true;
            appE.Visible = true;
            d1.Paragraphs.Add();
            Word.Range r1 = d1.Paragraphs[1].Range;
            r1.Bold = 1;
            r1.Font.Size = 16;
            r1.Text = "Стипендии";
            d1.Paragraphs[1].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            d1.Paragraphs.Add();
            d1.Paragraphs.Add();
            d1.Paragraphs[3].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            d1.Paragraphs[3].Range.Text = "Фамилия Стипендия";
            int k = 3;
            while (wb1.Worksheets[1].Cells[k - 1, 1].Value != null)
            {
                d1.Paragraphs.Add();
                st1.Surname = wb1.Worksheets[1].Cells[k - 1, 1].Value;
                st1.Mat = (int)wb1.Worksheets[1].Cells[k - 1, 2].Value;
                st1.Fiz = (int)wb1.Worksheets[1].Cells[k - 1, 3].Value;
                st1.Him = (int)wb1.Worksheets[1].Cells[k - 1, 4].Value;
                st1.Inf = (int)wb1.Worksheets[1].Cells[k - 1, 5].Value;
                st1.Citizenship = wb1.Worksheets[1].Cells[k - 1, 6].Value;
                double stip = st1.Stipa(1000);
                d1.Paragraphs[k + 1].Range.Text = wb1.Worksheets[1].Cells[k - 1, 1].Value + " " + stip.ToString() + " рублей";
                k++;
            }
            d1.SaveAs(saveFileDialog1.FileName);
            d1.Close();
            wb1.Close();
            appE.Quit();
            appW.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Word.Application appW = new Word.Application();
            Excel.Application appE = new Excel.Application();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            Excel.Workbook wb1 = appE.Workbooks.Open(openFileDialog1.FileName);
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            Word.Document d1 = appW.Documents.Add();
            appE.Visible = true;
            appW.Visible = true;
            d1.Paragraphs.Add();
            Word.Range r1 = d1.Paragraphs[1].Range;
            r1.Bold = 1;            
            r1.Font.Size = 16;
            d1.Paragraphs[1].Range.Text = "Стипендии";
            d1.Paragraphs[1].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            d1.Paragraphs.Add();
            d1.Paragraphs.Add();
            Word.Range r2 = d1.Paragraphs[3].Range;
            Word.Table t = d1.Tables.Add(r2, 1, 2, Word.WdTableFieldSeparator.wdSeparateByCommas);
            t.Cell(1, 1).Range.Text = "Фамилия";
            t.Cell(1, 2).Range.Text = "Стипендия";
            int k = 1;
            while (wb1.Worksheets[1].Cells[k + 1, 1].Value != null)
            {
                t.Rows.Add();
                st1.Surname = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                st1.Mat = (int)wb1.Worksheets[1].Cells[k + 1, 2].Value;
                st1.Fiz = (int)wb1.Worksheets[1].Cells[k + 1, 3].Value;
                st1.Him = (int)wb1.Worksheets[1].Cells[k + 1, 4].Value;
                st1.Inf = (int)wb1.Worksheets[1].Cells[k + 1, 5].Value;
                st1.Citizenship = wb1.Worksheets[1].Cells[k + 1, 6].Value;
                double stip = st1.Stipa(1000);
                t.Cell(k + 1, 1).Range.Text = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                t.Cell(k + 1, 2).Range.Text = stip.ToString() + " рублей";
                k++;
            }
            d1.SaveAs(saveFileDialog1.FileName);
            d1.Close();
            wb1.Close();
            appE.Quit();
            appW.Quit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Excel.Application appE = new Excel.Application();
            appE.Visible = true;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            Excel.Workbook wb1 = appE.Workbooks.Open(openFileDialog1.FileName);
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            PowerPoint.Application appPp = new PowerPoint.Application();
            appPp.Visible = MsoTriState.msoTrue;
            PowerPoint.Presentation p = appPp.Presentations.Add();
            PowerPoint.Slide sl1 = p.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            PowerPoint.TextRange tr = sl1.Shapes[1].TextFrame.TextRange;
            tr.Text = "Группа 24ИС03";
            tr.Font.Size = 48;
            tr.Font.Name = "Times New Roman";
            tr = sl1.Shapes[2].TextFrame.TextRange;
            tr.Text = "Итоги";
            tr.Font.Size = 36;
            tr.Font.Name = "Times New Roman";
            tr.Font.Color.RGB = 1;
            PowerPoint.Slide sl2 = p.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutText);
            sl2.Shapes[1].TextFrame.TextRange.Text = "Стипендии";
            int k = 1;
            string s = "";
            while (wb1.Worksheets[1].Cells[k + 1, 1].Value != null)
            {
                s += wb1.Worksheets[1].Cells[k + 1, 1].Value;
                st1.Mat = (int)wb1.Worksheets[1].Cells[k + 1, 2].Value;
                st1.Fiz = (int)wb1.Worksheets[1].Cells[k + 1, 3].Value;
                st1.Him = (int)wb1.Worksheets[1].Cells[k + 1, 4].Value;
                st1.Inf = (int)wb1.Worksheets[1].Cells[k + 1, 5].Value;
                st1.Citizenship = wb1.Worksheets[1].Cells[k + 1, 6].Value;
                double stip = st1.Stipa(1000);
                s = s + " " + stip.ToString() + " рублей" + "\r";
                k++;
            }
            k = 1;
            sl2.Shapes[2].TextFrame.TextRange.Text = s;
            PowerPoint.TextRange tr2 = sl2.Shapes[2].TextFrame.TextRange;
            tr2.Font.Size = 12;
            PowerPoint.Slide sl3 = p.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutTable);
            sl3.Shapes[1].TextFrame.TextRange.Text = "Стипендии";
            sl3.Shapes.AddTable(1, 2);
            PowerPoint.Table t = sl3.Shapes[2].Table;
            t.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Фамилия";
            t.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Стипендии";
            while (wb1.Worksheets[1].Cells[k + 1, 1].Value != null)
            {
                t.Rows.Add();
                st1.Surname = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                st1.Mat = (int)wb1.Worksheets[1].Cells[k + 1, 2].Value;
                st1.Fiz = (int)wb1.Worksheets[1].Cells[k + 1, 3].Value;
                st1.Him = (int)wb1.Worksheets[1].Cells[k + 1, 4].Value;
                st1.Inf = (int)wb1.Worksheets[1].Cells[k + 1, 5].Value;
                st1.Citizenship = wb1.Worksheets[1].Cells[k + 1, 6].Value;
                double stip = st1.Stipa(1000);
                t.Cell(k + 1, 1).Shape.TextFrame.TextRange.Text = wb1.Worksheets[1].Cells[k + 1, 1].Value;
                t.Cell(k + 1, 2).Shape.TextFrame.TextRange.Text = stip.ToString() + " рублей";
                k++;
            }
            PowerPoint.Slide sl4 = p.Slides.Add(4, PowerPoint.PpSlideLayout.ppLayoutObject);
            sl4.Shapes[1].TextFrame.TextRange.Text = "Фото";
            sl4.Shapes.AddPicture(@"C:\Users\artur\Desktop\Главная папка\проги для уника\ТП\фото\photo_2026-03-09_15-43-33.jpg", MsoTriState.msoFalse, MsoTriState.msoTrue, 120, 150, 500, 330);
            wb1.Close();
            appE.Quit();
            p.SaveAs(saveFileDialog1.FileName);
            p.Close();
            appPp.Quit();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Word.Application appW = new Word.Application();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            Word.Document d1 = appW.Documents.Open(openFileDialog1.FileName);
            appW.Visible = true;
            int n = d1.Paragraphs.Count;
            PowerPoint.Application appPp = new PowerPoint.Application();
            appPp.Visible = MsoTriState.msoTrue;
            PowerPoint.Presentation p1 = appPp.Presentations.Add();
            PowerPoint.Slide sl1 = p1.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            sl1.Shapes[1].TextFrame.TextRange.Text = d1.Paragraphs[1].Range.Text;
            sl1.Shapes[2].TextFrame.TextRange.Text = d1.Paragraphs[2].Range.Text;
            for (int k = 2; k < n; k++)
            { 
                sl1 = p1.Slides.Add(k, PowerPoint.PpSlideLayout.ppLayoutText);
                sl1.Shapes[1].TextFrame.TextRange.Text = "Абзац " + (k - 1).ToString();
                sl1.Shapes[2].TextFrame.TextRange.Text = d1.Paragraphs[k + 1].Range.Text;
            }
            d1.Close();
            appW.Quit();
            int m = (n + 2) / 4;
            int[] ind1 = new int[m];
            for (int k = 1; k <= m; k++)
                ind1[k - 1] = 4 * k - 3;
            PowerPoint.SlideShowTransition sst1 = p1.Slides.Range(ind1).SlideShowTransition;
            sst1.AdvanceOnTime = MsoTriState.msoTrue;
            sst1.AdvanceTime = 1;
            sst1.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverLeftDown;
            m = (n + 1) / 4;
            int[] ind2 = new int[m];
            for (int k = 1; k <= m; k++)
                ind2[k - 1] = 4 * k - 2;
            PowerPoint.SlideShowTransition sst2 = p1.Slides.Range(ind2).SlideShowTransition; sst2.AdvanceOnTime = MsoTriState.msoTrue;
            sst2.AdvanceTime = 1;
            sst2.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverRightDown;
            m = n / 4;
            int[] ind3 = new int[m];
            for (int k = 1; k <= m; k++)
                ind3[k - 1] = 4 * k - 1;
            PowerPoint.SlideShowTransition sst3 = p1.Slides.Range(ind3).SlideShowTransition;
            sst3.AdvanceOnTime = MsoTriState.msoTrue;
            sst3.AdvanceTime = 1;
            sst3.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverLeftUp;
            m = (n - 1) / 4;
            int[] ind4 = new int[m];
            for (int k = 1; k <= m; k++)
                ind4[k - 1] = 4 * k;
            PowerPoint.SlideShowTransition sst4 =
            p1.Slides.Range(ind4).SlideShowTransition;
            sst4.AdvanceOnTime = MsoTriState.msoTrue;
            sst4.AdvanceTime = 1;
            sst4.EntryEffect =
            PowerPoint.PpEntryEffect.ppEffectCoverRightUp;
            p1.SlideShowSettings.Run();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Student st1 = new Student();
            Excel.Application appE = new Excel.Application();
            Excel.Workbook wb1 = appE.Workbooks.Add();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            appE.Visible = true;
            Connection cnn = new Connection();
            cnn.Open(@"Provider=Microsoft.JET.OLEDB.4.0; Data Source=C:\Users\artur\Desktop\Главная папка\проги для уника\ТП\задание 1\оценки.mdb");
            Recordset rst = new Recordset();
            rst.Open("Select * From Оценки", cnn, CursorTypeEnum.adOpenKeyset, LockTypeEnum.adLockOptimistic);
            rst.MoveFirst();
            int k = 1;
            wb1.Worksheets[1].Cells[1, 1].Value = "Фамилия";
            wb1.Worksheets[1].Cells[1, 2].Value = "Стипендия";
            do
            {
                st1.Surname = rst.Fields["Фамилия"].Value;
                st1.Mat = rst.Fields["Математика"].Value;
                st1.Fiz = rst.Fields["Физика"].Value;
                st1.Him = rst.Fields["Химия"].Value;
                st1.Inf = rst.Fields["Информатика"].Value;
                st1.Citizenship = rst.Fields["Гражданство"].Value;
                double stip = st1.Stipa(1000);
                wb1.Worksheets[1].Cells[k + 1, 1].Value = st1.Surname;
                wb1.Worksheets[1].Cells[k + 1, 2].Value = stip.ToString() + " рублей";
                rst.MoveNext();
                k++;
            }
            while (!rst.EOF);
            wb1.SaveAs(saveFileDialog1.FileName);
            wb1.Close();
            appE.Quit();
            rst.Close();
            cnn.Close();
        }
    }
}
