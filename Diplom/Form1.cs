using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using HtmlAgilityPack;
using System.Diagnostics;
using mshtml;

namespace Diplom
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            excelDataSource1.Fill();
        }
        private void bTZIOpen_Click(object sender, EventArgs e)
        {
            string htmlCode = "";
            using (WebClient client = new WebClient())
            {
                client.Headers.Add(HttpRequestHeader.UserAgent, "AvoidError");
                htmlCode = UTF8ToWin1251(client.DownloadString(textBox1.Text));
            }
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(htmlCode);
            foreach (var row in doc.DocumentNode.SelectNodes("//tr[td]"))
                TableTZI.Rows.Add(row.SelectNodes("td").Select(td => td.InnerText).ToArray());
            while (TableTZI.Rows[0].ItemArray[0].ToString() != "\n  1\n  ")
            {
                TableTZI.Rows[0].Delete();
            }
            TableTZI.Rows[TableTZI.Rows.Count - 1].Delete();
            TableTZI.Rows[TableTZI.Rows.Count - 1].Delete();
            TableTZI.Rows[TableTZI.Rows.Count - 1].Delete();
            TableTZI.Rows[TableTZI.Rows.Count - 1].Delete();
            for (int i = 0; i < TableTZI.Rows.Count; i++)
            {
                for (int j = 0; j < TableTZI.Columns.Count; j++)
                {
                    TableTZI.Rows[i].SetField<String>(j, TableTZI.Rows[i].ItemArray[j].ToString().Replace("\n", " "));
                }
            }
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add();
            Excel.Worksheet workSheet = workBook.ActiveSheet;
            workSheet.Cells[1, "A"] = "Name";
            for (int i = 2; i<TableTZI.Rows.Count;i++)
            {
                workSheet.Cells[i, "A"] = TableTZI.Rows[i].ItemArray[1].ToString().Replace("\n", " ") + TableTZI.Rows[i].ItemArray[2].ToString().Replace("\n", " ");
            }
            string path = Directory.GetCurrentDirectory();
            if (File.Exists(path+"\\db\\text.xlsx"))
            {
                File.Delete(path+"\\db\\text.xlsx");
            }
            workBook.Close(true, path+"\\db\\text.xlsx");
            excelApp.Quit();
            using (Process myProcess = new Process())
            {
                myProcess.StartInfo.UseShellExecute = false;
                if (File.Exists("C:\\Program Files (x86)\\Java\\jre1.8.0_231\\bin\\java.exe"))
                myProcess.StartInfo.FileName = "C:\\Program Files (x86)\\Java\\jre1.8.0_231\\bin\\java.exe";
                else myProcess.StartInfo.FileName = "C:\\Program Files\\Java\\jre1.8.0_231\\bin\\java.exe";
                myProcess.StartInfo.Arguments = "-jar TextClassifier.jar";
                myProcess.Start();         
            }
            
        }
        static string UTF8ToWin1251(string sourceStr)
        {
            Encoding utf8 = Encoding.GetEncoding("UTF-8");
            Encoding win1251 = Encoding.GetEncoding("Windows-1251");
            byte[] utf8Bytes = win1251.GetBytes(sourceStr);
            byte[] win1251Bytes = Encoding.Convert(utf8, win1251, utf8Bytes);
            return win1251.GetString(win1251Bytes);
        }
        static private string Win1251ToUTF8(string source)
        {
            Encoding utf8 = Encoding.GetEncoding("utf-8");
            Encoding win1251 = Encoding.GetEncoding("Windows-1251");
            byte[] utf8Bytes = win1251.GetBytes(source);
            byte[] win1251Bytes = Encoding.Convert(win1251, utf8, utf8Bytes);
            source = win1251.GetString(win1251Bytes);
            return source;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.варіантTableAdapter.Fill(this.dBDataSet.Варіант);
            this.разделТКВІTableAdapter.Fill(this.dBDataSet.РазделТКВІ);
            this.обьектыTableAdapter.Fill(this.dBDataSet.Обьекты);
            this.оглавTableAdapter.Fill(this.dBDataSet.Оглав);
            this.обьектыTableAdapter.Fill(this.dBDataSet.Обьекты);
            excelDataSource1.Fill();

        }
        void upd()
        {
            try
            {
                this.оглавTableAdapter.Update(this.dBDataSet.Оглав);
                this.обьектыTableAdapter.Update(this.dBDataSet.Обьекты);
                this.разделТКВІTableAdapter.Update(this.dBDataSet.РазделТКВІ);
                this.оглавTableAdapter.Update(this.dBDataSet.Оглав);
                this.варіантTableAdapter.Update(this.dBDataSet.Варіант);
                this.профільTableAdapter1.Update(this.dBDataSet.Профіль);
                this.тквіTableAdapter1.Update(this.dBDataSet.ТКВІ);
                //оглавTableAdapter.Update();
            }
            catch (SqlException ex) { MessageBox.Show(ex.ToString()); }
            оглавTableAdapter.Fill(dBDataSet.Оглав);
            this.разделТКВІTableAdapter.Fill(this.dBDataSet.РазделТКВІ);
            this.варіантTableAdapter.Fill(this.dBDataSet.Варіант);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            upd();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int row = gridView2.GetSelectedRows()[0];
            upd();
            String str = gridView2.GetRowCellValue(row, "Код").ToString();
            dBDataSet.Tables["Обьекты"].Rows.Add(str);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            upd();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int row = gridView7.GetSelectedRows()[0];
            upd();
            String str = gridView7.GetRowCellValue(row, "Код").ToString();
            dBDataSet.Tables["ТКВІ"].Rows.Add(str);
        }

        int[] copy(int wind)
        {
            int[] prof = new int[dataGridView1.ColumnCount];
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                if (int.TryParse(dataGridView1.Rows[0].Cells[i].Value.ToString(), out prof[i]))
                    prof[i] = int.Parse(dataGridView1.Rows[0].Cells[i].Value.ToString());
                else { prof[i] = 10; continue; }
                int winp;
                if (int.TryParse(gridView1.GetRowCellValue(i, gridView1.Columns[2 + i]).ToString().PadRight(1), out winp))
                    if (prof[i] < winp)
                        prof[i] = 10;
            }
            return prof;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Profile form = new Profile();
            form.ShowDialog();
            if (form.DialogResult != DialogResult.OK)
                return;
            int row = gridView8.GetSelectedRows()[0];
            upd();
            String str = gridView8.GetRowCellValue(row, "Код").ToString();
            object[] data = new object[gridView1.Columns.Count];
            data[0] = str;
            row = form.row;
            for (int j = 0; j < form.gridView1.Columns.Count-1; j++)
                data[j + 1] = form.gridView1.GetRowCellValue(row, form.gridView1.Columns[j]);
            dBDataSet.Tables["Профіль"].Rows.Add(data);
            upd();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ReportObj report = new ReportObj();
            DevExpress.XtraReports.UI.ReportPrintTool tool = new DevExpress.XtraReports.UI.ReportPrintTool(report);
            tool.ShowRibbonPreview();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ReportTKVI report = new ReportTKVI();
            DevExpress.XtraReports.UI.ReportPrintTool tool = new DevExpress.XtraReports.UI.ReportPrintTool(report);
            tool.ShowRibbonPreview();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ReportProf report = new ReportProf();
            DevExpress.XtraReports.UI.ReportPrintTool tool = new DevExpress.XtraReports.UI.ReportPrintTool(report);
            tool.ShowRibbonPreview();
        }
    }
}
