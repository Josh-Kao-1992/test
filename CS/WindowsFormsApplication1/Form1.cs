using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraReports.UI;
using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace WindowsFormsApplication1 {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            XtraReport1 report = new XtraReport1();
            XtraReport2 report2 = new XtraReport2();
            report.CreateDocument(false);
            report2.CreateDocument(false);
            report.Pages.AddRange(report2.Pages);
            ReportPrintTool tool = new ReportPrintTool(report);
            tool.ShowPreviewDialog();
        }

        private void button2_Click(object sender, EventArgs e) {
            XtraReport1 report = new XtraReport1();
            XtraReport2 report2 = new XtraReport2();
            report.CreateDocument(false);
            report2.CreateDocument(false);
            report.ExportToXlsx("test1.xlsx", new DevExpress.XtraPrinting.XlsxExportOptions() { SheetName = "report1" });
            report2.ExportToXlsx("test2.xlsx", new DevExpress.XtraPrinting.XlsxExportOptions() { SheetName = "report2" });

            Workbook workbook = new DevExpress.Spreadsheet.Workbook();
            workbook.LoadDocument("test1.xlsx");

            Workbook workbook2 = new DevExpress.Spreadsheet.Workbook();
            workbook2.LoadDocument("test2.xlsx");

            workbook.Worksheets.Insert(1,"report2");
            workbook.Worksheets[1].CopyFrom(workbook2.Worksheets[0]);
            workbook.SaveDocument("test3.xlsx");
            Process.Start("test3.xlsx");
        }
    }
}
