using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;

namespace Diplom
{
    public partial class ReportProf : DevExpress.XtraReports.UI.XtraReport
    {
        public ReportProf()
        {
            InitializeComponent();
        }

        private void groupHeaderBand2_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {

        }
    }
}
