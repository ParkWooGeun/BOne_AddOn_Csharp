using System;
using System.Windows.Forms;

namespace PSH_BOne_AddOn.EXT_Form
{
    public partial class FrmRPT_Viewer1 : System.Windows.Forms.Form
    {
        public FrmRPT_Viewer1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// CRViewer 객체의 접근용 프로퍼티
        /// </summary>
        public CrystalDecisions.Windows.Forms.CrystalReportViewer ReportViewer
        {
            get { return this.CRViewer; }
        }

        private void FrmRPT_Viewer1_Activated(object sender, EventArgs e)
        {
            if(this.Created == true)
            {
                this.Activate(); //리포트뷰어용 Form의 활성화 및 포커스 이동(화면 제일 위에 띄우기)
                //this.WindowState = FormWindowState.Maximized; //시작시 폼 최대화
            }
        }

        private void FrmRPT_Viewer1_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
