
namespace PSH_BOne_AddOn.EXT_Form
{
    partial class FrmRPT_Viewer1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CRViewer = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // CRViewer
            // 
            this.CRViewer.ActiveViewIndex = -1;
            this.CRViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CRViewer.Cursor = System.Windows.Forms.Cursors.Default;
            this.CRViewer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CRViewer.Location = new System.Drawing.Point(0, 0);
            this.CRViewer.Name = "CRViewer";
            this.CRViewer.Size = new System.Drawing.Size(1280, 960);
            this.CRViewer.TabIndex = 0;
            this.CRViewer.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;
            // 
            // FrmRPT_Viewer1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1280, 960);
            this.Controls.Add(this.CRViewer);
            this.Name = "FrmRPT_Viewer1";
            this.Text = "Report Viewer";
            this.Activated += new System.EventHandler(this.FrmRPT_Viewer1_Activated);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmRPT_Viewer1_FormClosed);
            this.ResumeLayout(false);
        }

        #endregion

        private CrystalDecisions.Windows.Forms.CrystalReportViewer CRViewer;
    }
}