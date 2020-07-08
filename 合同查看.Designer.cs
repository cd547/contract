namespace 合同管理
{
    partial class 合同查看
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
            this.pdfDocumentViewer1 = new Spire.PdfViewer.Forms.PdfDocumentViewer();
            this.SuspendLayout();
            // 
            // pdfDocumentViewer1
            // 
            this.pdfDocumentViewer1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(229)))), ((int)(((byte)(229)))), ((int)(((byte)(229)))));
            this.pdfDocumentViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pdfDocumentViewer1.Location = new System.Drawing.Point(0, 0);
            this.pdfDocumentViewer1.MultiPagesThreshold = 60;
            this.pdfDocumentViewer1.Name = "pdfDocumentViewer1";
            this.pdfDocumentViewer1.Size = new System.Drawing.Size(807, 648);
            this.pdfDocumentViewer1.TabIndex = 0;
            this.pdfDocumentViewer1.Text = "pdfDocumentViewer1";
            this.pdfDocumentViewer1.Threshold = 60;
            this.pdfDocumentViewer1.ZoomMode = Spire.PdfViewer.Forms.ZoomMode.Default;
            // 
            // 合同查看
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(807, 648);
            this.Controls.Add(this.pdfDocumentViewer1);
            this.Name = "合同查看";
            this.Text = "合同查看";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.合同查看_FormClosing);
            this.Load += new System.EventHandler(this.合同查看_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Spire.PdfViewer.Forms.PdfDocumentViewer pdfDocumentViewer1;
    }
}