namespace PickingSystem_001
{
    partial class frmSystem
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSystem));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnSelect = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.rtbNotice = new System.Windows.Forms.RichTextBox();
            this.lblUpload = new System.Windows.Forms.Label();
            this.lblSelect = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnUpload
            // 
            resources.ApplyResources(this.btnUpload, "btnUpload");
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSelect
            // 
            resources.ApplyResources(this.btnSelect, "btnSelect");
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.UseVisualStyleBackColor = true;
            // 
            // txtPath
            // 
            resources.ApplyResources(this.txtPath, "txtPath");
            this.txtPath.Name = "txtPath";
            this.txtPath.ReadOnly = true;
            // 
            // rtbNotice
            // 
            resources.ApplyResources(this.rtbNotice, "rtbNotice");
            this.rtbNotice.Name = "rtbNotice";
            // 
            // lblUpload
            // 
            resources.ApplyResources(this.lblUpload, "lblUpload");
            this.lblUpload.Name = "lblUpload";
            // 
            // lblSelect
            // 
            resources.ApplyResources(this.lblSelect, "lblSelect");
            this.lblSelect.Name = "lblSelect";
            // 
            // frmSystem
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblSelect);
            this.Controls.Add(this.lblUpload);
            this.Controls.Add(this.rtbNotice);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.btnUpload);
            this.Name = "frmSystem";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmSystem_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.RichTextBox rtbNotice;
        private System.Windows.Forms.Label lblUpload;
        private System.Windows.Forms.Label lblSelect;
    }
}

