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
            this.btnReset = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.PickDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PickNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ProductCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PickQty = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
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
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
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
            // btnReset
            // 
            resources.ApplyResources(this.btnReset, "btnReset");
            this.btnReset.Name = "btnReset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.AllowDrop = true;
            resources.ApplyResources(this.dateTimePicker1, "dateTimePicker1");
            this.dateTimePicker1.Name = "dateTimePicker1";
            // 
            // dgv
            // 
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PickDate,
            this.PickNo,
            this.ProductCode,
            this.PickQty});
            resources.ApplyResources(this.dgv, "dgv");
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellContentClick);
            // 
            // PickDate
            // 
            this.PickDate.DataPropertyName = "PICKING_DATE";
            resources.ApplyResources(this.PickDate, "PickDate");
            this.PickDate.Name = "PickDate";
            // 
            // PickNo
            // 
            this.PickNo.DataPropertyName = "PICKING_CODE";
            resources.ApplyResources(this.PickNo, "PickNo");
            this.PickNo.Name = "PickNo";
            // 
            // ProductCode
            // 
            this.ProductCode.DataPropertyName = "ITEM_CODE";
            resources.ApplyResources(this.ProductCode, "ProductCode");
            this.ProductCode.Name = "ProductCode";
            // 
            // PickQty
            // 
            this.PickQty.DataPropertyName = "QTY";
            resources.ApplyResources(this.PickQty, "PickQty");
            this.PickQty.Name = "PickQty";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.dataGridView1, "dataGridView1");
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            // 
            // frmSystem
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.lblSelect);
            this.Controls.Add(this.lblUpload);
            this.Controls.Add(this.rtbNotice);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.btnUpload);
            this.Name = "frmSystem";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmSystem_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
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
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn PickDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn PickNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn ProductCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn PickQty;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

