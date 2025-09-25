using System;
using System.Data;
using System.Windows.Forms;

namespace PickingSystem_001
{
    public partial class frmSystem : Form
    {
        private Dac dac;
        private HandlerController handlerController;
        private string filePath = string.Empty;
        private string date = string.Empty;

        public frmSystem()
        {
            InitializeComponent();
            dac = new Dac();
            dac.MssqlOpen();
            handlerController = new HandlerController(this);
    }

        private void button1_Click(object sender, EventArgs e)
        {
            // 윈도우 기본 파일 열기 대화상자 
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 파일 첨부 화면 로드 시, 디폴트 경로 
                openFileDialog.InitialDirectory = @"C:\";
                // 확인 버튼 눌렀을 때 
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = openFileDialog.FileName;
                    filePath = txtPath.Text;
                    writeRtbNotice("파일 업로드중...");
                    handlerController.ExcelToDataTable(filePath);
                }
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            // DB 조회
            string strDate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            DataTable dt = handlerController.ReadStoredData(strDate);
            
            // DataGridView 
            dgv.AutoGenerateColumns = false;
            dgv.DataSource = dt;
        }

        public void writeRtbNotice(string str)
        {
            rtbNotice.AppendText(str + Environment.NewLine); // 줄바꿈
        }
        private void frmSystem_FormClosing(object sender, FormClosingEventArgs e)
        {
            // DB 연결 끊기
            dac.MssqlClose();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            rtbNotice.Clear();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
