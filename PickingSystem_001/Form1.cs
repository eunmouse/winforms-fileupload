using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PickingSystem_001
{
    public partial class frmSystem : Form
    {
        private DAC dac;
        private string filePath = string.Empty; // 빈문자열 ""

        public frmSystem()
        {
            InitializeComponent();
            dac = new DAC();
            // DB 연결
            dac.MssqlOpen();
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
                    writeRtbNotice("파일 업로드중... 시간 꽤 소요됨...");
                    HandlerController handlerController = new HandlerController(this);
                    handlerController.ReadExcelSheet(filePath);
                }
            }
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
    }
}
