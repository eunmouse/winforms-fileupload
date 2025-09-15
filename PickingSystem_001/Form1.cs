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
            dac.Dac();
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
                    filePath = openFileDialog.FileName;

                    HandlerController handlerController = new HandlerController();
                    handlerController.ReadExcelSheet(filePath);
                }
            }
        }
    }
}
