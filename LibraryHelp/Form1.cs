using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;



namespace LibraryHelp
{
    public partial class Form1 : Form
    {
        public int SheetNum = 1;
        public int ColNum = 1;
        public Form1()
        {
            InitializeComponent();
        }

        LoadData LD = new LoadData(); // 객체생성



        // 경로추출
        public string ShowFileOpenDialog()
        {
            //파일오픈창 생성 및 설정
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "파일을 선택하세요";
            ofd.FileName = "test";
            ofd.Filter = "엑셀 파일 (.xlsx, .xlsm,.xlsb) | *.xlsx; *.xlsm; *.xlsb; | 모든 파일 (*.*) | *.*";

            //파일 오픈창 로드
            DialogResult dr = ofd.ShowDialog();

            string seetNumOfString = this.textBox1.Text;
            SheetNum = int.Parse(seetNumOfString);
            string ColNumOfString = this.textBox2.Text;
            ColNum = int.Parse(ColNumOfString);

            //OK버튼 클릭시
            if (dr == DialogResult.OK)
            {
                //File명과 확장자를 가지고 온다.
                string fileName = ofd.SafeFileName;
                //File경로와 File명을 모두 가지고 온다.
                string fileFullName = ofd.FileName;
                //File경로만 가지고 온다.
                string filePath = fileFullName.Replace(fileName, "");

                //Excel 파일 불러오기
                LD.ReadExcelData(fileFullName, SheetNum, ColNum);

                //File경로 + 파일명 리턴
                return fileFullName;
            }
            //취소버튼 클릭시 또는 ESC키로 파일창을 종료 했을경우
            else if (dr == DialogResult.Cancel)
            {
                return "";
            }

            return "";
        }

        private void textBox_1_KeyPress(object sender, KeyPressEventArgs e)
        {

            textBox1.Text = "";
            
            //숫자와  " - "  표시가 아닌 다른문자는 입력되지 않습니다.  || 을사용하여 다른 문자를 포함할 수 있습니다. 

            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back) || e.KeyChar == '-'))
            {
                e.Handled = true;
            }

        }



        private void button2_Click(object sender, EventArgs e)
        {
            ShowFileOpenDialog();
            label3.Text = "다음 작업을 진행하세요";
        }


        private void textBox_2_Mouse(object sender, MouseEventArgs e)
        {
            textBox2.Text = "";
        }

    }

}



