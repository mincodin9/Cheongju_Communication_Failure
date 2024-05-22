using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;

namespace chungju
{
    public partial class Form1 : Form
    {
        Class1 xls = new Class1();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns.Add("count", "순번");
            dataGridView1.Columns.Add("suNum", "수용가번호(-)");
            dataGridView1.Columns.Add("chg_suNum", "수용가번호");
            dataGridView1.Columns.Add("juBeon", "주번호");
            dataGridView1.Columns.Add("buBeon", "부번호");
            dataGridView1.Columns.Add("chg_juBeon", "변경주번호");
            dataGridView1.Columns.Add("chg_buBeon", "변경부번호");
            dataGridView1.Columns.Add("memo", "사유");
            dataGridView1.Columns.Add("memo", "메모");
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            int selectedRow = e.RowIndex;

            if (selectedRow >= 0 && selectedRow < dataGridView1.Rows.Count)
            {
                string value1 = dataGridView1.Rows[selectedRow].Cells[3].Value.ToString();
                string value2 = dataGridView1.Rows[selectedRow].Cells[4].Value.ToString();

                textBox1.Text = value1;
                textBox2.Text = value2;
            }
        }

        //입력
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //수용가번호/단말주번호/단말부번호 입력
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog oDlg = new OpenFileDialog();
            oDlg.DefaultExt = "xlsx";
            oDlg.Filter = "Select Excel File|*.xlsx";
            oDlg.Multiselect = false;
            DialogResult result = oDlg.ShowDialog();
            if (result != DialogResult.OK) return;

            string xlsFilePath = oDlg.FileName;
            System.Data.DataTable dataTable = readExcel(xlsFilePath);

            // 기존의 DataTable을 사용하여 데이터 읽기
            try
            {
                int cnt = 1;
                foreach (DataRow roww in dataTable.Rows)
                {
                    DataGridViewRow row = new DataGridViewRow();

                    string suNum = roww["수용가번호"].ToString();
                    string juBeon = roww["단말주번호"].ToString();
                    string buBeon = roww["단말부번호"].ToString();

                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells[0].Value = cnt.ToString();

                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells[1].Value = suNum;

                    string edit_suNum = suNum.Replace("-", "");
                    edit_suNum = edit_suNum.Substring(0, edit_suNum.Length - 2);
                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells[2].Value = edit_suNum;

                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells[3].Value = juBeon;

                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells[4].Value = buBeon;

                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells.Add(new DataGridViewTextBoxCell());
                    row.Cells.Add(new DataGridViewTextBoxCell());

                    dataGridView1.Rows.Add(row);
                    cnt++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("실패: " + ex.Message);
            }


        }

        //단말주번호/단말부번호 변환
        private void button2_Click(object sender, EventArgs e)
        {
            string line1 = textBox3.Text;
            string line2 = textBox4.Text;

            line1 = line1.Trim();
            line2 = line2.Trim();
            line2 = line2.Replace("-", "");

            line1 = "0" + line1;
            line2 = "000000" + line2;

            int selectedRow = dataGridView1.CurrentCell.RowIndex;
            if (selectedRow >= 0 && selectedRow < dataGridView1.Rows.Count)
            {
                dataGridView1.Rows[selectedRow].Cells[5].Value = line1;
                dataGridView1.Rows[selectedRow].Cells[6].Value = line2;

                if (dataGridView1.Rows[selectedRow].Cells[5].Value.Equals(dataGridView1.Rows[selectedRow].Cells[3].Value))
                {
                    dataGridView1.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Green;
                }
                else
                {
                    dataGridView1.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Yellow;
                }
                if (dataGridView1.Rows[selectedRow].Cells[4].Value.Equals(dataGridView1.Rows[selectedRow].Cells[6].Value))
                {
                    dataGridView1.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Green;
                }
                else
                {
                    dataGridView1.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }

            textBox3.Text = string.Empty;
            textBox4.Text = string.Empty;


        }

        //메모입력
        private void button3_Click(object sender, EventArgs e)
        {
            string reason = textBox3.Text;
            string memo = textBox4.Text;
            reason = reason.Trim();
            memo = memo.Trim();


            int selectedRow = dataGridView1.CurrentCell.RowIndex;
            if (selectedRow >= 0 && selectedRow < dataGridView1.Rows.Count)
            {
                dataGridView1.Rows[selectedRow].Cells[7].Value = reason;
                dataGridView1.Rows[selectedRow].Cells[8].Value = memo;
            }

            textBox3.Text = string.Empty;
            textBox4.Text = string.Empty;
        }

        //엑셀 저장
        private void button4_Click(object sender, EventArgs e)
        {
            string filepath = "D:\\aaaaall\\ㅇㅅㅁ\\청주통신장애\\abcd.xlsx";
            SaveToExcel(dataGridView1, filepath);
        }


        private System.Data.DataTable readExcel(string xlsFilePath)
        {
            xls.ExcelConnection(xlsFilePath);

            string sheet = "Sheet1$";
            string query = string.Format("Select * from [{0}]", sheet);

            try
            {
                // 기존의 DataTable 및 DataSet을 사용하도록 수정
                DataSet dataSet = new DataSet(); // 수정: 새로운 DataSet 생성
                System.Data.DataTable dataTable = xls.XlsDataTable(query, null, dataSet, sheet, 2); // 수정: dataSet 인자 전달
                xls.ExcelClose();
                return dataTable;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + string.Format("Sheet:{0}.File:F{1}", sheet, xlsFilePath), ex);
            }
        }

        private void SaveToExcel(DataGridView dataGridView, string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 또는 LicenseContext.Commercial

                FileInfo fileInfo = new FileInfo(filePath);

                // 파일이 존재하는지 확인하고 존재하면 열어서 처리
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // 워크북에 Sheet1이 이미 있는지 확인하고 있다면 삭제
                    ExcelWorksheet existingWorksheet = package.Workbook.Worksheets["Sheet1"];
                    if (existingWorksheet != null)
                    {
                        package.Workbook.Worksheets.Delete("Sheet1");
                    }

                    // 새 워크시트 추가
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // 헤더 행 추가
                    for (int i = 1; i <= dataGridView.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i].Value = dataGridView.Columns[i - 1].HeaderText;
                    }

                    // 데이터 행 추가
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataGridView.Rows[i].Cells[j].Value;
                        }
                    }

                    package.Save();
                }

                MessageBox.Show("데이터가 성공적으로 저장되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("파일 저장 중 오류가 발생했습니다: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

}
