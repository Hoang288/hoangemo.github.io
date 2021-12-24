using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO; //sử dụng Dialog thì phải thêm cái IO này
using System.Data.OleDb; //riêng thao tác với Excel thì phải thêm cái này

namespace Ontap
{
    public partial class Form1 : Form
    {
        Class1[] solieu = new Class1[100];
        int stt;
        int flag;
        int index;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            grbthongtin.Enabled = false;
            btnluu.Enabled = false;
            btnsua.Enabled = false;
            btnxoa.Enabled = false;
        }

        static bool IsNumeric(string value)
        {
            try
            {
                double number;
                bool result = double.TryParse(value, out number);
                return result;
            }
            catch (Exception ex) { return false; }
        }

        private void btnthem_Click(object sender, EventArgs e)
        {
            flag = 10;
            grbthongtin.Enabled = true;
            btnluu.Enabled = true;
            txthoten.Text = "";
            txtmssv.Text = "";
            txtdqt.Text = "";
            txtdkt.Text = "";
        }

        private void btnluu_Click(object sender, EventArgs e)
        {
            stt++;
            solieu[stt] = new Class1();
            if (txthoten.Text == "" || txtmssv.Text == "" || txtdqt.Text == "" || txtdkt.Text == "")
            {
                lblnote.Text = "Chú ý: Hãy nhập đủ";
                lblnote.ForeColor = Color.Red;
            }
            else if (IsNumeric(txthoten.Text) == true| IsNumeric(txtmssv.Text) == false | IsNumeric(txtdqt.Text) == false | IsNumeric(txtdkt.Text) == false)
            {
                MessageBox.Show("Vui lòng nhập lại đúng kiểu dữ liệu");
            }
            else if (flag == 10)
            {
                //stt++;
                //solieu[stt] = new Class1();
                solieu[stt].Hoten = txthoten.Text;
                solieu[stt].Mssv = txtmssv.Text;
                solieu[stt].Dqt = Convert.ToDouble(txtdqt.Text);
                solieu[stt].Dkt = Convert.ToDouble(txtdkt.Text);
                dgvthongke.Rows.Add(txthoten.Text, txtmssv.Text, txtdqt.Text, txtdkt.Text, solieu[stt].Tongket());
                lblnote.Text = "Chú ý:";
                lblnote.ForeColor = Color.Black;
            }
            else if (flag == 20)
            {
                solieu[index].Hoten = txthoten.Text;
                solieu[index].Mssv = txtmssv.Text;
                solieu[index].Dqt = Convert.ToDouble(txtdqt.Text);
                solieu[index].Dkt = Convert.ToDouble(txtdkt.Text);
                dgvthongke.Rows[index].Cells[0].Value = solieu[index].Hoten;
                dgvthongke.Rows[index].Cells[1].Value = solieu[index].Mssv;
                dgvthongke.Rows[index].Cells[2].Value = solieu[index].Dqt;
                dgvthongke.Rows[index].Cells[3].Value = solieu[index].Dkt;
                dgvthongke.Rows[index].Cells[4].Value = solieu[index].Tongket();
                lblnote.Text = "Chú ý:";
                lblnote.ForeColor = Color.Black;
            }
            btnsua.Enabled = false;
            btnxoa.Enabled = false;
            txthoten.Text = "";
            txtmssv.Text = "";
            txtdqt.Text = "";
            txtdkt.Text = "";
        }

        private void dgvthongke_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            btnsua.Enabled = true;
            btnxoa.Enabled = true;
        }

        private void btnsua_Click(object sender, EventArgs e)
        {
            flag = 20;
            index = dgvthongke.CurrentRow.Index;
            txthoten.Text = Convert.ToString(dgvthongke.Rows[index].Cells[0].Value);
            txtmssv.Text = Convert.ToString(dgvthongke.Rows[index].Cells[1].Value);
            txtdqt.Text = Convert.ToString(dgvthongke.Rows[index].Cells[2].Value);
            txtdkt.Text = Convert.ToString(dgvthongke.Rows[index].Cells[3].Value);
            btnsua.Enabled = false;
        }

        private void dgvthongke_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnsua.Enabled = true;
            btnxoa.Enabled = true;
        }

        private void btnxoa_Click(object sender, EventArgs e)
        {
            int row;
            row = dgvthongke.CurrentRow.Index;
            if (row >= 0)
            {

                dgvthongke.Rows.RemoveAt(row);
                lblnote.Text = "Chú ý:";
                lblnote.ForeColor = Color.Black;
                btnxoa.Enabled = false;
            }
            else
            {
                lblnote.Text = "Chú ý: Chọn hàng cần xóa";
                lblnote.ForeColor = Color.Red;
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text file|*.txt|Word file|*.doc|All file|*.*";
            saveFileDialog1.ShowDialog();
            string name;
            name = saveFileDialog1.FileName;
            StreamWriter fw = new StreamWriter(name); //mở truy xuất file
            fw.WriteLine(dgvthongke.Rows.Count); //nội dung file
            for (int i = 0; i < dgvthongke.Rows.Count; i++)
            {
                fw.WriteLine(dgvthongke.Rows[i].Cells[0].Value);
                fw.WriteLine(dgvthongke.Rows[i].Cells[1].Value);
                fw.WriteLine(dgvthongke.Rows[i].Cells[2].Value);
                fw.WriteLine(dgvthongke.Rows[i].Cells[3].Value);
                fw.WriteLine(dgvthongke.Rows[i].Cells[4].Value);
            }
            fw.Close(); //đóng truy xuất file


           
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Text file|*.txt|Word file|*.doc|All file|*.*";
            openFileDialog1.ShowDialog();
            string name;
            name = openFileDialog1.FileName;
            StreamReader fr = new StreamReader(name);
            string strTemp;
            strTemp = fr.ReadLine(); //đọc nội dung file
            int soluong;
            soluong = Convert.ToInt32(strTemp);
            for (int i = 0; i < soluong; i++)
            {
                string Hoten;
                string MSSV;
                string DQT;
                string DKT;
                string DTK;
                Hoten = fr.ReadLine();
                MSSV = fr.ReadLine();
                DQT = fr.ReadLine(); 
                DKT = fr.ReadLine();
                DTK = fr.ReadLine();
                dgvthongke.Rows.Add(Hoten, MSSV, DQT, DKT, DTK);
            }
            fr.Close(); //đóng truy xuất file

        }

        private void mởFileExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
        
            string ExcelPath = "";
            DataTable ExcelTable = new DataTable();
            using (OpenFileDialog openFile = new OpenFileDialog() { Filter = "Excel File|*.xls;*.xlsx;*" })
            {
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    ExcelPath = openFile.FileName;
                }
            }
            dgvthongke.ClearSelection();
            ExcelTable = LoadExcelSheet(ExcelPath, "Sheet1");
            dgvthongke.DataSource = ExcelTable;
        }
        public DataTable LoadExcelSheet(string aExcelFilePath, string aSheetName)
        {
            using (OleDbConnection ExcelConnection = new OleDbConnection()
            { ConnectionString = $@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {aExcelFilePath };
            Extended Properties = 'Excel 12.0 Xml;HDR=YES;'" })
            {
                DataTable SheetData = new DataTable();
                OleDbCommand SQLCommand = new OleDbCommand($"Select * From [{aSheetName}$]", ExcelConnection);
                ExcelConnection.Open();
                ((OleDbDataAdapter)new OleDbDataAdapter(SQLCommand)).Fill(SheetData);
                ExcelConnection.Close();
                return SheetData;
            }
        }

        private void lưuDạngExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgvthongke.Rows.Count > 0)
            {

                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < dgvthongke.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = dgvthongke.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dgvthongke.Rows.Count-1; i++)
                {
                    for (int j = 0; j < dgvthongke.Columns.Count; j++)
                    {
                        xcelApp.Cells[i + 2, j + 1] = dgvthongke.Rows[i].Cells[j].Value.ToString();
                    }
                }
                xcelApp.Columns.AutoFit();
                xcelApp.Visible = true;
            }
        }

        private void txthoten_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
