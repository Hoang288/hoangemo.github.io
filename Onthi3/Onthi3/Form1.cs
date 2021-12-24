using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Xml;

namespace Onthi3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
        }

        DataSet ds = new DataSet(); // dành cho khi cần đụng tới XML


        int flag = 0;
        int stt = 0;

        public struct Tongketdiem
        {
            public string hoten;
            public string lop;
            public string mssv;
            public double dqt;
            public double dkt;
            public double tongket;
        }
        Tongketdiem[] DiemTK = new Tongketdiem[100];

        private void btnthem_Click(object sender, EventArgs e)
        {
            txthoten.Clear();
            txtmssv.Clear();
            txtlop.Clear();
            txtdqt.Clear();
            txtdkt.Clear();
            flag = 1;
        }

        private void btnluu_Click(object sender, EventArgs e)
        {
            DiemTK[stt].hoten = txthoten.Text;
            DiemTK[stt].lop = txtlop.Text;
            DiemTK[stt].mssv = txtmssv.Text;
            DiemTK[stt].dqt = Convert.ToDouble(txtdqt.Text);
            DiemTK[stt].dkt = Convert.ToDouble(txtdkt.Text);
            DiemTK[stt].tongket = 0.3 * DiemTK[stt].dqt + 0.7 * DiemTK[stt].dkt;
            dataGridView1.Rows.Add(DiemTK[stt].hoten, DiemTK[stt].lop, DiemTK[stt].mssv, DiemTK[stt].dqt, DiemTK[stt].dkt, DiemTK[stt].tongket);
            stt++;
        }

        private void maxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //double max = 0;
            //int count = dataGridView1.Rows.Count;
            //for(int i = 1; i <= count; i++)
            //{
            //    if (max <= DiemTK[i].tongket)
            //    {
            //        max = DiemTK[i].tongket;
            //    }
            //}
            //for(int i = 1; i<= count; i++)
            //{
            //    if(DiemTK[i].tongket == max)
            //    {
            //        txtdanhgia.Text = "Điểm tổng kết cao nhất là: " + max;
            //    }
            //}

            double[] columnData = new double[dataGridView1.Rows.Count];

            columnData = (from DataGridViewRow row in dataGridView1.Rows
                          where row.Cells[5].FormattedValue.ToString() != string.Empty
                          select Convert.ToDouble(row.Cells[5].FormattedValue)).ToArray();

            if (columnData != null)
            {
                txtdanhgia.Text = "Điểm tổng kết cao nhất là: "+ columnData.Max().ToString();
            }
            else
            {
                txtdanhgia.Text = "-1";
            }
        }

        private void minToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double[] columnData = new double[dataGridView1.Rows.Count];

            columnData = (from DataGridViewRow row in dataGridView1.Rows
                          where row.Cells[5].FormattedValue.ToString() != string.Empty
                          select Convert.ToDouble(row.Cells[5].FormattedValue)).ToArray();

            if (columnData != null)
            {
                txtdanhgia.Text = "Điểm tổng kết thấp nhất là: " + columnData.Min().ToString();
            }
            else
            {
                txtdanhgia.Text = "-1";
            }
        }

        private void notepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text file|*.txt|Word file|*.doc|All file|*.*";
            saveFileDialog1.ShowDialog();
            string name;
            name = saveFileDialog1.FileName;
            StreamWriter fw = new StreamWriter(name); //mở truy xuất file
            fw.WriteLine(dataGridView1.Rows.Count); //nội dung file
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                fw.WriteLine(dataGridView1.Rows[i].Cells[0].Value);
                fw.WriteLine(dataGridView1.Rows[i].Cells[1].Value);
                fw.WriteLine(dataGridView1.Rows[i].Cells[2].Value);
                fw.WriteLine(dataGridView1.Rows[i].Cells[3].Value);
                fw.WriteLine(dataGridView1.Rows[i].Cells[4].Value);
                fw.WriteLine(dataGridView1.Rows[i].Cells[5].Value);
            }
            fw.Close(); //đóng truy xuất file

        }

        private void notepadToolStripMenuItem1_Click(object sender, EventArgs e)
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
                string Lop;
                string MSSV;
                string DQT;
                string DKT;
                string DTK;
                Hoten = fr.ReadLine();
                Lop = fr.ReadLine();
                MSSV = fr.ReadLine();
                DQT = fr.ReadLine();
                DKT = fr.ReadLine();
                DTK = fr.ReadLine();
                dataGridView1.Rows.Add(Hoten,Lop, MSSV, DQT, DKT, DTK);
            }
            fr.Close(); //đóng truy xuất file
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

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {

                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        xcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                xcelApp.Columns.AutoFit();
                xcelApp.Visible = true;
            }
        }

        private void excelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            string ExcelPath = "";
            DataTable ExcelTable = new DataTable();
            using (OpenFileDialog openFile = new OpenFileDialog() { Filter = "Excel File|*.xls;*.xlsx;*" })
            {
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    ExcelPath = openFile.FileName;
                }
            }
            ExcelTable = LoadExcelSheet(ExcelPath, "Sheet1");
            dataGridView1.DataSource = ExcelTable;
        }

        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            var dt = new DataTable();
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (column.Visible)
                {
                    dt.Columns.Add();
                }
            }
            object[] cellValues = new object[dataGridView1.Columns.Count];
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                dt.Rows.Add(cellValues);
            }
            return dt;

        }

        private void xMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = GetDataTableFromDGV(dataGridView1);
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.WriteXml(File.OpenWrite(@"d:\TestDulieu.xml"));
            MessageBox.Show("Đã lưu");
        }

        private void xMLToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "XML|*.xml";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    XmlReader xmlFile = XmlReader.Create(ofd.FileName, new XmlReaderSettings());
                    ds.ReadXml(xmlFile);
                    dataGridView1.DataSource = ds.Tables[0].DefaultView;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
        }
    }
}
