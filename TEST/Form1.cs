using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using ClosedXML.Excel;
using Syncfusion.XlsIO;
using System.IO;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Ganss.Excel;

namespace TEST
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void X_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void Form1_Load(object sender, EventArgs e)
        {
            


            //Connection Object 
            SQLiteConnection con = new SQLiteConnection(@"Data source = C:\\valve-data.db;");

            //Command Object 

            string query = "SELECT* from Sheet1";
            SQLiteCommand cmd = new SQLiteCommand(query, con);


            //Datatable

            System.Data.DataTable dt = new System.Data.DataTable();
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            adapter.Fill(dt);

            


            dataGridView1.DataSource = dt;
           

            

            chart1.Series["J-1"].XValueMember = "Time";
            chart1.Series["J-1"].YValueMembers = "J-1";
            chart1.Series["J-2"].XValueMember = "Time";
            chart1.Series["J-2"].YValueMembers = "J-2";
            chart1.Series["J-4"].XValueMember = "Time";
            chart1.Series["J-4"].YValueMembers = "J-4";
            chart1.Series["J-6"].XValueMember = "Time";
            chart1.Series["J-6"].YValueMembers = "J-6";
            chart1.Series["SDO-2"].XValueMember = "Time";
            chart1.Series["SDO-2"].YValueMembers = "SDO-2";

            chart1.DataSource = dt;




        }

        private void button20_Click_1(object sender, EventArgs e)
        {
          
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
           
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
           
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            
            app.Visible = true;
            
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
           
            worksheet.Name = "Exported from Chart";
            
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
           
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
             
            workbook.SaveAs("c:\\Chart.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            app.Quit();




        }

        private void button13_Click(object sender, EventArgs e)
        {
            string chartImage = Environment.CurrentDirectory + "\\Chart.png";
            chart1.SaveImage(chartImage, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
            MessageBox.Show("Saved The Chart");




        }

        private void button21_Click(object sender, EventArgs e)
        {
           
            int yMin = int.Parse(textBox1.Text);
            int yMax = int.Parse(textBox2.Text);
            int xMin = int.Parse(textBox6.Text);
            int xMax = int.Parse(textBox5.Text);
            
            
            chart1.ChartAreas["ChartArea1"].AxisX.Minimum = xMin;
            chart1.ChartAreas["ChartArea1"].AxisX.Maximum = xMax;
            chart1.ChartAreas["ChartArea1"].AxisY.Minimum = yMin;
            chart1.ChartAreas["ChartArea1"].AxisY.Maximum = yMax;


            chart1.Series["J-1"].XValueMember = "Time";
            chart1.Series["J-1"].YValueMembers = "J-1";
            chart1.Series["J-2"].XValueMember = "Time";
            chart1.Series["J-2"].YValueMembers = "J-2";
            chart1.Series["J-4"].XValueMember = "Time";
            chart1.Series["J-4"].YValueMembers = "J-4";
            chart1.Series["J-6"].XValueMember = "Time";
            chart1.Series["J-6"].YValueMembers = "J-6";
            chart1.Series["SDO-2"].XValueMember = "Time";
            chart1.Series["SDO-2"].YValueMembers = "SDO-2";




        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {
            
        }
    }
}

