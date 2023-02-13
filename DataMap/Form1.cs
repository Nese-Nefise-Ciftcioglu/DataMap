using System;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Threading.Tasks;
namespace DataMap
{
    public partial class Form1 : Form
    {
        public string location;
        public string value;
        public static SqlConnection conn = new SqlConnection(@"Data Source=" + Environment.MachineName + @"\SQLEXPRESS" + " ;Initial Catalog= test_db ;Integrated Security=SSPI;MultipleActiveResultSets=True");
        public string addressLoc;
        public string nameLoc;
        public string surnameLoc;
        public Form1()
        {
            InitializeComponent();

            //Write the initial values from database to text boxes
            conn.Open();
            List<String> excelBlocks = new List<String>();
            string query = "Select * from ExcelMap";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string output = reader.GetValue(1).ToString();
                excelBlocks.Add(output.ToString()); ;
            }
            textBox1.Text = excelBlocks[0];
            textBox2.Text = excelBlocks[1];
            textBox3.Text = excelBlocks[2];
            conn.Close();
        }

        //--- Set values for ExcelBlock
        private void UpdateMapButton(object sender, EventArgs e)
        {
            //button3.BackColor = Color.DarkGray;
            conn.Open();
            SqlCommand cmdAdrs = new SqlCommand("UPDATE ExcelMap SET ExcelBlock = @adrs Where CloumnName = 'Address' ", conn);
            SqlCommand cmdName = new SqlCommand("UPDATE ExcelMap SET ExcelBlock = @Name Where CloumnName = 'Name' ", conn);
            SqlCommand cmdSurname = new SqlCommand("UPDATE ExcelMap SET ExcelBlock = @Surname Where CloumnName = 'Surname' ", conn);
            //SqlCommand cmd = new SqlCommand("insert into ExcelMap ([Address], [Name], [Surname]) values(@Address,@Name,@Surname)",conn);
            cmdAdrs.Parameters.AddWithValue("@adrs", (addressLoc));
            cmdName.Parameters.AddWithValue("@Name", (nameLoc));
            cmdSurname.Parameters.AddWithValue("@Surname", (surnameLoc));
            cmdAdrs.ExecuteNonQuery();
            cmdName.ExecuteNonQuery();
            cmdSurname.ExecuteNonQuery();
            conn.Close();
            //MessageBox.Show("Mappings saved successfully");
            notificationLabel.Text = String.Format("Mappings saved successfully");

        }


        private void addressBox(object sender, EventArgs e)
        {
            addressLoc = textBox1.Text;
        }

        private void nameBox(object sender, EventArgs e)
        {
            nameLoc = textBox2.Text;

        }

        private void surnameBox(object sender, EventArgs e)
        {
            surnameLoc = textBox3.Text;
        }

        private void SaveButton(object sender, EventArgs e)
        {
            conn.Open();
            List<String> excelBlocks = new List<String>();
            string query = "Select * from ExcelMap";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string output = reader.GetValue(1).ToString();
                excelBlocks.Add(output.ToString());
                //MessageBox.Show(output);               
            }
            foreach (String i in excelBlocks) { Console.WriteLine(i.ToString()); }

            List<String> info = new List<String>();
            string queryInfo = "Select * from Info";
            SqlCommand cmdInfo = new SqlCommand(queryInfo, conn);
            SqlDataReader readerInfo = cmdInfo.ExecuteReader();
            while (readerInfo.Read())
            {
                info.Add(readerInfo.GetValue(0).ToString());
                info.Add(readerInfo.GetValue(1).ToString());
                info.Add(readerInfo.GetValue(2).ToString());

            }
            conn.Close();

            //---Write to excel
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.DisplayAlerts = false;
            string fileName = "C:/Users/pc/source/repos/DataMap/DataMap/ExcelDnm.xlsx";
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(fileName);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

            worksheet.Range[excelBlocks[0]].Value = info[0].ToString();
            worksheet.Range[excelBlocks[1]].Value = info[1].ToString();
            worksheet.Range[excelBlocks[2]].Value = info[2].ToString();

            workbook.SaveAs(Filename: @"C:\Users\pc\source\repos\DataMap\DataMap\ExcelDnm.xlsx", AccessMode: XlSaveAsAccessMode.xlNoChange);
            //workbook.Close();
            //app.Quit();

            /*Dictionary<string, string> dictionary = new Dictionary<string, string>();
            dictionary.Add("Address", excelBlocks[0].ToString());
            dictionary.Add("Name", excelBlocks[1].ToString());
            dictionary.Add("Surname", excelBlocks[2].ToString());

            string addressMap = dictionary["Address"].ToString();
            string nameMap = dictionary["Name"].ToString();
            string surnameMap = dictionary["Surname"].ToString();
            Console.WriteLine(addressMap,surnameMap,nameMap);


            foreach (KeyValuePair<string, string> kvp in dictionary)
            {
                Console.WriteLine(kvp.Key, kvp.Value.ToString());
            } */
        }

        //--Make the window draggable
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HTCAPTION = 0x2;
        [DllImport("User32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void CloseButton(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void MinimizeButton(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
    }


}

