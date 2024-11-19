using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace readFromExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            string delimiter = "|";
            string tablename = "medTable";
            DataSet dataset = new DataSet();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "CSV Files (*.csv)|*.csv";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (MessageBox.Show("Are you sure you want to import the data from \n " + openFileDialog1.FileName + "?", "Are you sure?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    dataGridView1.AllowUserToAddRows = false;
                    filename = openFileDialog1.FileName;
                    StreamReader sr = new StreamReader(filename);
                    string csv = File.ReadAllText(openFileDialog1.FileName);
                    dataset.Tables.Add(tablename);
                    dataset.Tables[tablename].Columns.Add("Field 1");
                    dataset.Tables[tablename].Columns.Add("Field 2");
                    dataset.Tables[tablename].Columns.Add("Field 3");
                    dataset.Tables[tablename].Columns.Add("Field 4");
                    dataset.Tables[tablename].Columns.Add("Field 5");
                    dataset.Tables[tablename].Columns.Add("Field 6");
                    string allData = sr.ReadToEnd();
                    string[] rows = allData.Split("\r".ToCharArray());

                    foreach (string r in rows)
                    {
                        string[] items = r.Split(delimiter.ToCharArray());
                        if (r.Length > 1)
                        {
                            foreach (string i in items)
                            {
                                items = items.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                            }
                            dataset.Tables[tablename].Rows.Add(items);
                        }
                    }
                    this.dataGridView1.DataSource = dataset.Tables[0].DefaultView;
                    MessageBox.Show(filename + " was successfully imported. \n Please review all data before sending it to the database.", "Success!", MessageBoxButtons.OK);
                }
                else
                {
                    this.Close();
                }
            }
        }

        public string filename { get; set; }


        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Import_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /*
        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                str = @"INSERT INTO USERSTable (Id,UserName,Password,Type) VALUES ('" + dataGridView1.Rows[i].Cells["Id"].Value + "', '" + dataGridView1.Rows[i].Cells["UserName"].Value + "'," + dataGridView1.Rows[i].Cells["Password"].Value + ",'" + dataGridView1.Rows[i].Cells["Type"].Value + "');";
                try
                {
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlCommand com = new SqlCommand(str, con))
                        {
                            con.Open();
                            com.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }*/
    }
}