using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OleDBForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void LoadForm()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\testdatabase.xls;Extended Properties=\"Excel 8.0\"";
                String strSQL = "Select * from Radnik2";
                OleDbCommand newComm = new OleDbCommand();
                newComm.Connection = conn;
                newComm.CommandText = strSQL;
                OleDbDataReader reader;
                conn.Open();
                reader = newComm.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    //Console.WriteLine("Poruka: hello world");
                    //Console.WriteLine(reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString());
                    this.lvRadnici.Items.Add(reader["ID"].ToString());

                    this.lvRadnici.Items[i].SubItems.Add(reader["Ime"].ToString());
                    this.lvRadnici.Items[i].SubItems.Add(reader["Prezime"].ToString());
                    i++;
                }
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Postoji problem: " + e.Message);
            }
        }
        private void lvRadnici_Click(object sender, EventArgs e)
        {
            var item = this.lvRadnici.SelectedItems[0];
            string Ime = item.SubItems[1].Text;
            string Prezime = item.SubItems[2].Text;
            
            this.txtID.Text = item.SubItems[0].Text;
            this.txtIme.Text = Ime;
            this.txtPrezime.Text = Prezime;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.LoadForm();
        }

        private void btnPrikazi_Click(object sender, EventArgs e)
        {
            this.LoadForm();
        }
        // dodavanje zapisa
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string ComputerName = Environment.MachineName;
                //string ConString = @"Data Source=" + ComputerName + ";Initial Catalog=Preduzece;Integrated Security=True";
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\testdatabase.xls;Extended Properties=\"Excel 8.0\"";

                int id = Int32.Parse(txtID.Text);
                string Ime = txtIme.Text;
                string Prezime = txtPrezime.Text;
                string querystring = "INSERT INTO Radnik2 (id,Ime,Prezime)  VALUES (@id,@Ime,@Prezime)";

                conn.Open();
                OleDbCommand cmd = new OleDbCommand(querystring, conn);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.Parameters.AddWithValue("@Ime", Ime);
                cmd.Parameters.AddWithValue("@Prezime", Prezime);
                // izvrsi sql upit
                cmd.ExecuteNonQuery();
                conn.Close();
                // prikazi list view opet
                this.LoadForm();
                MessageBox.Show("Radnik uspesno unet u bazu podataka!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Postoji problem: " + ex.Message); ;
            }
        }
        // izmena zapisa
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string ComputerName = Environment.MachineName;
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\testdatabase.xls;Extended Properties=\"Excel 8.0\"";

                string id = txtID.Text;
                string Ime = txtIme.Text;
                string Prezime = txtPrezime.Text;
                string querystring = "UPDATE Radnik2 SET ID=@id,Ime=@Ime,Prezime=@Prezime WHERE ID=@id";

                conn.Open();
                OleDbCommand cmd = new OleDbCommand(querystring, conn);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.Parameters.AddWithValue("@Ime", Ime);
                cmd.Parameters.AddWithValue("@Prezime", Prezime);
                // izvrsi sql upit
                cmd.ExecuteNonQuery();
                conn.Close();
                // prikazi list view opet
                this.LoadForm();
                MessageBox.Show("Radnik uspesno unet u bazu podataka!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Postoji problem: " + ex.Message); ;
            }
        }
    }
}
