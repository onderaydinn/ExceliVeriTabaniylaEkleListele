using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb; //OFFİCE PROGRAMI İÇİN KULLANILIR
using System.Security.Cryptography;

namespace Excel_Test
{
    public partial class Form1 : Form
    {
        //EXCEL İLE ALAKALI BAĞLANTI DA PROBLEM YAŞARSAN excel data connectivity download yazıp indir!!
        //https://www.connectionstrings.com/excel/ 
        public Form1()
        {
            InitializeComponent();
        }
        //C:\Users\ASUS\Desktop\Yeni Microsoft Excel Çalışma Sayfası.xlsx
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ASUS\Desktop\Software\Project\Excel_Test\1.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES';");
        void listele()
        {
            baglanti.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Sayfa1$]", baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            baglanti.Close();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("insert into [Sayfa1$] (Saat,Ders)values(@p1,@p2)", baglanti); // Excel Sayfa başlıklarına göre saat,ders
            komut.Parameters.AddWithValue("@p1", textBox1.Text);
            komut.Parameters.AddWithValue("@p2", textBox2.Text);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Yeni Ders Bilgisi Eklendi");
            listele();

        }
    }
}
