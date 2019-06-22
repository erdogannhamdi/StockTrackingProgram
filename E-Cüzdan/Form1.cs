using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace E_Cüzdan
{
    public partial class Form1 : Form
    {
           
        public Form1()
        {
            InitializeComponent();
        }
       // SqlConnection baglan = new SqlConnection(@"Data Data Source=(LocalDB)\MSSQLLocalDB;Integrated Security=True;AttachDbFileName='|DataDirectory|\stok.mdf'");
        SqlConnection baglan = new SqlConnection(@"Data Source=HAMDI;Initial Catalog=stok;Integrated Security=True");
        DataTable tablo = new DataTable();

        void excel()
        {
            int satir = 1, sutun = 1, i, j;
            Excel.Application uygulama = new Excel.Application();
            uygulama.Visible = true;
            object Missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Workbook kitap = uygulama.Workbooks.Add(Missing);
            Microsoft.Office.Interop.Excel.Worksheet sayfa1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Range alan = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[satir, sutun + i];
                alan.Value2 = dataGridView1.Columns[i].HeaderText;
                alan=sayfa1.get_Range("a1", "f1");
                alan.EntireRow.Font.Size = 12;
                alan.EntireRow.Font.Bold = true;
            }
            satir++;
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range alan = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[satir + i, sutun + j];
                    alan.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    alan.Select();
                }
            }
        }

        public void listele()
        {
            tablo.Clear();
            baglan.Open();
            SqlCommand komut = new SqlCommand("select *from kayit", baglan);
            SqlDataReader oku = komut.ExecuteReader();
            while (oku.Read())
            {
                tablo.Rows.Add(oku["ad"].ToString(), oku["model"].ToString(), oku["seri"].ToString(), oku["adet"].ToString(), oku["tarih"].ToString(), oku["kisi"].ToString());
                dataGridView1.DataSource = tablo;
            }
            baglan.Close();            
        }
            

        private void Form1_Load(object sender, EventArgs e)
        {                    
            textTarih.Text = DateTime.Now.ToShortDateString();
            textTarih.Enabled = false;
            tablo.Columns.Add("Ad", typeof(string));
            tablo.Columns.Add("Model", typeof(string));
            tablo.Columns.Add("Seri No", typeof(string));
            tablo.Columns.Add("Adet", typeof(string));
            tablo.Columns.Add("Tarih", typeof(string));
            tablo.Columns.Add("Kişi", typeof(string));
            dataGridView1.DataSource = tablo;
            listele();
            //this.FormBorderStyle = None; // Form'un kenarlıklarını kadırır.
        }

        private void textAd_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textAd.Text == "" )
            {
                pictureUyari1.Visible = true;
                label9.Visible = true;
                label9.Text = "Lütfen zorunlu yerleri doldurunuz !";
            }
            if (textModel.Text == "")
            {
                pictureUyari2.Visible = true;
                label9.Visible = true;
                label9.Text = "Lütfen zorunlu yerleri doldurunuz !";
            }
            if (textKisi.Text == "")
            {
                pictureUyari3.Visible = true;
                label9.Visible = true;
                label9.Text = "Lütfen zorunlu yerleri doldurunuz !";
            }
            else
            {
                baglan.Open();
                SqlCommand komut = new SqlCommand("insert into kayit (ad,model,seri,adet,tarih,kisi,resim) values (@p1,@p2,@p3,@p4,@p5,@p6,@p7)", baglan);
                komut.Parameters.AddWithValue("@p1", textAd.Text.ToString());
                komut.Parameters.AddWithValue("@p2", textModel.Text.ToString());
                komut.Parameters.AddWithValue("@p3", textSerino.Text.ToString());
                komut.Parameters.AddWithValue("@p4", textAdet.Text.ToString());
                komut.Parameters.AddWithValue("@p5", textTarih.Text.ToString());
                komut.Parameters.AddWithValue("@p6", textKisi.Text.ToString());
                komut.Parameters.AddWithValue("@p7", textResim.Text.ToString());
                komut.ExecuteNonQuery();
                baglan.Close();
                MessageBox.Show("Kayıt eklendi.");
            }
            listele();
            textAd.Text = "";
            textModel.Text = "";
            textSerino.Text = "";
            textAdet.Text = "";
            textKisi.Text = "";
            pictureResim.ImageLocation = "C:\\Users\\hamdi\\Documents\\Visual Studio 2012\\Projects\\Stok_Takip\\E-Cüzdan\\bin\\Debug\\foto\\Photo_Add-RoundedWhite-128.png";

        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            pictureResim.ImageLocation = openFileDialog1.FileName;
            pictureResim.Image =null;
            textResim.Text = openFileDialog1.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            baglan.Open();
            SqlCommand komut = new SqlCommand("delete from kayit where seri LIKE @p1", baglan);
            komut.Parameters.AddWithValue("@p1", textSerino.Text);
            komut.ExecuteNonQuery();
            baglan.Close();
            MessageBox.Show("Kayıt silindi !");
            listele();
            textAd.Text = "";
            textModel.Text = "";
            textSerino.Text = "";
            textAdet.Text = "";
            textKisi.Text = "";
            pictureResim.ImageLocation = "C:\\Users\\hamdi\\Documents\\Visual Studio 2012\\Projects\\Stok_Takip\\E-Cüzdan\\bin\\Debug\\foto\\Photo_Add-RoundedWhite-128.png";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            baglan.Open();
            int index = dataGridView1.SelectedCells[0].RowIndex;
            textAd.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();
            textModel.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();
            textSerino.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            textAdet.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            textTarih.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            textKisi.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
            SqlCommand komut = new SqlCommand("select resim from kayit where seri LIKE @p1", baglan);
            komut.Parameters.AddWithValue("@p1", textSerino.Text.ToString());
            SqlDataReader oku = komut.ExecuteReader();
            while (oku.Read())
            {
                pictureResim.ImageLocation = oku["resim"].ToString();
                textResim.Text = oku["resim"].ToString();
            }
            baglan.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            baglan.Open();
            SqlCommand komut = new SqlCommand("update kayit set ad=@p1,model=@p2,kisi=@p3,resim=@p4,adet=@p6 where seri =@p5", baglan);
            komut.Parameters.AddWithValue("@p1", textAd.Text);
            komut.Parameters.AddWithValue("@p2", textModel.Text);
            komut.Parameters.AddWithValue("@p3", textKisi.Text);
            komut.Parameters.AddWithValue("@p4", textResim.Text);
            komut.Parameters.AddWithValue("@p5", textSerino.Text);
            komut.Parameters.AddWithValue("@p6", textAdet.Text);
            komut.ExecuteNonQuery();
            baglan.Close();
            listele();
            textAd.Text = "";
            textModel.Text = "";
            textSerino.Text = "";
            textAdet.Text = "";
            textKisi.Text = "";
            pictureResim.ImageLocation = "C:\\Users\\hamdi\\Documents\\Visual Studio 2012\\Projects\\Stok_Takip\\E-Cüzdan\\bin\\Debug\\foto\\Photo_Add-RoundedWhite-128.png";
        }

        public static string ToTittleCase(String ad)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(ad); //gelen adın baş harfini büyük yapar
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tablo.Clear();
            baglan.Open();
            SqlCommand komut = new SqlCommand("select *from kayit where ad LIKE @p1", baglan);
            komut.Parameters.AddWithValue("@p1", ToTittleCase(textBox1.Text));
            SqlDataReader oku =komut.ExecuteReader();
            
            while (oku.Read())
            {
                tablo.Rows.Add(oku["ad"].ToString(), oku["model"].ToString(), oku["seri"].ToString(), oku["adet"].ToString(), oku["tarih"].ToString(), oku["kisi"].ToString());
                dataGridView1.DataSource = tablo;
            }
            baglan.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if(textBox2.Text =="")
            {
                listele();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult sonuc;
            sonuc=MessageBox.Show("Çıkmak istediğinizden emin misiniz?", "Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (sonuc==DialogResult.Yes)
            {
                this.Hide();
                Application.Exit();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                listele();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
                tablo.Clear();
                baglan.Open();
                SqlCommand komut = new SqlCommand("select *from kayit where model LIKE @p1", baglan);
                komut.Parameters.AddWithValue("@p1", ToTittleCase(textBox2.Text));
                SqlDataReader oku = komut.ExecuteReader();

                while (oku.Read())
                {
                    tablo.Rows.Add(oku["ad"].ToString(), oku["model"].ToString(), oku["seri"].ToString(), oku["adet"].ToString(), oku["tarih"].ToString(), oku["kisi"].ToString());
                    dataGridView1.DataSource = tablo;
                }
                baglan.Close();
            
        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            textTarih.Text = DateTime.Now.ToShortDateString();
            textAd.Text = "";
            textModel.Text = "";
            textSerino.Text = "";
            textAdet.Text = "";
            textKisi.Text = "";
            pictureResim.ImageLocation = "C:\\Users\\hamdi\\Documents\\Visual Studio 2012\\Projects\\Stok_Takip\\E-Cüzdan\\bin\\Debug\\foto\\Photo_Add-RoundedWhite-128.png";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            excel();
        }

        public FormBorderStyle None { get; set; }

        private void button9_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

   
    }
}