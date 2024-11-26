using System.Collections;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelVTEntegrasyonProjesi
{
    public partial class btnExceldenOku : Form
    {
        SqlConnection baglanti = new SqlConnection(@"Data Source=DESKTOP-PI4NQA8\SQLEXPRESS;Initial Catalog=ProjelerVT;Integrated Security=True;Encrypt=False");
        public btnExceldenOku()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelUygulama = new Excel.Application();
            excelUygulama.Visible = true;
            Excel.Workbook workbook = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workbook.Sheets[1];

            string[] basliklar = { "Personel no ", "Ad", "Soyad", "Semt", "Þehir" };

            Excel.Range range;
            // hangi hücreye yazacaðýný belirlemek için Range oluþturulur...


            for (int i = 0; i < basliklar.Length; i++)
            {

                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
            }
            try
            {
                baglanti.Open();
                string sqlCumlesi = "SELECT PersonelNO, Ad, Soyad, Semt, Sehir FROM Personel";
                SqlCommand sqlKomut = new SqlCommand(sqlCumlesi, baglanti);
                SqlDataReader sdr = sqlKomut.ExecuteReader();

                int satir = 2;  // ilk satýr baþlýk olarak geldiði için ikinci satýrdan baþladýk
                while (sdr.Read())
                {
                    string pno = sdr[0].ToString();
                    string ad = sdr[1].ToString();
                    string soyad = sdr[2].ToString();
                    string semt = sdr[3].ToString();
                    string sehir = sdr[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + pno + " " + ad + " " + soyad + " " + semt + " " + sehir + "\n";

                    range = sayfa1.Cells[satir, 1];
                    range.Value2 = pno;
                    range = sayfa1.Cells[satir, 2];
                    range.Value2 = ad;
                    range = sayfa1.Cells[satir, 3];
                    range.Value2 = soyad;
                    range = sayfa1.Cells[satir, 4];
                    range.Value2 = semt;
                    range = sayfa1.Cells[satir, 5];
                    range.Value2 = sehir;
                    satir++;



                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Sql Query sýrasýnda bir hata oluþtu \n" + ex.ToString());
            }

            finally
            {
                if (baglanti != null)

                    baglanti.Close();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkBook;
            Excel.Worksheet exlWorkSheet;
            Excel.Range range;
            int rCount = 0;
            int cCount = 0;
            exlApp = new Excel.Application();
            exlWorkBook = exlApp.Workbooks.Open("C:\\Users\\Asus\\Desktop\\Kitap1.xlsx");
            exlWorkSheet = (Excel.Worksheet)exlWorkBook.Worksheets.get_Item(1);

            range = exlWorkSheet.UsedRange; // tüm excel hücrelerini seçmek için

            // ilk olarak richTextBox2 içeriðini temizle
            richTextBox2.Clear();

            try
            {
                baglanti.Open();

                // Satýrlarý döngü ile oku
                for (rCount = 2; rCount <= range.Rows.Count; rCount++)
                {
                    ArrayList list = new ArrayList();

                    for (cCount = 1; cCount <= range.Columns.Count; cCount++)
                    {
                        string okunanHucre = Convert.ToString((range.Cells[rCount, cCount] as Excel.Range).Value2);

                        // Eðer hücre boþ deðilse listeye ekle
                        if (!string.IsNullOrEmpty(okunanHucre))
                        {
                            list.Add(okunanHucre);
                        }

                        richTextBox2.Text = richTextBox2.Text + okunanHucre + " ";
                    }

                    // Her satýrdan sonra veritabanýna ekleme yap
                    richTextBox2.Text = richTextBox2.Text + "\n";

                    // Eðer list en az 5 eleman içeriyorsa veritabanýna ekle
                    if (list.Count >= 5)
                    {
                        SqlCommand sqlCommand = new SqlCommand("INSERT INTO Personel (PersonelNo, Ad, Soyad, Semt, Sehir) " +
                            "VALUES (@P1, @P2 , @P3, @P4, @P5)", baglanti);

                        sqlCommand.Parameters.AddWithValue("@P1", list[0]);
                        sqlCommand.Parameters.AddWithValue("@P2", list[1]);
                        sqlCommand.Parameters.AddWithValue("@P3", list[2]);
                        sqlCommand.Parameters.AddWithValue("@P4", list[3]);
                        sqlCommand.Parameters.AddWithValue("@P5", list[4]);

                        sqlCommand.ExecuteNonQuery(); // SQL komutunu çalýþtýr
                    }
                    else
                    {
                        // Eksik veri olduðunda hata mesajý verebilir veya o satýrý atlayabilirsiniz
                        MessageBox.Show("Excel dosyasýndaki bir satýr eksik veriye sahip, bu satýr iþlenmedi.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluþtu: \n" + ex.ToString());
            }
            finally
            {
                if (baglanti != null)
                {
                    baglanti.Close();
                }
            }

            exlApp.Quit();
            ReleaseObject(exlWorkSheet);
            ReleaseObject(exlWorkBook);
            ReleaseObject(exlApp);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }

            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}