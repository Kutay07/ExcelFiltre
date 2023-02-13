using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pasta
{
    
    public partial class Form1 : Form
    {          
        List<DataTable> sehirler = new List<DataTable>();

        DataTable tablo = new DataTable();

        String connectionString = "";
        public Form1()
        {
            InitializeComponent();
        }

        private List<Veriler> GrafikVerisiOlustur(int hangiSutun)
        {
            List<Veriler> verilerList = new List<Veriler>();
            foreach (DataRow rw in tablo.Rows)
            {
                var filtrele = verilerList.Where(veri => veri.metin == rw[hangiSutun].ToString().Trim());
                if (filtrele.Any())
                {
                    filtrele.First().sayi += 1;
                }
                else
                {
                    string[] s = rw[hangiSutun].ToString().Trim().Split(',');
                    foreach(string ss in s)
                    {
                        var filtrelee = verilerList.Where(veri => veri.metin == ss.Trim());
                        if (filtrelee.Any())
                        {
                            filtrelee.First().sayi += 1;
                        }
                        else
                        {
                            if (ss.ToString() != "")
                            {
                                Veriler v = new Veriler(1, ss.Trim());
                                verilerList.Add(v);
                            }
                        }
                    }
                }
            }
            return verilerList;
        }

        private void SutunaGoreYeniTablolarOlustur(DataTable genelTablo, int hangiSutun)
        {
            sehirler.Clear();
            comboBox2.Items.Clear();

            foreach (DataRow dr in genelTablo.Rows)
            {
                var filteredDataTables = sehirler.Where(dt => dt.TableName == dr[hangiSutun].ToString().Trim());
                if (filteredDataTables.Any())
                {
                    filteredDataTables.First().ImportRow(dr);
                }
                else
                {
                    DataTable yeniTablo = genelTablo.Clone();
                    yeniTablo.TableName = dr[hangiSutun].ToString();
                    yeniTablo.Rows.Clear();
                    yeniTablo.ImportRow(dr);
                    sehirler.Add(yeniTablo);
                    comboBox2.Items.Add(dr[hangiSutun].ToString());
                }
            }
        }

        private void ExceliIceAktar(string filePath)
        {
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    MessageBox.Show("Excel dosyasýnda hiçbir sayfa bulunamadý.");
                    return;
                }

                // Tüm Sheet Ýsimlerini Çekme
                DataTable dtSayfaAdlari = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // ComboBox'a Sheet Ýsimlerini Ekleme
                comboBox1.Items.Clear();
                foreach (DataRow row in dtSayfaAdlari.Rows)
                {
                    string sheetName = row["TABLE_NAME"].ToString();
                    comboBox1.Items.Add(sheetName);
                }

                OleDbCommand command = new OleDbCommand("SELECT * FROM [" + comboBox1.Items[0] + "]", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);

                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                comboBox3.Items.Clear();
                comboBox5.Items.Clear();
                foreach (DataColumn c in dataSet.Tables[0].Columns)
                {
                    comboBox3.Items.Add(c.ColumnName);
                    comboBox5.Items.Add(c.ColumnName);
                }
                tablo = dataSet.Tables[0];
                dataGridView1.DataSource = dataSet.Tables[0];
                SutunaGoreYeniTablolarOlustur(dataSet.Tables[0], 0);
                dataGridView2.DataSource = sehirler.First();
                grafikDoldur(dataSet.Tables[0], GrafikVerisiOlustur(6));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dosyalarý (*.xlsx)|*.xlsx";
            openFileDialog.Title = "Bir Excel dosyasý seçin";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ExceliIceAktar(openFileDialog.FileName);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand commandd = new OleDbCommand("SELECT * FROM [" + comboBox1.Text + "]", connection);
                OleDbDataAdapter adapterr = new OleDbDataAdapter(commandd);

                DataSet dataSett = new DataSet();
                adapterr.Fill(dataSett);
                tablo = dataSett.Tables[0];
                dataGridView1.DataSource = dataSett.Tables[0];
                comboBox3.Items.Clear();
                foreach (DataColumn c in dataSett.Tables[0].Columns)
                {
                    comboBox3.Items.Add(c.ColumnName);
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var filteredDataTables = sehirler.Where(dt => dt.TableName == comboBox2.Text);
            if (filteredDataTables.Any())
            {
                dataGridView2.DataSource = filteredDataTables.First();
                tablo = filteredDataTables.First();
                grafikDoldur(filteredDataTables.First(), GrafikVerisiOlustur(6));
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Text = comboBox3.SelectedIndex.ToString();
            SutunaGoreYeniTablolarOlustur(tablo, comboBox3.SelectedIndex);
            if (sehirler.Any())
            {
                dataGridView2.DataSource = sehirler.First();
                grafikDoldur(sehirler.First(), GrafikVerisiOlustur(6));
            }
        }

        void grafikDoldur(DataTable tb, List<Veriler> veriler)
        {
            GrafikVerisiOlustur(6);
            chart1.Series.Clear();
            
            chart1.Series.Add("Series1");

            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            //chart1.Series["Series1"]["PieLabelStyle"] = "Outside";
            foreach (Veriler v in veriler)
            {
                chart1.Series["Series1"].Points.AddXY(v.metin, v.sayi.ToString());
            }
         
        }

        private void chart1_MouseClick(object sender, MouseEventArgs e)
        {
            HitTestResult result = chart1.HitTest(e.X, e.Y);
            if (result.PointIndex >= 0)
            {
                string dataPoint = chart1.Series[0].Points[result.PointIndex].YValues[0].ToString();
                string dataCategory = chart1.Series[0].Points[result.PointIndex].AxisLabel;
                string message = "Data Point: " + dataPoint + "\nData Category: " + dataCategory;
                MessageBox.Show(message, "Data Information");
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            grafikDoldur(sehirler.First(), GrafikVerisiOlustur(comboBox5.SelectedIndex));
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(Convert.ToInt32(comboBox6.SelectedIndex))
            {
                case 0:
                    chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
                    break;
                case 1:
                    chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                    break;
                case 2:
                    chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                    break;
                case 3:
                    chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point; 
                    break;
                case 4:
                    chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bubble;
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox6.Items.Add("pasta");
            comboBox6.Items.Add("sütun");
            comboBox6.Items.Add("çizgi");
            comboBox6.Items.Add("nokta");
            comboBox6.Items.Add("balon");
        }
    }

    class Veriler
    {
        public Veriler(int sayi, string metin)
        {
            this.sayi = sayi;
            this.metin = metin;
        }
        public int sayi { get; set; }
        public string metin { get; set; }
    }
}

