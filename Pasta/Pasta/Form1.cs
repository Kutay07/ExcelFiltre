using OfficeOpenXml;
using System.Data;
using System.Windows.Forms.DataVisualization.Charting;

namespace Pasta
{

    public partial class Form1 : Form
    {
        List<DataTable> sehirler = new();
        List<DataTable> tablolar = new();
        List<Filtre> filtreler = new();

        public Form1()
        {
            InitializeComponent();
        }

        public static List<DataTable> ReadExcelFile(string filePath, ComboBox comboBox)
        {
            List<DataTable> tableList = new();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Dimension == null)
                    {
                        continue;
                    }
                    var dt = new DataTable(worksheet.Name);
                    bool hasHeader = true;
                    for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
                    {
                        dt.Columns.Add(hasHeader ? worksheet.Cells[1, i].Text : string.Format("Column {0}", i));
                    }
                    var startRow = hasHeader ? 2 : 1;
                    for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                        DataRow row = dt.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Value;
                        }
                    }
                    tableList.Add(dt);
                    comboBox.Items.Add(worksheet.Name.Trim());
                }
            }
            return tableList;
        }

        public static void WriteDataTableToExcel(DataTable table, string filePath)
        {
            using (ExcelPackage package = new())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(table.TableName);

                // Yazdýrýlacak sütun baþlýklarýný yaz
                for (int i = 1; i <= table.Columns.Count; i++)
                {
                    worksheet.Cells[1, i].Value = table.Columns[i - 1].ColumnName;
                }

                // Tablodaki tüm verileri yazdýr
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = table.Rows[row][col].ToString();
                    }
                }

                // Excel dosyasýný kaydet
                FileInfo excelFile = new(filePath);
                package.SaveAs(excelFile);
            }
        }


        // buraya bir tablo ve bir sütun indeksi gelir. çýktý olarak verilen indexteki verileri hangisinden kaç tane 
        // olduðunu sayarak oluþturduðu GRUP LISTESINI verir.
        private static List<Veriler> GrafikVerisiOlustur(DataTable tablo, int hangiSutun)
        {
            List<Veriler> verilerList = new();
            foreach (DataRow rw in tablo.Rows)
            {
                var filtrele = verilerList.Where(veri => veri.metin == rw[hangiSutun].ToString().Trim().ToLower());
                if (filtrele.Any())
                {
                    filtrele.First().sayi += 1;
                }
                else
                {
                    string[] s = rw[hangiSutun].ToString().Trim().Split(',');
                    foreach (string ss in s)
                    {
                        var filtrelee = verilerList.Where(veri => veri.metin == ss.Trim().ToLower());
                        if (filtrelee.Any())
                        {
                            filtrelee.First().sayi += 1;
                        }
                        else
                        {
                            if (ss.ToString() != "")
                            {
                                Veriler v = new(1, ss.Trim().ToLower());
                                verilerList.Add(v);
                            }
                        }
                    }
                }
            }
            return verilerList;
        }

        //gelen veriye göre grafiði oluþtur
        void GrafikDoldur(List<Veriler> veriler)
        {
            chart1.Series.Clear();

            chart1.Series.Add("Series1");

            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            //chart1.Series["Series1"]["PieLabelStyle"] = "Outside";
            foreach (Veriler v in veriler)
            {
                chart1.Series["Series1"].Points.AddXY(v.metin, v.sayi.ToString());
            }
            // ToolTip özelliðini ayarla
            foreach (DataPoint point in chart1.Series["Series1"].Points)
            {
                string dataPoint = point.YValues[0].ToString();
                string dataCategory = point.AxisLabel;
                string message = dataPoint + "baþvuru \n" + dataCategory;

                point.ToolTip = message;
            }

        }

        // excel dosyasý seçim butonu
        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new();
            openFileDialog.Filter = "Excel Dosyalarý (*.xlsx)|*.xlsx";
            openFileDialog.Title = "Bir Excel dosyasý seçin";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                tablolar = ReadExcelFile(openFileDialog.FileName, comboBox1);
                //ExceliIceAktar(openFileDialog.FileName, comboBox1);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var tablo = tablolar.Where(dt => dt.TableName == comboBox1.Text);
            if (tablo.Any())
            {
                // grafik için sütun seçilebilen nesneye sütun isimlerinin eklenmesi(combobox5)
                // filtre eklemeye yarayan nesneye isim ekleme (combobox3)

                comboBox5.Items.Clear();
                comboBox3.Items.Clear();
                foreach (DataColumn dc in tablo.First().Columns)
                {
                    comboBox5.Items.Add(dc.ColumnName);
                    comboBox3.Items.Add(dc.ColumnName);

                }
                dataGridView1.DataSource = tablo.First();
                dataGridView2.DataSource = tablo.First();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            KategoriOlustur((DataTable)dataGridView2.DataSource, comboBox3.SelectedIndex, comboBox4);
        }

        private void KategoriOlustur(DataTable gelenTablo, int sutun, ComboBox comboBox)
        {
            sehirler.Clear();
            comboBox.Items.Clear();

            foreach (DataRow dr in gelenTablo.Rows)
            {
                var filteredDataTables = sehirler.Where(dt => dt.TableName == dr[sutun].ToString().Trim());
                if (filteredDataTables.Any())
                {
                    filteredDataTables.First().ImportRow(dr);
                }
                else
                {
                    DataTable yeniTablo = gelenTablo.Clone();
                    yeniTablo.TableName = dr[sutun].ToString();
                    yeniTablo.Rows.Clear();
                    yeniTablo.ImportRow(dr);
                    sehirler.Add(yeniTablo);
                    comboBox.Items.Add(dr[sutun].ToString());
                }
            }
        }

        private DataTable BirdenFazlaFiltrelenmisTabloOlustur(DataTable gelenTablo)
        {
            DataTable degisenTablo = gelenTablo.Clone();
            foreach (DataRow dr in gelenTablo.Rows)
            {
                degisenTablo.ImportRow(dr);
            }

            for (int i = 0; i < filtreler.Count; i++)
            {
                try
                {
                    DataRow[] satirlar = degisenTablo.Select(filtreler[i].kategori + " LIKE '%" + filtreler[i].veri + "%'");
                    DataTable tempTablo = degisenTablo.Clone();
                    foreach (DataRow dr in satirlar)
                    {
                        tempTablo.ImportRow(dr);
                    }
                    degisenTablo.Rows.Clear();
                    degisenTablo.TableName += "-" + filtreler[i].veri;
                    foreach (DataRow dr in tempTablo.Rows)
                    {
                        degisenTablo.ImportRow(dr);
                    }
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error" + "laaaaps");
                }
            }
            return degisenTablo;
        }
        private void btn_click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string[] secilenfiltre = btn.Text.Split(':');

            int index = 0;

            for (int i = 0; i < filtreler.Count; i++)
            {
                if (filtreler[i].kategori == secilenfiltre[0].Trim() && filtreler[i].veri == secilenfiltre[1].Trim())
                {
                    index = i; break;
                }
            }
            filtreler.Remove(filtreler[index]);

            dataGridView2.DataSource = BirdenFazlaFiltrelenmisTabloOlustur((DataTable)dataGridView1.DataSource);
            flowLayoutPanel1.Controls.Remove(btn);
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            GrafikDoldur(GrafikVerisiOlustur((DataTable)dataGridView2.DataSource, comboBox5.SelectedIndex));
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (Convert.ToInt32(comboBox6.SelectedIndex))
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

        private void chart1_MouseClick(object sender, MouseEventArgs e)
        {
            HitTestResult result = chart1.HitTest(e.X, e.Y);
            if (result.PointIndex >= 0)
            {
                string dataPoint = chart1.Series[0].Points[result.PointIndex].YValues[0].ToString();
                string dataCategory = chart1.Series[0].Points[result.PointIndex].AxisLabel;
                string message = "Data Point: " + dataPoint + "\nData Category: " + dataCategory;
                Filtre f = new Filtre(comboBox5.Text, dataCategory);
                filtreler.Add(f);
                dataGridView3.DataSource = BirdenFazlaFiltrelenmisTabloOlustur((DataTable)dataGridView2.DataSource);
                filtreler.Remove(f);
                MessageBox.Show(message, "Data Information");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            comboBox6.Items.Add("pasta");
            comboBox6.Items.Add("sütun");
            comboBox6.Items.Add("çizgi");
            comboBox6.Items.Add("nokta");
            comboBox6.Items.Add("balon");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(dataGridView1);
            f2.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //Form2 f2 = new Form2(dataGridView3);
            //f2.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = (DataTable)dataGridView1.DataSource;
            filtreler.Clear();
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Filtre filtre = new Filtre(comboBox3.Text, comboBox4.Text);
            bool bayrak = true;
            foreach (Filtre filtre1 in filtreler)
            {
                if (filtre1.kategori == filtre.kategori)
                {
                    if (filtre1.veri != filtre.veri)
                    {
                        MessageBox.Show("her bir kolon için bir filitre birimi girebilirsiniz");
                        comboBox4.Text = filtre1.veri;
                    }
                    bayrak = false;
                }
            }
            if (bayrak)
            {
                Button btn = new Button();
                btn.Size = new Size((" " + filtre.kategori + " : " + filtre.veri).Length * 10, 30);
                btn.Text = " " + filtre.kategori + " : " + filtre.veri;
                btn.Click += new EventHandler(btn_click);
                flowLayoutPanel1.Controls.Add(btn);
                filtreler.Add(filtre);
                dataGridView2.DataSource = BirdenFazlaFiltrelenmisTabloOlustur((DataTable)dataGridView1.DataSource);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            openFileIleExcelKaydet(dataGridView1);
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(dataGridView1);
            f2.ShowDialog();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            openFileIleExcelKaydet(dataGridView2);
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(dataGridView2);
            f2.ShowDialog();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            openFileIleExcelKaydet(dataGridView3);
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(dataGridView3);
            f2.ShowDialog();
        }

        public void openFileIleExcelKaydet(DataGridView dataGrid)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "|*.xlsx";
            saveFileDialog.Title = "dýþa aktar";
            saveFileDialog.FileName = "List";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                WriteDataTableToExcel((DataTable)dataGrid.DataSource, saveFileDialog.FileName);
                MessageBox.Show("Excel beriltilen konuma kaydedildi.");
            }
            else
                MessageBox.Show("Bir Hata ile karþýlaþýldý.");
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

    class Filtre
    {
        public Filtre(string kategori, string veri)
        {
            this.kategori = kategori;
            this.veri = veri;
        }
        public string kategori { get; set; }
        public string veri { get; set; }
    }
}

