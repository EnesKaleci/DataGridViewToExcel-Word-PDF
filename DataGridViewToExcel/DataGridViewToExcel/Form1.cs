using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Document = iTextSharp.text.Document;
using PageSize = iTextSharp.text.PageSize;
using PdfWriter = iTextSharp.text.pdf.PdfWriter;
using Element = iTextSharp.text.Element;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

//NuGet Paket Yöneticisi'ni kullanarak iTextSharp ,Microsoft.Office.Interop.Excel ve Microsoft.Office.Interop.Word
//paketlerini indirmeniz gerekiyor. 
//ADO.NET ile Database bağlamanız gerekiyor.
//
namespace DataGridViewToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Sizin Entities ismiyle değiştirin(NorthwindEntities)
        NorthwindEntities ef = new NorthwindEntities();
        
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ef.Customers.ToList();
        }
        private void btnExcel_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1);
        }
        private void ExportToExcel(DataGridView dataGridView)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            try
            {
                int rowCount = dataGridView.Rows.Count;
                int columnCount = dataGridView.Columns.Count;

                // DataGridView'deki başlıkları yazdırma
                for (int j = 0; j < columnCount; j++)
                {
                    worksheet.Cells[1, j + 1] = dataGridView.Columns[j].HeaderText;
                }

                // DataGridView'deki verileri yazdırma
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value?.ToString();
                    }
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;
                saveDialog.RestoreDirectory = true;

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Excel dosyası başarıyla kaydedildi.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                workbook = null;
                excelApp = null;
            }
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.OverwritePrompt = false;
            save.Title = "PDF Dosyaları";
            save.DefaultExt = "pdf";
            save.Filter = "PDF Dosyaları (*.pdf)|*.pdf|Tüm Dosyalar(*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                PdfPTable pdfTable = new PdfPTable(dataGridView1.ColumnCount);

                // Bu alanlarla oynarak tasarımı iyileştirebilirsiniz.
                pdfTable.DefaultCell.Padding = 3; // hücre duvarı ve veri arasında mesafe
                pdfTable.WidthPercentage = 80; // hücre genişliği
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT; // yazı hizalaması
                pdfTable.DefaultCell.BorderWidth = 1; // kenarlık kalınlığı
                // Bu alanlarla oynarak tasarımı iyileştirebilirsiniz.

                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                    cell.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240); // hücre arka plan rengi
                    pdfTable.AddCell(cell);
                }
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.Value != null)
                            {
                                pdfTable.AddCell(cell.Value.ToString());

                            }
                            else
                            {
                                pdfTable.AddCell("");
                            }
                        }
                    }
                }
                catch (NullReferenceException)
                {
                }
                using (FileStream stream = new FileStream(save.FileName + ".pdf", FileMode.Create))
                {
                    Document pdfDoc = new Document(PageSize.A2, 10f, 10f, 10f, 0f);// sayfa boyutu.
                    PdfWriter.GetInstance(pdfDoc, stream);
                    pdfDoc.Open();
                    pdfDoc.Add(pdfTable);
                    pdfDoc.Close();
                    stream.Close();
                }
                

            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Yeni bir Word uygulaması oluştur
            Word.Application wordApp = new Word.Application();

            try
            {
                // Yeni bir belge oluştur
                Word.Document doc = wordApp.Documents.Add();

                //Belgeye bir tablo ekliyoruz
                Word.Table table = doc.Tables.Add(doc.Range(), dataGridView1.Rows.Count + 1, dataGridView1.Columns.Count);

                // DataGridView başlıklarını Word tablosuna kopyala
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    table.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
                }

                // DataGridView verilerini Word tablosuna kopyala
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        // DataGridView hücresinin değerini al
                        object value = dataGridView1.Rows[i].Cells[j].Value;

                        // Eğer hücre değeri null değilse, Word belgesine ekleyin
                        if (value != null)
                        {
                            table.Cell(i + 1, j + 1).Range.Text = value.ToString();
                        }
                        else
                        {
                            table.Cell(i + 1, j + 1).Range.Text = ""; // Null değer için boş bir dize ekle
                        }
                    }
                }

                // Belgeyi kaydet

                SaveFileDialog saveDiaglog = new SaveFileDialog();
                saveDiaglog.Filter = "Word Files|*.docx";
                if (saveDiaglog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDiaglog.FileName;
                    doc.SaveAs2(filePath);


                }

                doc.Close();
            }
            catch (Exception ex)
            {
                // Hata kontrol
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                //Word uygulamasını kapat
                wordApp.Quit();
            }
        }

    }
}

