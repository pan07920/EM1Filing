using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EM1Filing
{
    public partial class EM1FinalFile : Form
    {
        private string em1file = @"F:\raw\O_EM1L.txt";

        public EM1FinalFile()
        {
            InitializeComponent();
        }

        private void EM1FinalFile_Load(object sender, EventArgs e)
        {
            textBoxFile.Text = em1file;
            FileInfo fi = new FileInfo(em1file);
            textBoxTimeStamp.Text = fi.LastWriteTime.ToString();
           
        }


        private void buttonLoad_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            DataSet ds = new DataSet();
            try
            {

                string FileName = em1file;
                var dt = new DataTable();

                dt.Columns.Add("Ticker", typeof(String));
                dt.Columns.Add("CUSIP", typeof(String));
                dt.Columns.Add("Sedol", typeof(String));
                dt.Columns.Add("ISIN", typeof(String));
                dt.Columns.Add("Curr", typeof(String));
                dt.Columns.Add("Weight", typeof(decimal));

                string[] lines = System.IO.File.ReadAllLines(em1file);

                foreach (string line in lines)
                {
                    var cols = line.Split(',');
                    if (cols[0].ToString().ToUpper() != "TICKER")
                    {
                        DataRow dr = dt.NewRow();
                        for (int cIndex = 0; cIndex <6; cIndex++)
                        {
                            dr[cIndex] = cols[cIndex];
                        }

                        dt.Rows.Add(dr);
                    }
            
                }
                

                bindingSource1.DataSource = dt;


                dataGridView1.DataSource = bindingSource1;

                dataGridView1.Columns[1].ValueType = typeof(string);
                dataGridView1.Columns[4].ValueType = typeof(decimal);
                dataGridView1.Columns[4].DefaultCellStyle.Format = "N2";
                textBoxWeightTotal.Text = CellSum().ToString();
                textBoxRowTotal.Text = dataGridView1.Rows.Count.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading txt file: " + ex.ToString());
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }
        //private void buttonLoad_Click(object sender, EventArgs e)
        //{
        //    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
        //    DataSet ds = new DataSet();
        //    try
        //    {

        //        string FileName = em1file;
        //        var dt = new DataTable();
        //        var connString = string.Format("Provider=Microsoft.Jet.OleDb.4.0; Data Source={0};Extended Properties=\"Text;HDR=Yes;IMEX=1;FMT=Delimited\"", Path.GetDirectoryName(FileName));

        //        using (var conn = new OleDbConnection(connString))
        //        {
        //            conn.Open();
        //            var query = "SELECT * FROM [" + Path.GetFileName(FileName) + "]";
        //            using (var adapter = new OleDbDataAdapter(query, conn))
        //            {
        //                adapter.Fill(dt);
        //            }
        //        }

                
                
        //        //dt.Columns.Add("Ticker", typeof(String));
        //        //dt.Columns.Add("CUSIP", typeof(String));
        //        //dt.Columns.Add("Sedol", typeof(String));
        //        //dt.Columns.Add("ISIN", typeof(String));
        //        //dt.Columns.Add("Curr", typeof(String));
        //        //dt.Columns.Add("Weight", typeof(decimal));
                
        //        bindingSource1.DataSource = dt;

            
        //        dataGridView1.DataSource = bindingSource1;

        //        dataGridView1.Columns[1].ValueType = typeof(string);
        //        dataGridView1.Columns[4].ValueType = typeof(decimal);
        //        dataGridView1.Columns[4].DefaultCellStyle.Format = "N2";
        //        textBoxWeightTotal.Text = CellSum().ToString();
        //        textBoxRowTotal.Text = dataGridView1.Rows.Count.ToString();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error reading txt file: " + ex.ToString());
        //    }
        //    System.Windows.Forms.Cursor.Current = Cursors.Default;
        //}

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5)
            {
                textBoxWeightTotal.Text = CellSum().ToString();
                textBoxRowTotal.Text = dataGridView1.Rows.Count.ToString();
            }
            
        }
        private double CellSum()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                double d = 0;
                Double.TryParse(dataGridView1.Rows[i].Cells[5].Value.ToString(), out d);
                sum += d;
            }
            return sum;
        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            textBoxWeightTotal.Text = CellSum().ToString();
            textBoxRowTotal.Text = dataGridView1.Rows.Count.ToString();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double weight;
            Double.TryParse(textBoxWeightTotal.Text, out weight);

            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                double d = 0;
                Double.TryParse(dataGridView1.Rows[i].Cells[5].Value.ToString(), out d);
                dataGridView1.Rows[i].Cells[5].Value = Math.Round(d * 100 / weight, 6);
            }

        }

        private void textBoxWeightTotal_TextChanged(object sender, EventArgs e)
        {
            if (textBoxWeightTotal.Text == "100")
            {
                button1.Enabled = false;
                buttonFinalFile.Enabled = true;
            }
            else
            {
                button1.Enabled = true;
                buttonFinalFile.Enabled = false;
            }

        }

        private void buttonFinalFile_Click(object sender, EventArgs e)
        {
            //Program.Export2Excel(this.dataGridView1, true);
            try
            {
                string fileReport = @"F:\raw\" + DateTime.Now.ToString("yyyyMMdd") + "EM1_final.xlsx";
                if (File.Exists(fileReport))
                    File.Delete(fileReport);

                FileInfo newFile = new FileInfo(fileReport);
                ExcelFinalFile(newFile);
                FileInfo fi = new FileInfo(fileReport);
                if (fi.Exists)
                {
                    System.Diagnostics.Process.Start(fileReport);
                }
                else
                {
                    //file doesn't exist
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Create Excel file failed: " + ex.ToString());
                return;
            }
        }

        private void ExcelFinalFile(FileInfo newFile)
        {

            // ok, we can run the real code of the sample now
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                // uncomment this line if you want the XML written out to the outputDir
                //xlPackage.DebugMode = true; 

                // get handle to the existing worksheet
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("EM1Final");

                if (worksheet != null)
                {
                    worksheet.Cells["A1"].Value = "Ticker";
                    worksheet.Cells["B1"].Value = "CUSIP";
                    worksheet.Cells["C1"].Value = "Sedol";
                    worksheet.Cells["D1"].Value = "ISIN";
                    worksheet.Cells["E1"].Value = "Curr";
                    worksheet.Cells["F1"].Value = "Weight";
                    


                    using (ExcelRange headerrow = worksheet.Cells["A1:F1"])
                    {
                        //headerrow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        // headerrow.Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                        // row.Style.Font.Color.SetColor(Color.White);
                        headerrow.Style.Font.Bold = true;
                        // headerrow.Style.Font.UnderLine = true;
                    }

                    int nRow = dataGridView1.RowCount;
                    for (int m = 0; m < nRow; m++)
                    {
                        for (int j = 0; j < 6; j++) //14 column
                        {
                            worksheet.Cells[m + 2, j + 1].Value = dataGridView1.Rows[m].Cells[j].Value;
                        }

                    }
                    string range = "B2:B" + nRow.ToString();
                    using (ExcelRange rtrnRows = worksheet.Cells[range])
                    {
                        rtrnRows.Style.Numberformat.Format = "@";
                    }

                    range = "F2:F" + nRow.ToString();
                    using (ExcelRange rtrnRows = worksheet.Cells[range])
                    {
                        rtrnRows.Style.Numberformat.Format = "##0.000000";
                    }

                    for (int j = 1; j <= 6; j++)
                    {
                        worksheet.Column(j).AutoFit();

                    }
                    //worksheet.View.FreezePanes(2, 1);
                    
                }

                xlPackage.Save();
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Int32 rowToDelete = dataGridView1.Rows.GetFirstRow(DataGridViewElementStates.Selected);
            dataGridView1.Rows.RemoveAt(rowToDelete);
            dataGridView1.ClearSelection();
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Add this
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                // Can leave these here - doesn't hurt
                dataGridView1.Rows[e.RowIndex].Selected = true;
                dataGridView1.Focus();
            }
        }

        private void textBoxFile_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxTimeStamp_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
