using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using OfficeOpenXml;
namespace EM1Filing
{
    public partial class EM1Trading : Form
    {
        //Test for Git
      
        public EM1Trading()
        {
            InitializeComponent();
        }

        private void EM1Trading_Load(object sender, EventArgs e)
        {
            LoadPurchase();
            LoadSale();

        }

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            //autoload at form loat
            //LoadPurchase();
            //LoadSale();
            //LoadStruckList();
        }
        private void LoadPurchase()
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            try
            {
                string sql = "Exec [jian].[SP_GetEM1Trade] @Code";
             

                SqlDataAdapter da = new SqlDataAdapter(sql, global::EM1Filing.Properties.Settings.Default.dbConnectionString);
                da.SelectCommand.CommandTimeout = 0;

                da.SelectCommand.Parameters.AddWithValue("@Code", 0);

                var dt = new DataTable();
                da.Fill(dt);

                bindingSource1.DataSource = dt;
                dataGridViewPurchase.DataSource = bindingSource1;


                dataGridViewPurchase.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridViewPurchase.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridViewPurchase.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
              
                
                dataGridViewPurchase.Columns[0].ValueType = typeof(int);
                dataGridViewPurchase.Columns[1].ValueType = typeof(string);
                dataGridViewPurchase.Columns[2].ValueType = typeof(string);
                dataGridViewPurchase.Columns[3].ValueType = typeof(string);
                dataGridViewPurchase.Columns[4].ValueType = typeof(int);
                dataGridViewPurchase.Columns[4].DefaultCellStyle.Format = "N0";
                dataGridViewPurchase.Columns[5].ValueType = typeof(decimal);
                dataGridViewPurchase.Columns[5].DefaultCellStyle.Format = "N2";

                dataGridViewPurchase.Columns[6].ValueType = typeof(int);
                dataGridViewPurchase.Columns[6].DefaultCellStyle.Format = "N0";

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading txt file: " + ex.ToString());
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }
        private void LoadSale()
        {

            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;

            try
            {
                string sql = "Exec [jian].[SP_GetEM1Trade] @Code";


                SqlDataAdapter da = new SqlDataAdapter(sql, global::EM1Filing.Properties.Settings.Default.dbConnectionString);
                da.SelectCommand.CommandTimeout = 0;

                da.SelectCommand.Parameters.AddWithValue("@Code", 3);

                var dt = new DataTable();
                da.Fill(dt);

                bindingSource2.DataSource = dt;
                dataGridViewSale.DataSource = bindingSource2;

                dataGridViewSale.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridViewSale.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridViewSale.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                dataGridViewSale.Columns[0].ValueType = typeof(int);
                dataGridViewSale.Columns[1].ValueType = typeof(string);
                dataGridViewSale.Columns[2].ValueType = typeof(string);
                dataGridViewSale.Columns[3].ValueType = typeof(string);
                dataGridViewSale.Columns[4].ValueType = typeof(int);
                dataGridViewSale.Columns[4].DefaultCellStyle.Format = "N0";
                dataGridViewSale.Columns[5].ValueType = typeof(decimal);
                dataGridViewSale.Columns[5].DefaultCellStyle.Format = "N2";

                dataGridViewSale.Columns[6].ValueType = typeof(int);
                dataGridViewSale.Columns[6].DefaultCellStyle.Format = "N0";

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading txt file: " + ex.ToString());
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }
        private void LoadStruckList()
        {
            
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            try
            {
                string fileReport = @"F:\raw\EM1_trades_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                if (File.Exists(fileReport))
                    File.Delete(fileReport);

                FileInfo newFile = new FileInfo(fileReport);
                ExcelTradingBY(newFile);
                //ExcelTradingSL(newFile);
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
        private void ExcelTradingBY(FileInfo newFile)
        {

            // ok, we can run the real code of the sample now
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                // uncomment this line if you want the XML written out to the outputDir
                //xlPackage.DebugMode = true; 

                // get handle to the existing worksheet
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("EM1");

                if (worksheet != null)
                {
                    worksheet.Cells["A1"].Value = "Account";
                    worksheet.Cells["B1"].Value = "Side";
                    worksheet.Cells["C1"].Value = "Ticker";
                    worksheet.Cells["D1"].Value = "Shares";
                    worksheet.Cells["E1"].Value = "Price";
                    worksheet.Cells["F1"].Value = "Value";

                    using (ExcelRange headerrow = worksheet.Cells["A1:F1"])
                    {
                        headerrow.Style.Font.Bold = true;
                    }

                    int nRowPurchase = dataGridViewPurchase.RowCount;

                    for (int m = 0; m < nRowPurchase; m++)
                    {
                        for (int j = 0; j < 6; j++) //6 column
                        {
                            worksheet.Cells[m + 2, j + 1].Value = dataGridViewPurchase.Rows[m].Cells[j+1].Value;
                        }
                    }
                    string formula = "SUM(F2:F" + (nRowPurchase + 1).ToString() + ")";
                    worksheet.Cells[nRowPurchase + 2, 6].Formula = formula;


                    int nRowSale = dataGridViewSale.RowCount;

                    for (int m = 0; m < nRowSale; m++)
                    {
                        for (int j = 0; j < 6; j++) //6 column
                        {
                            worksheet.Cells[nRowPurchase + 2 + m + 2, j + 1].Value = dataGridViewSale.Rows[m].Cells[j + 1].Value;
                        }
                    }
                    formula = "SUM(F" + (nRowPurchase + 4).ToString() + ":F" + (nRowPurchase + 1 + nRowSale + 1).ToString() + ")";
                    worksheet.Cells[nRowPurchase + 2 + nRowSale + 1, 6].Formula = formula;







                    string range = "F2:F" + (nRowPurchase + nRowSale + 3).ToString();
                    using (ExcelRange rtrnRows = worksheet.Cells[range])
                    {
                        rtrnRows.Style.Numberformat.Format = "\"$\"#,##0.00;[Red]\"$\"#,##0.00";
                    }

                    range = "D2:D" + (nRowPurchase + nRowSale + 3).ToString();
                    using (ExcelRange rtrnRows = worksheet.Cells[range])
                    {
                        rtrnRows.Style.Numberformat.Format = "#,##0";
                    }

                    // for (int j = 1; j <= 6; j++)
                    {
                        worksheet.Column(6).AutoFit();

                    }
                    worksheet.Calculate();
                    //range = "F2:F" + nRow.ToString();
                    //using (ExcelRange rtrnRows = worksheet.Cells[range])
                    //{
                    //    rtrnRows.Style.Numberformat.Format = "##0.000000";
                    //}

                    //for (int j = 1; j <= 6; j++)
                    //{
                    //    worksheet.Column(j).AutoFit();

                    //}
                    //worksheet.View.FreezePanes(2, 1);

                }

                xlPackage.Save();
            }
        }

        private void ExcelTradingSL(FileInfo newFile)
        {

            // ok, we can run the real code of the sample now
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                // uncomment this line if you want the XML written out to the outputDir
                //xlPackage.DebugMode = true; 

                // get handle to the existing worksheet
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[0];

                if (worksheet != null)
                {
                    worksheet.Cells["A1"].Value = "Account";
                    worksheet.Cells["B1"].Value = "Side";
                    worksheet.Cells["C1"].Value = "Ticker";
                    worksheet.Cells["D1"].Value = "Shares";
                    worksheet.Cells["E1"].Value = "Price";
                    worksheet.Cells["F1"].Value = "Value";

                    int nRowBY = dataGridViewSale.RowCount;
                    int nRow = dataGridViewSale.RowCount;
                    for (int m = 0; m < nRow; m++)
                    {
                        for (int j = 0; j < 6; j++) //14 column
                        {
                            worksheet.Cells[nRowBY + 3 + m, j + 1].Value = dataGridViewSale.Rows[m].Cells[j].Value;
                        }

                    }
                    //string range = "B2:B" + nRow.ToString();
                    //using (ExcelRange rtrnRows = worksheet.Cells[range])
                    //{
                    //    rtrnRows.Style.Numberformat.Format = "@";
                    //}

                    //range = "F2:F" + nRow.ToString();
                    //using (ExcelRange rtrnRows = worksheet.Cells[range])
                    //{
                    //    rtrnRows.Style.Numberformat.Format = "##0.000000";
                    //}

                    //for (int j = 1; j <= 6; j++)
                    //{
                    //    worksheet.Column(j).AutoFit();

                    //}
                    ////worksheet.View.FreezePanes(2, 1);

                }

                xlPackage.Save();
            }
        }
    }
}
