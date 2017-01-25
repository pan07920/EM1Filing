using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
namespace EM1Filing
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainEM1Filing());
        }
        // Hanching Comment 1
        #region Export to Excel
        /// <summary>
        /// Exports a passed datagridview to an Excel worksheet.
        /// If captions is true, grid headers will appear in row 1.
        /// Data will start in row 2.
        /// </summary>
        /// <param name="datagridview"></param>
        /// <param name="captions"></param>
        public static void Export2Excel(DataGridView datagridview, bool captions)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            object objApp_Late;
            object objBook_Late;
            object objBooks_Late;
            object objSheets_Late;
            object objSheet_Late;
            object objRange_Late;
            object[] Parameters;
            string[] headers = new string[datagridview.ColumnCount];
            string[] columns = new string[datagridview.ColumnCount];

            int i = 0;
            int c = 0;
            string col = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int A;
            int B;
            for (c = 0; c < datagridview.ColumnCount; c++)
            {
                headers[c] = datagridview.Rows[0].Cells[c].OwningColumn.Name.ToString();
                A = c / 26;
                B = c % 26 + 1;
                columns[c] = (A == 0 ? "" : col.Substring(A - 1, 1)) + (B == 0 ? "Z" : col.Substring(B - 1, 1));
            }

            try
            {
                // Get the class type and instantiate Excel.
                Type objClassType;
                objClassType = Type.GetTypeFromProgID("Excel.Application");
                objApp_Late = Activator.CreateInstance(objClassType);
                //Get the workbooks collection.
                objBooks_Late = objApp_Late.GetType().InvokeMember("Workbooks",
                BindingFlags.GetProperty, null, objApp_Late, null);
                //Add a new workbook.
                objBook_Late = objBooks_Late.GetType().InvokeMember("Add",
                BindingFlags.InvokeMethod, null, objBooks_Late, null);
                //Get the worksheets collection.
                objSheets_Late = objBook_Late.GetType().InvokeMember("Worksheets",
                BindingFlags.GetProperty, null, objBook_Late, null);
                //Get the first worksheet.
                Parameters = new Object[1];
                Parameters[0] = 1;
                objSheet_Late = objSheets_Late.GetType().InvokeMember("Item",
                BindingFlags.GetProperty, null, objSheets_Late, Parameters);

                if (captions)
                {
                    // Create the headers in the first row of the sheet
                    for (c = 0; c < datagridview.ColumnCount; c++)
                    {
                        //Get a range object that contains cell.
                        Parameters = new Object[2];
                        Parameters[0] = columns[c] + "1";
                        Parameters[1] = Missing.Value;
                        objRange_Late = objSheet_Late.GetType().InvokeMember("Range",
                        BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                        //Write Headers in cell.
                        Parameters = new Object[1];
                        Parameters[0] = headers[c];
                        objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty,
                        null, objRange_Late, Parameters);
                    }
                }

                // Now add the data from the grid to the sheet starting in row 2
                for (i = 0; i < datagridview.RowCount; i++)
                {
                    //Debug.WriteLine(i.ToString());
                    //if(i == 233)
                    //    Debug.WriteLine("here" + i.ToString());
                    for (c = 0; c < datagridview.ColumnCount; c++)
                    {
                        //Get a range object that contains cell.
                        Parameters = new Object[2];
                        Parameters[0] = columns[c] + Convert.ToString(i + 2);
                        Parameters[1] = Missing.Value;
                        objRange_Late = objSheet_Late.GetType().InvokeMember("Range",
                        BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                        //Write Headers in cell.
                        Parameters = new Object[1];
                        if (datagridview.Rows[i].Cells[headers[c]].Value == null)
                            Parameters[0] = "";
                        else
                            Parameters[0] = datagridview.Rows[i].Cells[headers[c]].Value.ToString();
                        objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty,
                        null, objRange_Late, Parameters);
                    }
                }

                //Return control of Excel to the user.
                Parameters = new Object[1];
                Parameters[0] = true;
                objApp_Late.GetType().InvokeMember("Visible", BindingFlags.SetProperty,
                null, objApp_Late, Parameters);
                objApp_Late.GetType().InvokeMember("UserControl", BindingFlags.SetProperty,
                null, objApp_Late, Parameters);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
        }
        #endregion
    }
}
