using System;
using System.Windows.Forms;
using System.Runtime.InteropServices; // Marshal
using XLS = Microsoft.Office.Interop.Excel; // Alias wg. Überschneidung bei Bezeichnern

namespace TimeCalc
{
    class clsInteropExcel
    {
        public static void WriteValues2Excel(DataGridView gridView)
        {
            try
            {
                string winTitle = "Stunden- und Tagerechner";
                XLS.Application xlApp = new XLS.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel kann nicht gestartet werden!`n`nIst das Programm installiert?", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    xlApp.Visible = true; // dafür sorgen, dass Excel sichbar ist
                    XLS.Workbook wb = xlApp.Workbooks.Add(XLS.XlWBATemplate.xlWBATWorksheet);
                    XLS.Worksheet ws = (XLS.Worksheet)wb.Worksheets[1];
                    try
                    {
                        XLS.Range allCells = ws.Cells; // > 23:59 kennt Excel nicht => alles als Text formatieren
                        allCells.NumberFormat = "@"; // set each cell's format to Text
                        allCells.HorizontalAlignment = XLS.XlHAlign.xlHAlignRight; // reset horizontal alignment to the right
                        int j = 0, i = 0, x = 0;
                        for (j = 0; j < gridView.Columns.Count; j++)
                        {
                            XLS.Range rHeader = (XLS.Range)ws.Cells[1, 1 + j];
                            rHeader.Value2 = gridView.Columns[j].HeaderText;
                        }
                        object[,] values = new object[gridView.Rows.Count, gridView.Columns.Count];
                        foreach (DataGridViewRow row in gridView.Rows)
                        {// alle Werte in einen 2-dimensionalem Array speichern
                            while (x < gridView.Columns.Count)
                            {
                                values[i, x] = row.Cells[x].Value != null ? row.Cells[x].Value : string.Empty;
                                x++;
                            }
                            x = 0; i++; //next row
                        }
                        ws.get_Range("A2", "G" + i.ToString()).Value2 = values; // nur 1x COM-Aufruf => gute Peformance
                        if (!xlApp.UserControl) { xlApp.UserControl = true; } // erlauben, dass Anwender Excel schließt
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                    finally { releaseObject(ws); releaseObject(wb); releaseObject(xlApp); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void releaseObject(object obj)
        {
            try { Marshal.ReleaseComObject(obj); obj = null; }
            catch { obj = null; }
            finally { GC.Collect(); }
        }

    }
}
