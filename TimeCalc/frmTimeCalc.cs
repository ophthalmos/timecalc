using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Reflection; // Assembly

namespace TimeCalc
{
    public partial class FrmTimeCalc : Form
    {
        string decText = "Gesamtsumme als Dezimalzahl: ";
        string oldText = string.Empty;
        bool nothingToSave = true;
        bool numPadDecimal = false;
        bool isFileLoading = false;
        bool modusStunden = true;
        IFormatProvider deCulture = new CultureInfo("de-DE", true);
        Version curVersion = Assembly.GetExecutingAssembly().GetName().Version;
        clsUtilities.DateDiff ddfSum = new clsUtilities.DateDiff();
        int numRows = 9; // IntDigits(numRows + 1) != IntDigits(e.RowIndex + 1)
        string winTitle;
        string currFile = string.Empty;
        List<int> dtEdits = new List<int>();
        StringFormat strFormat; // Used to format the grid rows.
        ArrayList arrColumnLefts = new ArrayList(); // Used to save left coordinates of columns
        ArrayList arrColumnWidths = new ArrayList(); // Used to save column widths
        int iCellHeight = 0; // Used to get/set the datagridview cell height
        int iTotalWidth = 0; //
        int intRow = 0; // Used as counter
        bool bFirstPage = false; // Used to check whether we are printing first page
        bool bNewPage = false; // Used to check whether we are printing a new page
        int iHeaderHeight = 0; // Used for the header height
        bool isInEditMode;
        Regex rgxImportDate = new Regex(@"([012][0-9]|3[01])\.(0[1-9]|1[012])\.(18[0-9]|19[0-9]|20[0-9])\d", RegexOptions.Compiled);
        Regex rgxImportTime = new Regex(@"([01][0-9]|2[0-3])\:([0-5][0-9])", RegexOptions.Compiled);
        Regex rgxValidDate = new Regex(@"^\d{1,2}\.\d{1,2}\.\d{4}$", RegexOptions.Compiled);
        Regex rgxValidTime = new Regex(@"^\d{1,2}\:\d{2}$", RegexOptions.Compiled);

        public FrmTimeCalc()
        {
            InitializeComponent();
            dtEdits.Add(0); // Index der Felder, in denen Datumsangaben möglich sind (z.B. mit "F5")
            dtEdits.Add(1);
            dtEdits.Add(2);
            dtEdits.Add(4);
            dGV.Columns[5].DefaultCellStyle.BackColor = dGV.Columns[6].DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dGV.Columns[5].DefaultCellStyle.Alignment = dGV.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            winTitle = clsUtilities.GetDescription() + " " + new Regex(@"^\d+\.\d+").Match(curVersion.ToString()).Value;
            Text = "Unbenannt - " + winTitle;
            dGV.TopLeftHeaderCell.ToolTipText = "Tabelle";
            dGV.Columns[0].ToolTipText = "Anfangdatum";
            dGV.Columns[1].ToolTipText = "Enddatum";
            dGV.Columns[2].ToolTipText = "Pausenstunden/Fehltage";
            dGV.Columns[3].ToolTipText = "Mit Hilfe des Multiplikators können mehrere Perso-\nnen gleichzeitig bei der Zeiterfassung berücksichtigt\nwerden. Wenn der Multiplikator '0' beträgt, wird die\nZeile von der Saldorechnung ausgeschlossen.";
            dGV.Columns[4].ToolTipText = "Drücken Sie <F5>, um das aktuelle Datum als\nNotiz einzugeben. Anschließend können Sie\nden Wert mit der Minus- oder Plustaste ändern.";
            dGV.Columns[5].ToolTipText = "Zeitraum von Anfang- bis Enddatum";
            dGV.Columns[6].ToolTipText = "Aufsummierung der Zeiten";
        }

        private void DGV_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var dGrid = sender as DataGridView;
            string rowText = (e.RowIndex + 1).ToString() + ". ";
            if (dGV.Rows[e.RowIndex].IsNewRow && IntDigits(numRows + 1) != IntDigits(e.RowIndex + 1))
            {// nur wenn sich die Anzahl der Stellen ändert
                numRows = e.RowIndex;
                SizeF Sz = e.Graphics.MeasureString(rowText, e.InheritedRowStyle.Font, 0, new StringFormat(StringFormatFlags.MeasureTrailingSpaces));
                dGrid.RowHeadersWidth = (int)Sz.Width + 25; // 25
            }
            var centerFormat = new StringFormat(StringFormatFlags.MeasureTrailingSpaces) // default: exclude the space at the end of each line
            {
                Alignment = StringAlignment.Far, // Bei einem Layout mit Ausrichtung von links nach rechts ist die weit entfernte Position rechts.
                LineAlignment = StringAlignment.Center // vertikale Ausrichtung der Zeichenfolge 
            };
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dGrid.RowHeadersWidth, e.RowBounds.Height);
            using (SolidBrush sBrush = new SolidBrush(dGrid.RowHeadersDefaultCellStyle.ForeColor))
            {// the using statement automatically disposes the brush
                e.Graphics.DrawString(rowText, e.InheritedRowStyle.Font, sBrush, headerBounds, centerFormat);
            }
            //if (e.RowIndex % 2 != 0)
            //{// AlternatingCellStyle
            //    dGrid.Rows[e.RowIndex].Cells[0].Style.BackColor = dGrid.Rows[e.RowIndex].Cells[1].Style.BackColor = dGrid.Rows[e.RowIndex].Cells[2].Style.BackColor = dGrid.Rows[e.RowIndex].Cells[3].Style.BackColor = Color.WhiteSmoke;
            ////    if (dGV.Rows[e.RowIndex].Cells[dGV.CurrentCell.ColumnIndex].IsInEditMode)
            //}
        }

        public static int IntDigits(int number)
        {// number = Math.Abs(number); // if negative
            int length = 1;
            while ((number /= 10) >= 1) { length++; }
            return length;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.D | Keys.Control:
                    {
                        AddShortcutToDesktop();
                        return true;
                    }
                case Keys.P | Keys.Control:
                    {
                        ShowPrintDialog();
                        return true;
                    }
                case Keys.E | Keys.Control:
                    {
                        ToolStripMenuItemExcel_Click(null, null);
                        return true;
                    }
                case Keys.P | Keys.Shift | Keys.Control:
                    {
                        ToolStripMenuItemPrintPreview_Click(null, null); // printPreviewDialog.ShowDialog();
                        return true;
                    }
                case Keys.I | Keys.Control:
                    {
                        ImportToolStripMenuItem_Click(null, null);
                        return true;
                    }
                case Keys.W | Keys.Control:
                    {
                        WebseiteToolStripMenuItem_Click(null, null);
                        return true;
                    }
                case Keys.J | Keys.Control:
                case Keys.F5:
                    if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex) && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {
                        UhrzeitDatumToolStripMenuItem_Click(null, null);
                        return true;
                    }
                    else break;
                case Keys.V | Keys.Control:
                    if (ActiveControl == dGV && dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex))
                    {// EditMode ist ausgeschlossen => ActiveControl == null!
                        PasteToolStripMenuItem_Click(null, null);
                        return true;
                    }
                    else break;
                case Keys.F10 | Keys.Shift:
                case Keys.Apps:
                    if (ActiveControl == dGV && dGV.CurrentCell != null)
                    {// EditMode ist ausgeschlossen => ActiveControl == null!
                        DataGridViewCell ccell = dGV.CurrentCell;
                        Rectangle r = ccell.DataGridView.GetCellDisplayRectangle(ccell.ColumnIndex, ccell.RowIndex, false);
                        Point p = new Point(r.X + r.Width / 2, r.Y + r.Height / 2);
                        ShowContextMenu(ccell, ccell.DataGridView, p);
                        return true;
                    }
                    else break;
                case Keys.Add:
                case Keys.Oemplus:
                    if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex) && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {// nicht in letzten beiden Splaten
                        ChangeCellValueByKey(false, true); // Shift, Add
                        return true;
                    }
                    else break;
                case Keys.Subtract:
                case Keys.OemMinus:
                    if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex) && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {// nicht in letzten beiden Splaten
                        ChangeCellValueByKey(false, false); // Shift, Add
                        return true;
                    }
                    else break;
                case Keys.Add | Keys.Shift:
                case Keys.Oemplus | Keys.Shift:
                    if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex) && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {// nicht in letzten beiden Splaten
                        ChangeCellValueByKey(true, true); // Shift, Add
                        return true;
                    }
                    else break;
                case Keys.Subtract | Keys.Shift:
                case Keys.OemMinus | Keys.Shift:
                    if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex) && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {// nicht in letzten beiden Splaten
                        ChangeCellValueByKey(true, false); // Shift, Add
                        return true;
                    }
                    else break;
                case Keys.Return: // case Keys.Enter:
                case Keys.Tab:
                    if (dGV.CurrentCell != null && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {
                        EnterToolStripMenuItem_Click(null, null);
                        return true;
                    }
                    else break;
                case Keys.Delete:
                    if (ActiveControl == dGV && dGV.CurrentCell != null)
                    {// EditMode ist ausgeschlossen => ActiveControl == null!
                        if (dGV.SelectedRows.Count > 0) { DeleteRowToolStripMenuItem_Click(null, null); }
                        else { LöschenToolStripMenuItem_Click(null, null); }
                        return true;
                    }
                    else break; // Delete-Taste hat eine Funktion in EditMode!
                case Keys.Delete | Keys.Shift:
                    if (dGV.CurrentCell != null && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {
                        DeleteRowToolStripMenuItem_Click(null, null);
                    }
                    return true;
                case Keys.Delete | Keys.Control | Keys.Shift:
                    if (dGV.CurrentCell != null && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {
                        AllesLöschenContextMenuItem_Click(null, null);
                    }
                    return true;
                case Keys.Insert | Keys.Shift:
                    if (dGV.CurrentCell != null && (ActiveControl == dGV || dGV.IsCurrentCellInEditMode))
                    {
                        InsertRowToolStripMenuItem_Click(null, null);
                    }
                    return true;
                case Keys.Tab | Keys.Control:
                case Keys.F12:
                    ChangeModus();
                    return true;
                case Keys.S | Keys.Control:
                    SpeichernToolStripMenuItem_Click(null, null);
                    return true;
                case Keys.S | Keys.Shift | Keys.Control:
                    SpeichernUnterToolStripMenuItem_Click(null, null);
                    return true;
                //case Keys.I | Keys.Control:
                //    infoToolStripMenuItem_Click(null, null);
                //    return true;
                case Keys.F1:
                    HelpMessageBoxShow();
                    return true;
                case Keys.Escape:
                    if (dGV.CurrentCell != null && dGV.IsCurrentCellInEditMode)
                    {
                        dGV.EndEdit();
                        dGV.CurrentCell.Selected = true;
                    }
                    else { Close(); }
                    return true;
            }// switch (keyData)
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ChangeCellValueByKey(bool shiftMode, bool addMode)
        {
            DateTime dt = new DateTime();
            if (dGV.IsCurrentCellInEditMode)
            {
                string tbText = ((TextBox)dGV.EditingControl).Text;
                if (modusStunden)
                {// Stundenrechner
                    string dtFormat = dGV.CurrentCell.ColumnIndex < 3 ? "H:mm" : "d.M.yyyy";
                    if (dGV.CurrentCell.ColumnIndex < 2) { dt = clsUtilities.NormalizeTime(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                    else if (dGV.CurrentCell.ColumnIndex == 2) { dt = clsUtilities.NormalizePause(tbText.Length == 0 ? DateTime.Today.ToString(dtFormat) : tbText); }
                    else { dt = clsUtilities.NormalizeDate(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                    if (dt != DateTime.MinValue) { dGV.CurrentCell.Value = dt.ToString(dtFormat); }
                    else { dGV.CurrentCell.Value = null; }
                }
                else
                {// Tagerechner
                    if (dGV.CurrentCell.ColumnIndex == 2)
                    {
                        dGV.CurrentCell.Value = int.TryParse(dGV.CurrentCell.Value.ToString(), out int temp) ? temp : 0;
                    }
                    else
                    {
                        string dtFormat = "d.M.yyyy";
                        if (dGV.CurrentCell.ColumnIndex < 2) { dt = clsUtilities.NormalizeTime(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                        else { dt = clsUtilities.NormalizeDate(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                        if (dt != DateTime.MinValue) { dGV.CurrentCell.Value = dt.ToString(dtFormat); }
                        else { dGV.CurrentCell.Value = null; }
                    }
                }
            }
            else { dGV.BeginEdit(true); }
            try
            {
                object cVal = dGV.CurrentCell.Value;
                {
                    if (modusStunden)
                    {// Stundenrechner
                        string dtFormat = dGV.CurrentCell.ColumnIndex < 3 ? "H:mm" : "d.M.yyyy"; // letzteres für "Notiz"
                        if (dGV.CurrentCell.ColumnIndex == 2)
                        {
                            dt = DateTime.ParseExact(cVal == null || cVal.ToString().Length == 0 ? DateTime.Today.ToString(dtFormat) : cVal.ToString(), dtFormat, deCulture);
                        }
                        else
                        {
                            dt = DateTime.ParseExact(cVal == null || cVal.ToString().Length == 0 ? DateTime.Now.ToString(dtFormat) : cVal.ToString(), dtFormat, deCulture);
                        }
                        if (shiftMode)
                        {
                            if (addMode) { dt = dGV.CurrentCell.ColumnIndex < 3 ? dt.AddHours(1) : dt.AddMonths(1); }
                            else { dt = dGV.CurrentCell.ColumnIndex < 3 ? dt.AddHours(-1) : dt.AddMonths(-1); }
                        }
                        else
                        {
                            if (addMode) { dt = dGV.CurrentCell.ColumnIndex < 3 ? dt.AddMinutes(1) : dt.AddDays(1); }
                            else { dt = dGV.CurrentCell.ColumnIndex < 3 ? dt.AddMinutes(-1) : dt.AddDays(-1); }
                        }
                        dGV.CurrentCell.Value = dt.ToString(dtFormat);
                    }
                    else
                    {// Tagerechner
                        string dtFormat = "d.M.yyyy";
                        if (dGV.CurrentCell.ColumnIndex == 2 && int.TryParse(dGV.CurrentCell.Value.ToString(), out int value))
                            if (shiftMode)
                            {
                                if (addMode) { dGV.CurrentCell.Value = (value + 10).ToString(); }
                                else { dGV.CurrentCell.Value = (value - 10).ToString(); }
                            }
                            else
                            {
                                if (addMode) { dGV.CurrentCell.Value = (value + 1).ToString(); }
                                else { dGV.CurrentCell.Value = (value - 1).ToString(); }
                            }
                        else
                        {
                            dt = DateTime.ParseExact(cVal == null || cVal.ToString().Length == 0 ? DateTime.Now.ToString(dtFormat) : cVal.ToString(), dtFormat, deCulture);
                            if (shiftMode)
                            {
                                if (addMode) { dt = dt.AddMonths(1); }
                                else { dt = dt.AddMonths(-1); }
                            }
                            else
                            {
                                if (addMode) { dt = dt.AddDays(1); }
                                else { dt = dt.AddDays(-1); }
                            }
                            dGV.CurrentCell.Value = dt.ToString(dtFormat);
                        }
                    }
                }
                if (dGV.IsCurrentCellInEditMode)
                {// neuer Wert wird nicht automatisch angezeigt! Code wird nur ausgeführt, wenn das parsen funktioniert hat.
                    ((TextBox)dGV.EditingControl).Text = dGV.CurrentCell.Value.ToString();
                    ((TextBox)dGV.EditingControl).SelectAll();
                }
            }
            catch { Console.Beep(); }
        }

        private void FrmTimeCalc_Load(object sender, EventArgs e)
        {// Width = 444 //MessageBox.Show(dGV.Columns[2].Width.ToString()); // => primär 105
            MinimumSize = new Size(Width, Height);
            MaximumSize = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            string[] args = Environment.GetCommandLineArgs();
            string cmdPath = string.Empty;
            for (int i = 1; i < args.Length; i++)
            {
                cmdPath = string.Concat(cmdPath, " ", args[i]); // Pfad mit Leerzeichen ohne Anführungsstriche
                if (cmdPath.Length > 0 && File.Exists(cmdPath))
                {
                    isFileLoading = true;
                    ReadTextFile(cmdPath);
                    Text = currFile + " - " + winTitle; // Text = Path.GetFileName(currFile) + " - " + winTitle;
                    if (dGV.Rows[0].Cells[0].Value != null)
                    {
                        if (rgxValidDate.Match(dGV.Rows[0].Cells[0].Value.ToString()).Success)
                        { UpdateLastDateColumns(); }
                        else if (rgxValidTime.Match(dGV.Rows[0].Cells[0].Value.ToString()).Success)
                        { modusStunden = true; UpdateLastTimeColumns(); } // modusStunden = true;
                    }
                    break;
                }
            }
            dGV.FirstDisplayedScrollingRowIndex = dGV.RowCount - 1;  //dGV.BeginEdit(true); <= besser NICHT (Cursor fehlt)
            isFileLoading = false;
        }

        private void DGV_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dGV.Tag = dGV.Rows[e.RowIndex].Cells[e.ColumnIndex].Value; //Save old value to datagridview.Tag
            dGV.CurrentCell.Style.ForeColor = Color.Empty;
            dGV.ClearSelection(); //  dGV.CurrentCell.Selected = true;
            toolStripButtonComplete.Enabled = true;
            isInEditMode = true;
        }

        private void DGV_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {// Event ist nicht in den letzten beiden Spalten auslösbar, da die diese READONLY sind. => CellLeave-Ereignis
            CalculateSaldo(dGV.Rows[e.RowIndex].Cells[e.ColumnIndex]);
            string oldString = dGV.Tag == null ? "" : dGV.Tag.ToString();
            string newString = dGV.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null ? "" : dGV.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            if (nothingToSave && newString != oldString)
            {
                ShowCellChangeInTitle();
            }
            toolStripButtonComplete.Enabled = false;
            isInEditMode = false;
        }

        private void DGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {// to detect changes in the DataGridView; CurrentCellDirtyStateChanged to checkbox type columns
            if (!isInEditMode && nothingToSave && dtEdits.Contains(e.ColumnIndex) && e.RowIndex > -1) // bei Programmstart tritt das Ereignis 3x ein => e.RowIndex -1
            {
                ShowCellChangeInTitle();
            }
        }

        private void ShowCellChangeInTitle()
        {
            Text = Regex.Replace(Text, @" - " + winTitle + "$", "* - " + winTitle);
            nothingToSave = false;
        }

        private void CalculateSaldo(DataGridViewCell myCell)
        {// aktualisiert u.a. Spalte 'Saldo'
            DataGridViewCell ccell1 = dGV.Rows[myCell.RowIndex].Cells[0];
            DataGridViewCell ccell2 = dGV.Rows[myCell.RowIndex].Cells[1];
            DataGridViewCell ccell3 = dGV.Rows[myCell.RowIndex].Cells[2];

            if (modusStunden) // Stunden berechnen
            {
                if (myCell.ColumnIndex == 0 && ccell1.Value != null)
                {// 1. Zelle normalisieren
                    DateTime date1 = clsUtilities.NormalizeTime(ccell1.Value.ToString());
                    if (date1 != DateTime.MinValue) { ccell1.Value = date1.ToString("H:mm"); }
                    else { TidyUpOnError(ccell1); }
                }
                else if (myCell.ColumnIndex == 1 && ccell2.Value != null)
                {// 2. Zelle normalisieren
                    DateTime date2 = clsUtilities.NormalizeTime(ccell2.Value.ToString());
                    if (date2 != DateTime.MinValue) { ccell2.Value = date2.ToString("H:mm"); }
                    else { TidyUpOnError(ccell2); }
                }
                else if (myCell.ColumnIndex == 2 && ccell3.Value != null)
                {// 3. Zelle normalisieren
                    DateTime date3 = clsUtilities.NormalizePause(ccell3.Value.ToString());
                    if (date3 != DateTime.MinValue) { ccell3.Value = date3.ToString("H:mm"); }
                    else { TidyUpOnError(ccell3); }
                }
                UpdateLastTimeColumns();
            }
            else // Tage berechnen
            {
                if (myCell.ColumnIndex == 0 && ccell1.Value != null)
                {// 1. Zelle normalisieren
                    DateTime date1 = clsUtilities.NormalizeDate(ccell1.Value.ToString());
                    if (date1 != DateTime.MinValue) { ccell1.Value = date1.ToString("d.M.yyyy"); }
                    else { TidyUpOnError(ccell1); }
                }
                else if (myCell.ColumnIndex == 1 && ccell2.Value != null)
                {// 2. Zelle normalisieren
                    DateTime date2 = clsUtilities.NormalizeDate(ccell2.Value.ToString());
                    if (date2 != DateTime.MinValue) { ccell2.Value = date2.ToString("d.M.yyyy"); }
                    else { TidyUpOnError(ccell2); }
                }
                else if (myCell.ColumnIndex == 2 && ccell3.Value != null)
                {// 3. Zelle normalisieren
                    if (!int.TryParse(ccell3.Value.ToString(), out int temp))
                    {
                        ccell3.Value = "0";
                        TidyUpOnError(ccell3); // TODO
                    }
                }
                UpdateLastDateColumns();
            }
        }

        private void UpdateLastTimeColumns()
        {// Spalte 'Saldo' und 'Spanne' werden komplett aktualisiert!
            TimeSpan tsCell7 = TimeSpan.Zero, tsCell6 = TimeSpan.Zero; string foo = string.Empty; TimeSpan pause = TimeSpan.Zero;
            for (int row = 0; row < dGV.RowCount - 1; ++row)
            {
                DataGridViewCell dgvCell1 = dGV.Rows[row].Cells[0]; // von
                DataGridViewCell dgvCell2 = dGV.Rows[row].Cells[1]; // bis
                DataGridViewCell dgvCell3 = dGV.Rows[row].Cells[2]; // Pause
                DataGridViewCell dgvCell4 = dGV.Rows[row].Cells[3]; // Faktor
                DataGridViewCell dgvCell6 = dGV.Rows[row].Cells[5]; // Spanne

                if (dgvCell1.Value != null && dgvCell2.Value != null && dgvCell4.Value != null && rgxValidTime.Match(dgvCell1.Value.ToString()).Success
                    && rgxValidTime.Match(dgvCell2.Value.ToString()).Success && Int32.TryParse(dgvCell4.Value.ToString(), out int rowFactor))
                {
                    DateTime d1 = DateTime.MinValue; DateTime d2 = DateTime.MinValue; DateTime d3 = DateTime.MinValue;
                    try
                    {// prüfen ob beide Zellen verwertbar sind
                        d1 = DateTime.ParseExact(dgvCell1.Value.ToString(), "H:mm", deCulture);
                        d2 = DateTime.ParseExact(dgvCell2.Value.ToString(), "H:mm", deCulture);
                        //d3 = DateTime.ParseExact(dgvCell3.Value.ToString(), "H:mm", deCulture);
                        if (d1 != DateTime.MinValue && d2 != DateTime.MinValue)
                        {
                            pause = DateTime.ParseExact(dgvCell3.Value.ToString(), "H:mm", deCulture).TimeOfDay;
                            if (DateTime.Compare(d2, d1) < 0) { d2 = d2.AddDays(1); } // Arbeitszeit über Nacht
                            TimeSpan ts = (d2 - d1).Subtract(pause);
                            dgvCell6.Value = String.Format("{0:00}:{1:00}", ts.Days * 24 + ts.Hours, ts.Minutes); // 24:00 ist keine gültige Zeitangabe! => hours = ts.Days * 24 + ts.Hours
                        }
                        else { dgvCell6.Value = null; }
                        int myMinute = 0; // 24:00 ist keine gültige Zeit! Deshalb wird ParseExact umgangen!
                        foo = dgvCell6.Value != null ? dgvCell6.Value.ToString() : string.Empty;
                        if (foo.Equals("24:00")) { foo = "23:59"; myMinute = 1; }
                        try { tsCell6 = foo != string.Empty ? DateTime.ParseExact(foo, "H:mm", deCulture).TimeOfDay : TimeSpan.Zero; }
                        catch (FormatException) { tsCell6 = TimeSpan.Zero; }
                        //if (tsCell6 != TimeSpan.Zero) // && rowFactor > 0
                        //{
                        tsCell6 = tsCell6.Add(TimeSpan.FromMinutes(myMinute));
                        tsCell7 += TimeSpan.FromTicks(tsCell6.Ticks * rowFactor);
                        dGV.Rows[row].Cells[6].Value = Math.Abs((int)tsCell7.TotalHours).ToString("00") + ":" + Math.Abs(tsCell7.Minutes).ToString("00");
                        //}
                        //else { dGV.Rows[row].Cells[6].Value = null; }
                    }
                    catch { dGV.Rows[row].Cells[5].Value = null; dGV.Rows[row].Cells[6].Value = null; }
                }
                else { dGV.Rows[row].Cells[5].Value = null; dGV.Rows[row].Cells[6].Value = null; }
            }
            if (tsCell7.TotalHours != 0)
            {
                linkLblDecText.Text = decText + tsCell7.TotalHours.ToString("#0.00");
                linkLblDecText.ToolTipText = "Ergebnis in die Zwischenablage kopieren";
                linkLblDecText.IsLink = true;
            }
            else
            {
                linkLblDecText.Text = String.Empty;
                linkLblDecText.ToolTipText = String.Empty;
                linkLblDecText.IsLink = false;
            }
            //if (!oldText.Equals(String.Empty) && !oldText.Equals(linkLblDecText.Text)) { nothingToSave = false; }
            //oldText = linkLblDecText.Text;
        }

        private void UpdateLastDateColumns()
        {// Spalte 'Saldo' und 'Spanne' werden komplett aktualisiert!
            int intCell7 = 0, intCell6 = 0, rowFactor = 0;
            DateTime d1 = DateTime.MinValue; DateTime d2 = DateTime.MinValue; //DateTime d1 = new DateTime(); DateTime d2 = new DateTime();
            ddfSum.years = 0; ddfSum.months = 0; ddfSum.days = 0;
            for (int row = 0; row < dGV.RowCount - 1; ++row)
            {
                DataGridViewCell dgvCell1 = dGV.Rows[row].Cells[0];
                DataGridViewCell dgvCell2 = dGV.Rows[row].Cells[1];
                DataGridViewCell dgvCell3 = dGV.Rows[row].Cells[2]; // Pause
                DataGridViewCell dgvCell4 = dGV.Rows[row].Cells[3];
                DataGridViewCell dgvCell6 = dGV.Rows[row].Cells[5];
                if (dgvCell1.Value != null && dgvCell2.Value != null && dgvCell4.Value != null && rgxValidDate.Match(dgvCell1.Value.ToString()).Success
                    && rgxValidDate.Match(dgvCell2.Value.ToString()).Success && int.TryParse(dgvCell4.Value.ToString(), out rowFactor))
                {
                    try
                    {
                        d1 = DateTime.ParseExact(dgvCell1.Value.ToString(), "d.M.yyyy", deCulture);
                        d2 = DateTime.ParseExact(dgvCell2.Value.ToString(), "d.M.yyyy", deCulture);
                        clsUtilities.DateDiff ddf = clsUtilities.calcDateDiff(d1, d2);
                        if (d1 != DateTime.MinValue && d2 != DateTime.MinValue)
                        {
                            TimeSpan ts = d2 - d1; // if (Convert.ToInt32(dGV.Rows[i].Cells[4].Value) != ts.Days) // ((int)dGV.Rows[i].Cells[4].Value funkt nicht
                            if (dgvCell3 != null && dgvCell3.Value != null && int.TryParse(dgvCell3.Value.ToString(), out int pause))
                            {// MessageBox.Show(pause.ToString());
                                ts = ts.Subtract(TimeSpan.FromDays(pause));
                            }
                            dgvCell6.Value = ts.Days; // MessageBox.Show(i.ToString() + ".: " + ddfSum.years.ToString() + " | " + ddfSum.months.ToString() + " | " + ddfSum.days.ToString());
                            if (rowFactor != 0) // if (ts != TimeSpan.Zero && rowFactor != 0)
                            {
                                ddfSum.years += ddf.years;
                                ddfSum.months += ddf.months;
                                ddfSum.days += ddf.days;
                                if (ddfSum.months > 12) { ddfSum.years++; ddfSum.months -= 12; }
                                else if (ddfSum.months < -12) { ddfSum.years--; ddfSum.months += 12; }
                            }
                            try { intCell6 = Convert.ToInt32(dgvCell6.Value); }
                            catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                            //if (intCell6 != 0)
                            //{
                            intCell7 += intCell6 * rowFactor;
                            dGV.Rows[row].Cells[6].Value = intCell7;
                            //}
                        }
                        else { dgvCell6.Value = null; dGV.Rows[row].Cells[6].Value = null; }
                    }
                    catch //  (Exception ex)
                    {//  MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        dgvCell6.Value = null; dGV.Rows[row].Cells[6].Value = null;
                    }
                }
                else { dgvCell6.Value = null; dGV.Rows[row].Cells[6].Value = null; }
            }
            if (intCell7 != 0) // irgendwo gibt es bereits einen Summenwert für Tage
            {
                linkLblDecText.Text = "" +
    (!ddfSum.years.Equals(0) ? ddfSum.years.ToString() + ((Math.Abs(ddfSum.years).Equals(1) ? " Jahr" : " Jahre") +
    (ddfSum.months.Equals(0) && ddfSum.days.Equals(0) ? "" : ", ")) : "") +
    (!ddfSum.months.Equals(0) ? ddfSum.months.ToString() + ((Math.Abs(ddfSum.months).Equals(1) ? " Monat" : " Monate") +
    (ddfSum.days.Equals(0) ? "" : ", ")) : "") +
    (!ddfSum.days.Equals(0) ? ddfSum.days.ToString() + (Math.Abs(ddfSum.days).Equals(1) ? " Tag" : " Tage") : "");
                linkLblDecText.ToolTipText = "Ergebnis in die Zwischenablage kopieren";
                linkLblDecText.IsLink = true;
            }
            else
            { // nirgendwo gibt es einen Summenwert in 4. Spalte
                linkLblDecText.Text = String.Empty;
                linkLblDecText.ToolTipText = String.Empty;
                linkLblDecText.IsLink = false;
            }
        }

        private void TidyUpOnError(DataGridViewCell myCell)
        {
            if (myCell != null & myCell.ColumnIndex < 3) { myCell.Style.ForeColor = Color.Red; }
            dGV.Rows[myCell.RowIndex].Cells[5].Value = null; // Spanne
            foreach (DataGridViewRow myRow in dGV.Rows)
            {
                myRow.Cells[6].Value = null; // Saldo
                //if (myRow.Index % 2 != 0)
                //{// AlternatingCellStyle 
                //    dGV.Rows[myRow.Index].Cells[0].Style.BackColor = dGV.Rows[myRow.Index].Cells[1].Style.BackColor = dGV.Rows[myRow.Index].Cells[2].Style.BackColor = Color.WhiteSmoke;
                //}
                //else
                //{// nach dem Löschen von 1,3 oder einer anderen ungeraden Zahl von Rows
                //    dGV.Rows[myRow.Index].Cells[0].Style.BackColor = dGV.Rows[myRow.Index].Cells[1].Style.BackColor = dGV.Rows[myRow.Index].Cells[2].Style.BackColor = Color.White;
                //}
            }
        }

        private void ChangeModus()
        {// ehemals: rbStundenTageCheckedChange
            if (!clsUtilities.isDGVEmpty(dGV) && !isFileLoading)  // Lösung für Programmstart mit Dateipfadangabe und 'Tage'-Tabelle
            {// dGV.Refresh();
                MessageBox.Show("Der Wechsel vom Stunden- zum Tagerechner\nund v. v. ist nur bei leerer Tabelle möglich!\n\nLöschen Sie alle Daten oder wählen Sie »Neue\nDatei« und versuchen Sie es dann noch mal.\n\nBy the way: Beim Laden einer Datei erkennt das\nProgramm automatisch den korrekten Modus.", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                modusStunden = !modusStunden;
                if (dGV.IsCurrentCellInEditMode) { dGV.EndEdit(); }
                DataGridViewComboBoxColumn comboList = (DataGridViewComboBoxColumn)dGV.Columns[3]; // ["Faktor"];
                comboList.Items.Clear();
                if (modusStunden)
                {
                    comboList.Items.AddRange("0", "1", "2", "3", "4", "5", "6", "7", "8", "9");
                    //toolStripButtonModus.Image = global::TimeCalc.Properties.Resources.ModHours;
                    labelHeading.Text = "Stundenrechner";
                }
                else
                {// DataGridViewComboBoxColumn comboList = (DataGridViewComboBoxColumn)dGV.Columns
                    comboList.Items.AddRange("0", "1");
                    //toolStripButtonModus.Image = global::TimeCalc.Properties.Resources.ModDays;
                    labelHeading.Text = "Tagerechner";
                }
                toolStripButtonModus.Image.RotateFlip(RotateFlipType.RotateNoneFlipXY);
                tsSymbolleiste.PerformLayout(); // sonst funktioniert vorstehende Zeile nicht bei F12
                tageToolStripMenuItem.Enabled = stundenToolStripMenuItem.Checked = modusStunden;
                stundenToolStripMenuItem.Enabled = tageToolStripMenuItem.Checked = !modusStunden;
                linkLblDecText.Text = String.Empty;
                linkLblDecText.ToolTipText = String.Empty;
                linkLblDecText.IsLink = false;
                dGV.ClearSelection();
                dGV.CurrentCell = null;
                dGV.Rows[0].Cells[0].Selected = true;
                //if (!clsUtilities.isDGVEmpty(dGV) && !isFileLoading)  // Lösung für Programmstart mit Dateipfadangabe und 'Tage'-Tabelle
                //{// dGV.Refresh();
                //    if (MessageBox.Show("Möchten Sie alle Daten löschen?", winTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                //    {
                //        dGV.Rows.Clear(); //dGV.Refresh();
                //    }
                //}
                dGV.CurrentCell = dGV.Rows[0].Cells[0];
                if (dGV.ContainsFocus) { dGV.BeginEdit(true); }
                else dGV.Focus();
            }
        }

        private void DGV_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //foreach (DataGridViewRow myRow in dGV.Rows)
            //{// AlternatingCellStyle 
            //    if (myRow.Index % 2 == 0) { dGV.Rows[myRow.Index].Cells[0].Style.BackColor = dGV.Rows[myRow.Index].Cells[1].Style.BackColor = dGV.Rows[myRow.Index].Cells[2].Style.BackColor = Color.White; }
            //    else { dGV.Rows[myRow.Index].Cells[0].Style.BackColor = dGV.Rows[myRow.Index].Cells[1].Style.BackColor = dGV.Rows[myRow.Index].Cells[2].Style.BackColor = Color.WhiteSmoke; }
            //}
            if (modusStunden) { UpdateLastTimeColumns(); }
            else { UpdateLastDateColumns(); }
        }

        private void HelpMessageBoxShow()
        {
            StringBuilder messageBoxCS = new StringBuilder();
            messageBoxCS.Append("Bei der Eingabe von Daten gelten folgende Zeichen als Separatoren:" +
                Environment.NewLine + "Punkt, Doppelpunkt, Komma, Semikolon, Schrägstrich sowie Stern-," +
                Environment.NewLine + "und Minuszeichen. Unvollständige Eingaben werden komplettiert!");
            messageBoxCS.AppendLine();
            messageBoxCS.AppendLine();
            messageBoxCS.Append("Durch Drücken der Enter- oder Tab-Taste gelangen Sie zur nächsten" +
                Environment.NewLine + "Zelle. Dabei erfolgt automatisch eine Interpretation der Eingabe. Die" +
                Environment.NewLine + "Umwandlung in ein gültiges Zeit- beziehungsweise Datumsformat" +
                Environment.NewLine + "kann auch durch Eingabe eines Doppelkreuzes (#) ausgelöst werden.");
            messageBoxCS.AppendLine();
            messageBoxCS.AppendLine();
            messageBoxCS.Append("Die aktuelle Uhrzeit bzw. das aktuelle Datum lässt sich durch Drü-" +
                Environment.NewLine + "cken der Tasten \"F5\" oder \"Strg+J\" (\"Jetzt\") eintragen." +  //\"J\" (\"Jetzt\"), 
                Environment.NewLine + "Ein bereits eingegebenes Datum kann durch Drücken der Plus- oder" +
                Environment.NewLine + "Minus-Taste verändert werden. Wenn zusätzlich die Shift-Taste ge-" +
                Environment.NewLine + "drückt wird, ändert sich der größere Wert (z.B. Stunde statt Minute).");
            messageBoxCS.AppendLine();
            messageBoxCS.AppendLine();
            messageBoxCS.Append("Die errechnete Gesamtstundenzahl kann durch Anklicken des Links" +
                Environment.NewLine + "in der Statuszeile kurzerhand in die Zwischenablage kopiert werden.");
            MessageBox.Show(messageBoxCS.ToString(), "Hilfe - " + winTitle);
        }

        private DateTime RetrieveLinkerTimestamp() // lässt sich nicht in Utilities.cs verlagern
        {
            string filePath = Assembly.GetCallingAssembly().Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            byte[] b = new byte[2048];
            Stream s = null;
            try
            {
                s = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally { if (s != null) s.Close(); }
            int i = BitConverter.ToInt32(b, c_PeHeaderOffset);
            int secondsSince1970 = BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours);
            return dt;
        }

        private void DGV_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {// ContextMenu bei RightClick anzeigen
            int eRow = e.RowIndex; int eCol = e.ColumnIndex;
            if (eRow != -1 && e.Button == MouseButtons.Right) // Ignore if a column is clicked; 
            {// MessageBox.Show(eRow.ToString() + " | " + eCol.ToString());
                eCol = eCol == -1 ? 0 : eCol; // allows click to RowHeader
                DataGridViewCell clickedCell = dGV.Rows[eRow].Cells[eCol];
                ShowContextMenu(clickedCell, dGV, dGV.PointToClient(Cursor.Position));
            }
            else if (eCol == -1 && eRow == -1)
            {// DataGridViewTopLeftHeaderCell
                if (dGV.IsCurrentCellInEditMode)
                {
                    dGV.EndEdit();
                    dGV.CurrentCell.Selected = true;
                }
            }
        }

        private void ShowContextMenu(DataGridViewCell cCell, Control ctrl, Point pos)
        {
            if (cCell != null && !cCell.IsInEditMode)
            {
                dGV.EndEdit();
                dGV.ClearSelection();
                cCell.Selected = true;
                dGV.CurrentCell = cCell;
                foreach (ToolStripItem item in contextMenuStrip.Items) { item.Enabled = true; } // Reset contextMenuStrip
                if (dGV.Rows.Count == 1 && dGV.CurrentRow.Cells[0].Value == null && dGV.CurrentRow.Cells[1].Value == null)
                {// reduziertes Menü zeigen wenn Click auf NeuerZeile als einziger Zeile ohne Inhalt erfolgt
                    allesLöschenContextMenuItem.Enabled = false;
                    zeileLöschenContextMenuItem.Enabled = false;
                }
                if (dGV.Rows[dGV.CurrentRow.Index].IsNewRow)
                { zeileLöschenContextMenuItem.Enabled = false; }
                contextMenuStrip.Show(ctrl, pos);
            }
        }

        private void AllesSpeichernContextMenuItem_Click(object sender, EventArgs e)
        {
            if (currFile.Length > 0) { SaveTextFile(currFile); }
            else
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (SaveTextFile(saveFileDialog.FileName)) { AskForDeskShortcut(); }
                }
            }
        }

        private void AllesLöschenContextMenuItem_Click(object sender, EventArgs e)
        {
            nothingToSave = false;
            dGV.Rows.Clear();
            dGV.Refresh();
            dGV.CurrentCell = dGV.Rows[0].Cells[0];
            dGV.BeginEdit(true);
        }

        private void ZeileEinfügenContextMenuItem_Click(object sender, EventArgs e)
        {
            dGV.Rows.Insert(dGV.CurrentCell.RowIndex, 1);
        }

        private void ZeileLöschenContextMenuItem_Click(object sender, EventArgs e)
        {// if (!dGV.Rows[dGV.CurrentRow.Index].IsNewRow) //if (dGV.CurrentCell.RowIndex != dGV.Rows.Count - 1)
            DeleteRowToolStripMenuItem_Click(null, null);
        }

        private void DGV_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            dGV.CurrentCell = dGV[0, dGV.CurrentRow.Index];
            dGV.BeginEdit(true);
        }

        private void LinkLblDecText_Click(object sender, EventArgs e)
        {
            if (linkLblDecText.IsLink == true)
            {
                string cbText = linkLblDecText.Text.Replace(decText, "");
                try
                {
                    if (cbText.Length > 0)
                    {
                        Clipboard.SetText(cbText, TextDataFormat.Text);
                        MessageBox.Show(cbText, "Inhalt der Zwischenablage", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else { MessageBox.Show("Es liegt keine Berechnung vor!", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
            }
        }

        private void DGV_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                dGV.CurrentCell.Style.BackColor = Color.MistyRose;
                toolStripButtonIncrease.Enabled = true;
                toolStripButtonIncreaseMuch.Enabled = true;
                toolStripButtonDecrease.Enabled = true;
                toolStripButtonDecreaseMuch.Enabled = true;
                toolStripButtonNow.Enabled = true;
            }
        }

        private void DGV_CellLeave(object sender, DataGridViewCellEventArgs e)
        {// wenn ein Wert in den letzten beiden Spalten gelöscht wurde, soll aktualisiert werden
            if (dGV.CurrentCell.ColumnIndex > dGV.ColumnCount - 2) // in letzten beiden Spalten 
            {// MessageBox.Show(dGV.CurrentCell.Value.ToString() + " | " + e.RowIndex.ToString() + " | " + e.ColumnIndex.ToString());
                CalculateSaldo(dGV.Rows[e.RowIndex].Cells[e.ColumnIndex]);
            }
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                dGV.CurrentCell.Style.BackColor = dGV.DefaultCellStyle.BackColor;
            }
            if (dGV.CurrentCell.ColumnIndex == 2 && modusStunden && (dGV.Rows[e.RowIndex].Cells[0].Value != null || dGV.Rows[e.RowIndex].Cells[1].Value != null)) // Pause
            {
                dGV.CurrentCell.Value = dGV.CurrentCell.Value == null || !rgxValidTime.Match(dGV.CurrentCell.Value.ToString()).Success ? "0:00" : dGV.CurrentCell.Value;
            }
            else if (dGV.CurrentCell.ColumnIndex == 2 && !modusStunden && (dGV.Rows[e.RowIndex].Cells[0].Value != null || dGV.Rows[e.RowIndex].Cells[1].Value != null)) // Pause
            {
                dGV.CurrentCell.Value = dGV.CurrentCell.Value == null || !int.TryParse(dGV.CurrentCell.Value.ToString(), out int temp) ? "0" : dGV.CurrentCell.Value;
            }
            if (dGV.CurrentCell.ColumnIndex == 3 && (dGV.Rows[e.RowIndex].Cells[0].Value != null || dGV.Rows[e.RowIndex].Cells[1].Value != null) && dGV.CurrentCell.Value == null) // Faktor
            {
                dGV.CurrentCell.Value = "1";
            }
            //if (dGV.CurrentCell.ColumnIndex == 3 && (dGV.Rows[e.RowIndex].Cells[0].Value != null || dGV.Rows[e.RowIndex].Cells[1].Value != null)) // Faktor
            //{
            //    dGV.CurrentCell.Value = dGV.CurrentCell.Value == null ? "1" : dGV.CurrentCell.Value;
            //}
            toolStripButtonIncrease.Enabled = false;
            toolStripButtonIncreaseMuch.Enabled = false;
            toolStripButtonDecrease.Enabled = false;
            toolStripButtonDecreaseMuch.Enabled = false;
            toolStripButtonNow.Enabled = false;
        }

        private void DGV_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {// MessageBox.Show(dGV.SelectionMode.ToString());
            if (dGV.IsCurrentCellInEditMode)
            {
                dGV.EndEdit();
                dGV.CurrentCell.Selected = true;
            }
            if (dGV.SelectedRows.Count > 0) { dGV.ClearSelection(); }
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                if (dGV.CurrentCell.ColumnIndex == e.ColumnIndex)
                {// erst disselect
                    dGV.CurrentCell.Selected = false;
                }// dann erneut select
                for (int r = 0; r < dGV.RowCount; r++)
                {
                    if (dGV[e.ColumnIndex, r].Selected) { dGV[e.ColumnIndex, r].Selected = false; }
                    else { dGV[e.ColumnIndex, r].Selected = true; }
                }
            }
            else
            {
                if (dGV.CurrentCell.ColumnIndex != e.ColumnIndex)
                {// CurrentCell darf nie null sein!
                    dGV.CurrentCell = dGV[e.ColumnIndex, 0];
                }
                for (int r = 0; r < dGV.RowCount; r++) { dGV[e.ColumnIndex, r].Selected = true; }
            }
        }

        private void TageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeModus();
        }

        private void StundenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeModus();
        }

        private void NeuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveGridViewIfNecessary();
            {
                AllesLöschenContextMenuItem_Click(null, null);
                currFile = string.Empty;
                Text = "Unbenannt - " + winTitle;
            }
        }

        private void SaveGridViewIfNecessary()
        {
            if (!nothingToSave) // && !clsUtilities.isDGVEmpty(dGV))
            {
                if (currFile.Length > 0)
                {
                    if (MessageBox.Show("Möchten Sie die Änderungen an " + currFile + " speichern?", winTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                    { nothingToSave = SaveTextFile(currFile); }
                }
                else
                {
                    if (MessageBox.Show("Möchten Sie die Eingaben speichern?", winTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                    {
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            if (SaveTextFile(saveFileDialog.FileName)) { AskForDeskShortcut(); }
                        }
                    }
                }
            }
        }

        private void SpeichernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (currFile.Length > 0) { SaveTextFile(currFile); }
            else
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (SaveTextFile(saveFileDialog.FileName)) { AskForDeskShortcut(); }
                }
            }
        }

        private void SpeichernUnterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (currFile.Length > 0)
            {
                saveFileDialog.FileName = Path.GetFileName(currFile); // set a default file name
                saveFileDialog.InitialDirectory = Path.GetDirectoryName(currFile);
            }
            else { saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); }
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string tempFile = currFile; // saveTextFile aktualisiert currFile
                if (SaveTextFile(saveFileDialog.FileName))
                {
                    if (saveFileDialog.FileName != tempFile) { AskForDeskShortcut(); }
                }
            }
        }

        private void AskForDeskShortcut()
        {
            if (MessageBox.Show("Möchten Sie eine Desktopverknüpfung anlegen?", winTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                AddShortcutToDesktop();
            }
        }

        private void BeendenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void UhrzeitDatumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                int iCol = dGV.CurrentCell.ColumnIndex;
                int iRow = dGV.CurrentCell.RowIndex;
                if (dtEdits.Contains(iCol) && iRow == dGV.Rows.Count - 1) // if (iCol < 3 && iRow == dGV.Rows.Count - 1)
                {
                    dGV.Rows.Add();
                    dGV.CurrentCell = dGV.Rows[dGV.CurrentCell.RowIndex - 1].Cells[iCol];
                    dGV.CurrentCell.Selected = true;
                }
                dGV.CurrentCell.Value = null; // ""
                if (modusStunden && dGV.CurrentCell.ColumnIndex < 2)
                {
                    dGV.CurrentCell.Value = DateTime.Now.ToString("H:mm");
                }
                else
                {
                    dGV.CurrentCell.Value = DateTime.Now.ToString("d.M.yyyy");
                }
                dGV.RefreshEdit();
                if (!dGV.IsCurrentCellInEditMode) { dGV.BeginEdit(true); }
            }
        }

        private void InsertRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.IsCurrentCellInEditMode) { dGV.EndEdit(); }
            dGV.Rows.Insert(dGV.CurrentCell.RowIndex, 1);
        }

        private void DeleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {// NICHT wenn Cursor in letzter (neu angelegter) Zeile steht!
            if (dGV.CurrentCell.RowIndex != dGV.Rows.Count - 1)
            {
                if (dGV.IsCurrentCellInEditMode) { dGV.EndEdit(); }
                //int oldRowIndex = dGV.CurrentCellAddress.X;
                if (dGV.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow row in dGV.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            for (int i = 0; i < 3; i++)
                            {// nur wg. nothingToSave
                                if (row.Cells[i].Value != null)
                                {
                                    nothingToSave = false;
                                    break;
                                }
                            }
                            dGV.Rows.Remove(row);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < 3; i++)
                    {// nur wg. nothingToSave
                        if (dGV.CurrentCell.OwningRow.Cells[i].Value != null)
                        {
                            nothingToSave = false;
                            break;
                        }
                    }
                    dGV.Rows.RemoveAt(dGV.CurrentCell.RowIndex);
                }
                TidyUpOnError(dGV.CurrentCell);
                dGV.BeginEdit(true); // Workaround damit neue Zelle validiert wird
                dGV.EndEdit();
                //if (oldRowIndex < dGV.RowCount) { dGV.CurrentCell = dGV.Rows[oldRowIndex + 1].Cells[0]; } // CurrentRow eine Zeile tiefer
                dGV.CurrentCell.Selected = true;
            }
            else { Console.Beep(); }
        }

        private void AusschneidenToolStripMenuItem_Click(object sender, EventArgs e)
        {// dGV.Focus(); darf hier nicht stehen weil Zelle in EditMode sonst den Fokus verliert!
            if (!dGV.CurrentRow.Selected && !dGV.IsCurrentCellInEditMode)
            {
                if (dGV.CurrentCell.Value != null)
                {
                    Clipboard.SetText(dGV.CurrentCell.Value.ToString(), TextDataFormat.Text);
                    dGV.CurrentCell.Value = null;
                    TidyUpOnError(dGV.CurrentCell);
                    dGV.BeginEdit(true);
                    nothingToSave = false;
                }
                else { Console.Beep(); }
            }
            else if (dGV.IsCurrentCellInEditMode)
            {
                if (((TextBox)dGV.EditingControl).SelectionLength > 0)
                {
                    ((TextBox)dGV.EditingControl).Cut();
                }
            }
        }

        private void KopierenToolStripMenuItem_Click(object sender, EventArgs e)
        {// dGV.Focus(); darf hier nicht stehen weil Zelle in EditMode sonst den Fokus verliert!
            if (!dGV.CurrentRow.Selected && !dGV.IsCurrentCellInEditMode)
            {
                if (dGV.CurrentCell.Value != null)
                {
                    try
                    {
                        Clipboard.Clear();
                        Clipboard.SetText(dGV.CurrentCell.Value.ToString(), TextDataFormat.Text);
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                }
                else { Console.Beep(); }
            }
            else if (dGV.IsCurrentCellInEditMode)
            {// Ensure that text is selected in the text box.
                if (((TextBox)dGV.EditingControl).SelectionLength > 0)
                {// SendKeys.Send("^{c}");
                    ((TextBox)dGV.EditingControl).Copy();
                }
            }
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {// dGV.Focus(); darf hier nicht stehen weil Zelle in EditMode sonst den Fokus verliert!
            if (!dGV.IsCurrentCellInEditMode)
            {
                int iCol = dGV.CurrentCell.ColumnIndex;
                int iRow = dGV.CurrentCell.RowIndex;
                if (iCol < 2 && iRow == dGV.Rows.Count - 1)
                {
                    dGV.Rows.Add();
                    dGV.CurrentCell = dGV.Rows[dGV.CurrentCell.RowIndex - 1].Cells[iCol];
                    dGV.CurrentCell.Selected = true;
                }
                dGV.CurrentCell.Value = null;
                try
                {
                    dGV.CurrentCell.Value = Clipboard.GetText();
                    dGV.BeginEdit(true);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                dGV.RefreshEdit();
            }
            else if (dGV.IsCurrentCellInEditMode)
            {// Determine if there is any text in the Clipboard to paste into the text box.
                if (Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) == true)
                {
                    ((TextBox)dGV.EditingControl).Paste();
                }
            }
        }

        private void LöschenToolStripMenuItem_Click(object sender, EventArgs e)
        {// dGV.Focus(); darf hier nicht stehen weil Zelle in EditMode sonst den Fokus verliert!
            if (dGV.IsCurrentCellInEditMode)
            {// Ensure that text is selected in the text box.   
                if (((TextBox)dGV.EditingControl).SelectionLength > 0)
                {// nothingToSave = false;
                    ((TextBox)dGV.EditingControl).SelectedText = "";
                }
            }
            else
            {
                foreach (DataGridViewCell cell in dGV.SelectedCells)
                {
                    //if (cell.Value != null) { nothingToSave = false; }
                    cell.Value = null;
                    //if (dGV.CurrentCell.ColumnIndex != 2 && dGV.CurrentCell.ColumnIndex != 4) { TidyUpOnError(cell); } // nicht in Notizspalte
                }
                //dGV.Refresh();
                dGV.ClearSelection();
                dGV.CurrentCell.Selected = true;
                dGV.BeginEdit(true);
            }
        }

        private void HilfeAnzeigenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HelpMessageBoxShow();
        }

        private void DGV_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dGV.IsCurrentCellInEditMode)
            {
                dGV.EndEdit();
                dGV.CurrentCell.Selected = true;
            }
        }

        private void AllesAuswählenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dGV.SelectAll();
        }

        private void WertErhöhenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                ChangeCellValueByKey(false, true); // Shift, Add
            }
        }

        private void WertVerringernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                ChangeCellValueByKey(false, false); // Shift, Add
            }
        }

        private void GroßenWertErhöhenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                ChangeCellValueByKey(true, true); // Shift true, Add true
            }
        }

        private void GroßenWertVerringernToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten
            {
                ChangeCellValueByKey(true, false); // Shift true, Add false
            }
        }

        private void AllesLöschenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AllesLöschenContextMenuItem_Click(null, null);
        }

        private void ToolStripMenuItemShortcut_Click(object sender, EventArgs e) { AddShortcutToDesktop(); }

        private void MainMenuStrip_MenuDeactivate(object sender, EventArgs e)
        {
            ausschneidenToolStripMenuItem.Enabled = false;
            kopierenToolStripMenuItem.Enabled = false;
            löschenToolStripMenuItem.Enabled = false;
            pasteToolStripMenuItem.Enabled = false;
            wertErhöhenToolStripMenuItem.Enabled = false;
            wertVerringernToolStripMenuItem.Enabled = false;
            großenWertErhöhenToolStripMenuItem.Enabled = false;
            großenWertVerringernToolStripMenuItem.Enabled = false;
            deleteRowToolStripMenuItem.Enabled = false;
            uhrzeitDatumToolStripMenuItem.Enabled = false;
        }

        private void MainMenuStrip_MenuActivate(object sender, EventArgs e)
        {
            if (dGV.IsCurrentCellInEditMode)
            {
                enterToolStripMenuItem.Text = "&Verlassen";
                enterToolStripMenuItem.ShortcutKeyDisplayString = "Eingabtaste";
                enterToolStripMenuItem.Image = Properties.Resources.GoToNext;
            }
            else
            {
                enterToolStripMenuItem.Text = "E&ditieren";
                enterToolStripMenuItem.ShortcutKeyDisplayString = "F2 oder Eingabtaste";
                enterToolStripMenuItem.Image = Properties.Resources.EditTable;
            }
            if (dGV.CurrentCell.ColumnIndex != 2) // dGV.CurrentCell.EditType == typeof(DataGridViewComboBoxEditingControl funkt nicht in EditMode
            {
                uhrzeitDatumToolStripMenuItem.Enabled = true;
                DateTime dt = new DateTime();
                if (dGV.CurrentCell == null) { dGV.CurrentCell = dGV.Rows[0].Cells[0]; } // Fehler wenn CurrentCell nicht existiert: 
                string dtFormat = modusStunden && dGV.CurrentCell.ColumnIndex < 2 ? "H:mm" : "d.M.yyyy"; //string[] dtFormats = { "H:mm", "d.M.yyyy" };

                if (modusStunden && dGV.CurrentCell.ColumnIndex < 2)
                {
                    wertErhöhenToolStripMenuItem.Text = "&Minute steigern";
                    wertVerringernToolStripMenuItem.Text = "M&inute mindern";
                    großenWertErhöhenToolStripMenuItem.Text = "&Stunde steigern";
                    großenWertVerringernToolStripMenuItem.Text = "S&tunde mindern";
                    uhrzeitDatumToolStripMenuItem.Text = "A&kt. Uhrzeit";
                }
                else
                {
                    wertErhöhenToolStripMenuItem.Text = "&Tag steigern";
                    wertVerringernToolStripMenuItem.Text = "Ta&g mindern";
                    großenWertErhöhenToolStripMenuItem.Text = "&Monat steigern";
                    großenWertVerringernToolStripMenuItem.Text = "M&onat mindern";
                    uhrzeitDatumToolStripMenuItem.Text = "A&kt. Datum";
                }

                if (dGV.IsCurrentCellInEditMode && ((TextBox)dGV.EditingControl).SelectionLength > 0)
                {
                    ausschneidenToolStripMenuItem.Enabled = true;
                    kopierenToolStripMenuItem.Enabled = true;
                    löschenToolStripMenuItem.Enabled = true;
                }
                else if (dGV.CurrentCell.Value != null)
                {
                    ausschneidenToolStripMenuItem.Enabled = true;
                    kopierenToolStripMenuItem.Enabled = true;
                    löschenToolStripMenuItem.Enabled = true;
                }

                if (dGV.IsCurrentCellInEditMode)
                {
                    string tbText = ((TextBox)dGV.EditingControl).Text;
                    if (modusStunden && dGV.CurrentCell.ColumnIndex < 2) { dt = clsUtilities.NormalizeTime(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                    else { dt = clsUtilities.NormalizeDate(tbText.Length == 0 ? DateTime.Now.ToString(dtFormat) : tbText); }
                    if (dt != DateTime.MinValue)
                    {
                        wertErhöhenToolStripMenuItem.Enabled = true;
                        wertVerringernToolStripMenuItem.Enabled = true;
                        großenWertErhöhenToolStripMenuItem.Enabled = true;
                        großenWertVerringernToolStripMenuItem.Enabled = true;
                    }
                }
                else if (!dGV.IsCurrentCellInEditMode)
                {
                    if (dGV.CurrentCell.Value == null || DateTime.TryParseExact(dGV.CurrentCell.Value.ToString(), dtFormat, deCulture, DateTimeStyles.None, out dt))
                    {
                        wertErhöhenToolStripMenuItem.Enabled = true;
                        wertVerringernToolStripMenuItem.Enabled = true;
                        großenWertErhöhenToolStripMenuItem.Enabled = true;
                        großenWertVerringernToolStripMenuItem.Enabled = true;
                    }
                }

                if (Clipboard.ContainsText()) { pasteToolStripMenuItem.Enabled = true; }

                if (!dGV.Rows[dGV.CurrentRow.Index].IsNewRow) { deleteRowToolStripMenuItem.Enabled = true; }
            }
        }

        private void EnterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int iCol = dGV.CurrentCell.ColumnIndex;
            int iRow = dGV.CurrentCell.RowIndex;
            if (iCol < dGV.ColumnCount - 2) // nicht in letzten beiden Spalten
            {
                if (dGV.IsCurrentCellInEditMode)
                {
                    if (iCol == dGV.ColumnCount - 3) // in 3. = letzte EditSpalte
                    {
                        if (iRow + 1 == dGV.RowCount) { dGV.Rows.Add(); }
                        dGV.CurrentCell = dGV[0, iRow + 1];
                    }
                    else { dGV.CurrentCell = dGV[iCol + 1, iRow]; }
                }
                else { dGV.BeginEdit(true); }
            }
            else // in den letzten beiden Spalten
            {// CurrentCell soll in jedem Fall in neuer Zeile starten
                if (iRow + 1 == dGV.RowCount) { dGV.Rows.Add(); }
                dGV.CurrentCell = dGV[0, iRow + 1];
            }
        }

        private void ÖffnenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveGridViewIfNecessary();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dGV.Rows.Clear(); //dGV.AllowUserToAddRows = false; // wird später wieder auf true gesetzt
                File.SetLastAccessTime(openFileDialog.FileName, DateTime.Now); // The NtfsDisableLastAccessUpdate registry setting is enabled by default
                ReadTextFile(openFileDialog.FileName);
                if (dGV.Rows[0].Cells[0].Value != null && rgxValidDate.Match(dGV.Rows[0].Cells[0].Value.ToString()).Success) // Regex.Match(dGV.Rows[0].Cells[0].Value.ToString(), @"^\d{1,2}\.\d{1,2}\.\d{4}$", RegexOptions.Compiled).Success)
                { modusStunden = false; }
                else if (dGV.Rows[0].Cells[0].Value != null && rgxValidTime.Match(dGV.Rows[0].Cells[0].Value.ToString()).Success) // Regex.Match(dGV.Rows[0].Cells[0].Value.ToString(), @"^\d{1,2}\:\d{2}$", RegexOptions.Compiled).Success)
                { modusStunden = true; }
            }
        }

        private void FrmTimeCalc_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!nothingToSave && e.CloseReason == CloseReason.UserClosing) // && !clsUtilities.isDGVEmpty(dGV))
            {
                if (currFile.Length > 0)
                {
                    DialogResult dlgResult = MessageBox.Show("Möchten Sie die Änderungen an " + currFile + " speichern?", winTitle, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            if (!SaveTextFile(currFile)) { e.Cancel = true; }
                            break;
                        case DialogResult.No:
                            break;
                        default:
                            e.Cancel = true; // cancel the closure of the form.
                            break;
                    }
                }
                else
                {
                    if (MessageBox.Show("Möchten Sie die Eingaben speichern?", winTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                    {
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string tempFile = currFile; // saveTextFile aktualisiert currFile
                            if (SaveTextFile(saveFileDialog.FileName))
                            {
                                if (saveFileDialog.FileName != tempFile) { AskForDeskShortcut(); }
                            }
                            else { e.Cancel = true; } // Speichern führte zu Fehler
                        }
                        else { e.Cancel = true; } // User möchte nicht speichern
                    }
                }
            }
        }

        private bool SaveTextFile(string fullFilename)
        {
            try
            {
                if (dGV.IsCurrentCellInEditMode)
                {
                    dGV.EndEdit();
                    dGV.CurrentCell.Selected = true;
                }
                clsUtilities.removeEmptyRows(dGV, dtEdits);
                using (StreamWriter strmWriter = new StreamWriter(fullFilename))
                {
                    string tab;
                    string ext = Path.GetExtension(fullFilename);
                    string sep = ext.Equals(".csv", StringComparison.OrdinalIgnoreCase) ? ";" : "\t";
                    for (int j = 0; j < dGV.Columns.Count; j++)
                    {// Export titles
                        tab = (j > 0) ? sep : String.Empty;
                        strmWriter.Write(tab + Convert.ToString(dGV.Columns[j].HeaderText));
                    }
                    strmWriter.WriteLine();
                    for (int i = 0; i < dGV.RowCount - 1; i++)
                    {
                        for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)
                        {
                            tab = (j > 0) ? sep : String.Empty;
                            strmWriter.Write(tab + Convert.ToString(dGV.Rows[i].Cells[j].Value));
                        }
                        strmWriter.WriteLine();
                    }
                }
                nothingToSave = true;
                currFile = fullFilename;
                Text = currFile + " - " + winTitle;  // Text = Path.GetFileName(fullFilename) + " - " + winTitle;
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Fehlermeldung", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            return false;
        }

        private bool ReadTextFile(string fullFilename)
        {
            try
            {
                using (StreamReader sReader = new StreamReader(fullFilename))
                {// You're better off using the using keyword; then you don't need to explicitly close anything.
                    String sLine = "";
                    sLine = sReader.ReadLine(); // erste Zeile lesen (gelangt nicht ins DataGridView)
                    bool ohnePause = sLine.Contains("bis;x;"); // stellt Kompatiblität zu Version < 1.7 her
                    sLine = sReader.ReadLine(); // zweite Zeile
                    while (sLine != null)
                    {
                        if (ohnePause)
                        {
                            string[] tokens = sLine.Split(';');
                            sLine = tokens[0] + ";" + tokens[1] + ";0:00;" + tokens[2] + ";" + tokens[3] + ";" + tokens[4] + ";" + tokens[5];
                        }
                        dGV.Rows.Add(sLine.Split(';', '\t'));
                        sLine = sReader.ReadLine();
                    }
                } //dGV.AllowUserToAddRows = true; // fügt leere Bearbeitungszeile am Ende hinzu
                nothingToSave = true;
                currFile = fullFilename;
                Text = currFile + " - " + winTitle;  // Text = Path.GetFileName(fullFilename) + " - " + winTitle;
                return true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Fehlermeldung", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            return false;
        }

        private void DGV_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {// Kein Semikolon in TextBoxen erlauben
            e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress); // remove an existing event-handler, if present
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress); // add the event handler
            e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown); // viel Aufwand um NumpadDecimal-Komma mit Punkt zu ersetzen
            e.Control.KeyDown += new KeyEventHandler(Control_KeyDown); // s.o.
        }

        private void Control_KeyDown(object sender, KeyEventArgs e)
        {
            var tB = sender as DataGridViewTextBoxEditingControl;
            if (e.KeyValue == 110) // NumpadDecimal
            {
                numPadDecimal = true;
                int iPos = tB.SelectionStart; // Cursorpositon in TextBox
                if (tB.SelectionLength > 0)
                {// falls Text markiert ist, wird er gelöscht
                    tB.Text = tB.Text.Remove(iPos, tB.SelectionLength);
                }
                tB.Text = tB.Text.Insert(iPos, ".");
                tB.SelectionStart = iPos + 1; // reposition cursor
            }
        }

        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {// This event occurs after the KeyDown event and can be used to prevent characters from entering the control
            if (numPadDecimal == true)
            {// Stop the character from being entered into the control
                e.Handled = true;
                numPadDecimal = false;
            }
            if (dGV.CurrentCell.EditType == typeof(DataGridViewComboBoxEditingControl))
            {
                dGV.BeginEdit(true);
                ((ComboBox)dGV.EditingControl).DroppedDown = true;
            }
            else
            {
                var txtBox = (TextBox)sender;
                if (e.KeyChar == ';') { e.KeyChar = ','; } //{ Console.Beep(); e.Handled = true; } // if (char.IsNumber(e.KeyChar))
                if (e.KeyChar == '#')
                {
                    DateTime dt = new DateTime();
                    string dtFormat = modusStunden && dGV.CurrentCell.ColumnIndex < 2 ? "H:mm" : "d.M.yyyy";
                    string strTxtBox = txtBox.TextLength == 0 ? DateTime.Now.ToString(dtFormat) : txtBox.Text;
                    if (modusStunden && dGV.CurrentCell.ColumnIndex < 2) { dt = clsUtilities.NormalizeTime(strTxtBox); }
                    else { dt = clsUtilities.NormalizeDate(strTxtBox); }
                    if (dt != DateTime.MinValue)
                    {
                        e.Handled = true; // Doppelkreuz '#' weglassen
                        txtBox.Text = dt.ToString(dtFormat); // Textbox Inhalt ersetzten
                        txtBox.Select(txtBox.Text.Length, 0); // Cursor ans Ende setzten (Markierungslänge = 0)
                    }
                }
            }
        }

        private void WebseiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try { Process.Start("https://timetool.codeplex.com/"); }
            catch (Exception ex) { MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private void InfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string mbText =
@"Das Programm ist Freeware. Sie dürfen es kostenlos nutzen
und  weitergeben, aber nicht verändern.
Die Benutzung erfolgt auf eigene Gefahr! Der Autor ist nicht
für Schäden verantwortlich, die durch Verwendung oder Ver-
breitung der Software verursacht werden.
In keinem Fall ist der Autor verantwortlich für entgangenen
Umsatz, Gewinn oder andere finanzielle Folgen, den Verlust
von Daten sowie unmittelbare oder mittelbare Folgeschäden,
die durch den Gebrauch der Software verursacht wurden.

Autor/Copyright: Dr. Wilhelm Happe, Kiel
";
            mbText = String.Concat(mbText, Environment.NewLine + "Programmversion: " + curVersion + ", Build vom " + RetrieveLinkerTimestamp().ToString("d.M.yyyy"));
            MessageBox.Show(mbText, winTitle, MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void DGV_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dGV.Rows[e.RowIndex].Cells[e.ColumnIndex].EditType == typeof(DataGridViewComboBoxEditingControl))
            {
                dGV.ClearSelection();
                dGV.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "1";
                dGV.RefreshEdit(); // Aktualisiert den Wert der aktiven Zelle mit dem zugrunde liegenden Zellenwert.
            }
            else
            {
                MessageBox.Show("Fehler in Zeile " + (e.RowIndex + 1).ToString() + ", Spalte "
                    + (e.ColumnIndex + 1).ToString() + ".", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void PanelHeader_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            LinearGradientBrush myBrush = new LinearGradientBrush(new Point(0, 0), new Point(Width, Height), Color.AliceBlue, Color.LightSteelBlue);
            g.FillRectangle(myBrush, ClientRectangle);
        }

        private void ToolStripButtonNewFile_Click(object sender, EventArgs e) { NeuToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonSave_Click(object sender, EventArgs e) { SpeichernToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonOpenFile_Click(object sender, EventArgs e) { ÖffnenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonImport_Click(object sender, EventArgs e) { ImportToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonSaveAs_Click(object sender, EventArgs e) { SpeichernUnterToolStripMenuItem_Click(null, null); }
        private void ToolStripButton2_Click(object sender, EventArgs e) { ToolStripMenuItemExcel_Click(null, null); }
        private void ToolStripButtonCut_Click(object sender, EventArgs e) { AusschneidenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonPaste_Click(object sender, EventArgs e) { PasteToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonCopy_Click(object sender, EventArgs e) { KopierenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDelete_Click(object sender, EventArgs e) { LöschenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonIncrease_Click(object sender, EventArgs e) { WertErhöhenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonIncreaseMuch_Click(object sender, EventArgs e) { GroßenWertErhöhenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDecrease_Click(object sender, EventArgs e) { WertVerringernToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDecreaseMuch_Click(object sender, EventArgs e) { GroßenWertVerringernToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonNow_Click(object sender, EventArgs e) { UhrzeitDatumToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonNewRow_Click(object sender, EventArgs e) { InsertRowToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDeleteRow_Click(object sender, EventArgs e) { DeleteRowToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonSelectAll_Click(object sender, EventArgs e) { AllesAuswählenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDeleteAll_Click(object sender, EventArgs e) { AllesLöschenContextMenuItem_Click(null, null); }
        private void ToolStripButtonModus_Click(object sender, EventArgs e) { ChangeModus(); }
        private void ToolStripButtonHelp_Click(object sender, EventArgs e) { HilfeAnzeigenToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonDeskLink_Click(object sender, EventArgs e) { AddShortcutToDesktop(); }
        private void ToolStripButtonPrint_Click(object sender, EventArgs e) { ShowPrintDialog(); }
        private void ToolStripButtonPreview_Click(object sender, EventArgs e) { ToolStripMenuItemPrintPreview_Click(null, null); }
        private void ToolStripButtonExit_Click(object sender, EventArgs e) { Close(); }
        private void ToolStripButtonInfo_Click(object sender, EventArgs e) { InfoToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonWeb_Click(object sender, EventArgs e) { WebseiteToolStripMenuItem_Click(null, null); }
        private void ToolStripButtonComplete_Click(object sender, EventArgs e)
        {
            if (dGV.CurrentCell != null && dtEdits.Contains(dGV.CurrentCell.ColumnIndex)) // nicht in letzten beiden Splaten // dGV.CurrentCell.EditType == typeof(DataGridViewTextBoxEditingControl)
            {// if (!dGV.IsCurrentCellInEditMode) { dGV.BeginEdit(true); }
                var txtBox = (TextBox)dGV.EditingControl;
                DateTime dt = new DateTime();
                string dtFormat = modusStunden && dGV.CurrentCell.ColumnIndex < 2 ? "H:mm" : "d.M.yyyy";
                string strTxtBox = txtBox.TextLength == 0 ? DateTime.Now.ToString(dtFormat) : txtBox.Text;
                if (modusStunden && dGV.CurrentCell.ColumnIndex < 2) { dt = clsUtilities.NormalizeTime(strTxtBox); }
                else { dt = clsUtilities.NormalizeDate(strTxtBox); }
                if (dt != DateTime.MinValue)
                {
                    txtBox.Text = dt.ToString(dtFormat); // Textbox Inhalt ersetzten
                    txtBox.Select(txtBox.Text.Length, 0); // Cursor ans Ende setzten (Markierungslänge = 0)
                }
                else { Console.Beep(); }
            }
        }

        private void AddShortcutToDesktop()
        {
            string deskText = currFile.Length > 0 ? Path.GetFileName(currFile) : clsUtilities.GetDescription();
            string linkFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), deskText + ".lnk");
            using (var shellShortcut = new ShellShortcut(linkFileName)
            {
                Path = Application.ExecutablePath,
                WorkingDirectory = Application.StartupPath,
                Arguments = (currFile.Length > 0 ? currFile.Contains(" ") ? "\"" + currFile + "\"" : currFile : ""),
                IconPath = Application.ExecutablePath,
                IconIndex = 0,
                Description = deskText,
            })
                try
                {
                    shellShortcut.Save();
                    MessageBox.Show("Die Desktopverknüpfung wurde erfolgreich angelegt:" + Environment.NewLine + linkFileName, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
        }

        public void ToolStripMenuItemPrintPreview_Click(object sender, EventArgs e)
        {
            if (dGV.IsCurrentCellInEditMode)
            {
                dGV.EndEdit();
                dGV.CurrentCell.Selected = true;
            }
            using (frmPrintPreview f2 = new frmPrintPreview(this))
            {
                f2.Text = currFile.Length > 0 ? currFile + " - Seitenanzeige" : "Seitenanzeige";
                f2._printPreviewControl.Document = printDocument;
                f2.ShowDialog(this);
            }
        }

        private void ToolStripMenuItemPrint_Click(object sender, EventArgs e)
        {
            ShowPrintDialog();
        }

        internal void ShowPrintDialog()
        {// how to call non-static method on form1 from form2
            if (dGV.IsCurrentCellInEditMode)
            {
                dGV.EndEdit();
                dGV.CurrentCell.Selected = true;
            }
            using (frmPrintConfig f3 = new frmPrintConfig(this))
            {
                f3.ShowDialog(this);
            }
        }

        internal void PrintNowFromChild()
        {// how to call non-static method on form1 from form2
            try
            {
                if (!clsUtilities.isDGVEmpty(dGV))
                {
                    printDocument.DocumentName = currFile.Length > 0 ? Path.GetFileName(currFile) : "";
                    printDocument.Print();
                }
                else { MessageBox.Show("Es gibt nichts zu drucken!", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
            }
            catch { MessageBox.Show("Fehler beim Drucken", "Fehler"); }
        }

        public void PrintDocument_BeginPrint(object sender, PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat()
                {
                    Alignment = StringAlignment.Near,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter
                };
                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                intRow = 0;
                bFirstPage = true;
                bNewPage = true;
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dGV.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        public void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                int iLeftMargin = e.MarginBounds.Left;
                int iTopMargin = e.MarginBounds.Top;
                bool bMorePagesToPrint = false; // Whether more pages have to print or not
                int iTmpWidth = 0;
                if (bFirstPage)
                {// For the first page to print set the cell width and header height
                    foreach (DataGridViewColumn GridCol in dGV.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width / (double)iTotalWidth * (double)iTotalWidth * ((double)e.MarginBounds.Width / (double)iTotalWidth))));
                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText, GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;
                        arrColumnLefts.Add(iLeftMargin); // Save width and height of headres
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                while (intRow <= dGV.Rows.Count - 2)
                {// Loop till all the grid rows not get printed
                    DataGridViewRow GridRow = dGV.Rows[intRow];
                    iCellHeight = GridRow.Height + 5; // Set the cell height
                    int iCount = 0;
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {// Check whether the current page settings allo more rows to print
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {// Draw Header
                            e.Graphics.DrawString(currFile, new Font(dGV.Font, FontStyle.Bold), Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top - e.Graphics.MeasureString(currFile, new Font(dGV.Font, FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            String strDate = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
                            e.Graphics.DrawString(strDate, new Font(dGV.Font, FontStyle.Bold), Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width - e.Graphics.MeasureString(strDate, new Font(dGV.Font, FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top - e.Graphics.MeasureString("Customer Summary", new Font(new Font(dGV.Font, FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            iTopMargin = e.MarginBounds.Top; // Draw Columns
                            foreach (DataGridViewColumn GridCol in dGV.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray), new Rectangle((int)arrColumnLefts[iCount], iTopMargin, (int)arrColumnWidths[iCount], iHeaderHeight));
                                e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount], iTopMargin, (int)arrColumnWidths[iCount], iHeaderHeight));
                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font, new SolidBrush(GridCol.InheritedStyle.ForeColor), new RectangleF((int)arrColumnLefts[iCount], iTopMargin, (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {// Draw Columns Contents                
                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font, new SolidBrush(Cel.InheritedStyle.ForeColor), new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin, (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount], iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));
                            iCount++;
                        }
                    }
                    intRow++;
                    iTopMargin += iCellHeight;
                }
                if (bMorePagesToPrint) { e.HasMorePages = true; } // If more lines exist, print another page.
                else { e.HasMorePages = false; }

            }
            catch (Exception exc) { MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ToolStripMenuItemExcel_Click(object sender, EventArgs e)
        {// nur durch die Verfrachtung in eine Klasse, in der allein die Assembly referenziert wird, erhält man die gewünschte Fehlermeldung, falls die Referenz fehlt
            try
            {// falls Assembly nicht gefunden
                clsInteropExcel interopExcel = new clsInteropExcel();
                clsInteropExcel.WriteValues2Excel(dGV);
            }
            catch (FileNotFoundException) { MessageBox.Show("Die Funktion auf diesem System nicht verfügbar.", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private void ImportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (modusStunden)
            {
                using (frmImportIntro IntroForm = new frmImportIntro())
                {
                    if (IntroForm.ShowDialog() == DialogResult.OK)
                    {

                        if (IntroForm.impFromFile)
                        {
                            if (!ImportTextFile(IntroForm.ImportForm_fileDialog.FileName))
                            { MessageBox.Show("Die Datei »" + IntroForm.ImportForm_fileDialog.FileName + "« enthält keine verwertbare Daten.", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Information); }
                        }
                        else
                        {
                            if (!ImportClipboardText()) // MessageBox.Show("Not yet implemented");
                            { MessageBox.Show("Die Windows-Zwischenablage enthält keine für den Import geeignete Daten.", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Information); }
                        }
                    }
                }
            }
            else { MessageBox.Show("Es kann nur in den Stundenrechner importiert werden,\nnicht in den Tagerechner! Bitte ändern Sie den Modus.", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private bool ImportTextFile(string fullFilename)
        {
            try
            {
                int index = 0;
                using (StreamReader sReader = new StreamReader(fullFilename))
                {
                    String sLine; Match match;
                    string varVon = string.Empty, varBis = string.Empty, varNotiz = string.Empty, varPause = string.Empty;
                    bool newBlock = false;
                    DateTime result = DateTime.MinValue;
                    while ((sLine = sReader.ReadLine()) != null)
                    {
                        if (!String.IsNullOrEmpty(sLine) && sLine.Contains("Datum:"))
                        {
                            newBlock = true;
                            match = rgxImportDate.Match(sLine);
                            if (match.Success)
                            {
                                if (DateTime.TryParse(match.ToString(), out result))
                                {
                                    varNotiz = result.ToString("d.M.yyyy");
                                }
                            }
                        }
                        else if (newBlock && !String.IsNullOrEmpty(sLine) && sLine.Contains("Arbeitsbeginn:"))
                        {
                            match = rgxImportTime.Match(sLine);
                            if (match.Success)
                            {
                                if (DateTime.TryParse(match.ToString(), out result))
                                {
                                    varVon = result.ToString("H:mm");
                                }
                            }
                        }
                        else if (newBlock && !String.IsNullOrEmpty(sLine) && sLine.Contains("Arbeitsende:"))
                        {
                            match = rgxImportTime.Match(sLine);
                            if (match.Success)
                            {
                                if (DateTime.TryParse(match.ToString(), out result))
                                {
                                    varBis = result.ToString("H:mm");
                                }
                            }
                        }
                        else if (newBlock && !String.IsNullOrEmpty(sLine) && sLine.Contains("Pausendauer:"))
                        {// MessageBox.Show(varVon + "|" + varBis + "|" + varNotiz);
                            match = rgxImportTime.Match(sLine);
                            if (match.Success)
                            {
                                if (DateTime.TryParse(match.ToString(), out result))
                                {
                                    varPause = result.ToString("H:mm");
                                }
                            }
                            dGV.Rows.Add(varVon, varBis, varPause, "1", varNotiz);
                            UpdateLastTimeColumns();
                            newBlock = false;
                            index++;
                        }
                    }
                } //dGV.AllowUserToAddRows = true; // fügt leere Bearbeitungszeile am Ende hinzu
                if (index != 0)
                {
                    nothingToSave = false;
                    return true;
                }
                else { return false; }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Fehlermeldung", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            return true;
        }

        private bool ImportClipboardText()
        {
            try
            {
                int index = 0;
                if (Clipboard.ContainsText())
                {
                    string text = Clipboard.GetText(TextDataFormat.Text);
                    Match match;
                    if (text.Length > 1)
                    {
                        text = text.Replace("\r\n", "\r").Replace("\n", "\r"); //  unify all line breaks to \r
                        string[] textArray = text.Split('\r'); //  create an array of lines
                        string varVon = string.Empty, varBis = string.Empty, varNotiz = string.Empty, varPause = string.Empty;
                        DateTime result = DateTime.MinValue;
                        bool newBlock = false;
                        foreach (string sLine in textArray)
                        {
                            if (sLine.Contains("Datum:"))
                            {
                                newBlock = true;
                                match = rgxImportDate.Match(sLine);
                                if (match.Success)
                                {
                                    if (DateTime.TryParse(match.ToString(), out result))
                                    {
                                        varNotiz = result.ToString("d.M.yyyy");
                                    }
                                }
                            }
                            else if (newBlock && sLine.Contains("Arbeitsbeginn:"))
                            {
                                match = rgxImportTime.Match(sLine);
                                if (match.Success)
                                {
                                    if (DateTime.TryParse(match.ToString(), out result))
                                    {
                                        varVon = result.ToString("H:mm");
                                    }
                                }
                            }
                            else if (newBlock && sLine.Contains("Arbeitsende:"))
                            {
                                match = rgxImportTime.Match(sLine);
                                if (match.Success)
                                {
                                    if (DateTime.TryParse(match.ToString(), out result))
                                    {
                                        varBis = result.ToString("H:mm");
                                    }
                                }
                            }
                            else if (newBlock && sLine.Contains("Pausendauer:"))
                            {// MessageBox.Show(varVon + "|" + varBis + "|" + varNotiz);
                                match = rgxImportTime.Match(sLine);
                                if (match.Success)
                                {
                                    if (DateTime.TryParse(match.ToString(), out result))
                                    {
                                        varPause = result.ToString("H:mm");
                                    }
                                }
                                dGV.Rows.Add(varVon, varBis, varPause, "1", varNotiz);
                                UpdateLastTimeColumns();
                                newBlock = false;
                                index++;
                            }
                        }
                    }
                    if (index != 0)
                    {
                        nothingToSave = false;
                        return true;
                    }
                    else { return false; }
                }
                else { MessageBox.Show("Die Zwischenablage enhält keinen Text!", winTitle, MessageBoxButtons.OK, MessageBoxIcon.Information); return true; }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Fehlermeldung", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            return true;
        }

        private void DGV_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) { dGV.FirstDisplayedScrollingRowIndex = dGV.Rows[dGV.Rows.Count - 1].Index; }
    }
}
