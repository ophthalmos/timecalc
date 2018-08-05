using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace TimeCalc
{
    public partial class frmImportIntro : Form
    {
        public frmImportIntro()
        {
            InitializeComponent();
        }
        // Control als Eigenschaft offenlegen (OpenFileDialog in diesem Formular!):
        internal OpenFileDialog ImportForm_fileDialog { get { return importFileDialog; } }
        internal bool impFromFile = false;

        private void btnFileImport_Click(object sender, EventArgs e)
        {
            impFromFile = true;
            if (importFileDialog.ShowDialog() == DialogResult.OK) { this.DialogResult = DialogResult.OK; }
        }

        private void btnClipboardImport_Click(object sender, EventArgs e)
        {
            impFromFile = false;
            this.DialogResult = DialogResult.OK;
        }

        private void frmImportIntro_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { this.DialogResult = DialogResult.Cancel; }
        }

        private void pictureBoxBMAS_Click(object sender, EventArgs e)
        {
            try { Process.Start("http://www.der-mindestlohn-wirkt.de/ml/DE/Service/App-Zeiterfassung/inhalt.html"); }
            catch (Exception ex) { MessageBox.Show(ex.Message, clsUtilities.GetDescription(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private void pictureBoxAppStore_Click(object sender, EventArgs e)
        {
            try { Process.Start("https://itunes.apple.com/de/app/bmas-app-einfach-erfasst/id1012872512?l=de&mt=10"); }
            catch (Exception ex) { MessageBox.Show(ex.Message, clsUtilities.GetDescription(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private void pictureBoxPlayStore_Click(object sender, EventArgs e)
        {
            try { Process.Start("https://play.google.com/store/apps/details?id=de.bmas.einfach_erfasst"); }
            catch (Exception ex) { MessageBox.Show(ex.Message, clsUtilities.GetDescription(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }

        private void pictureBoxWinStore_Click(object sender, EventArgs e)
        {
            try { Process.Start("https://www.microsoft.com/de-de/store/apps/bmas-app-einfach-erfasst/9nblggh682df"); }
            catch (Exception ex) { MessageBox.Show(ex.Message, clsUtilities.GetDescription(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
        }
    }
}
