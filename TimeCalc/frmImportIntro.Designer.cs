namespace TimeCalc
{
    partial class frmImportIntro
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmImportIntro));
            this.lblImportInfo = new System.Windows.Forms.Label();
            this.btnFileImport = new System.Windows.Forms.Button();
            this.btnClipboardImport = new System.Windows.Forms.Button();
            this.groupBoxImport = new System.Windows.Forms.GroupBox();
            this.importFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.pictureBoxWinStore = new System.Windows.Forms.PictureBox();
            this.pictureBoxPlayStore = new System.Windows.Forms.PictureBox();
            this.pictureBoxAppStore = new System.Windows.Forms.PictureBox();
            this.pictureBoxBMAS = new System.Windows.Forms.PictureBox();
            this.lblImportLogos = new System.Windows.Forms.Label();
            this.groupBoxImport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWinStore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPlayStore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAppStore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxBMAS)).BeginInit();
            this.SuspendLayout();
            // 
            // lblImportInfo
            // 
            this.lblImportInfo.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.lblImportInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblImportInfo.Location = new System.Drawing.Point(0, 0);
            this.lblImportInfo.Name = "lblImportInfo";
            this.lblImportInfo.Padding = new System.Windows.Forms.Padding(3, 3, 2, 2);
            this.lblImportInfo.Size = new System.Drawing.Size(284, 204);
            this.lblImportInfo.TabIndex = 0;
            this.lblImportInfo.Text = resources.GetString("lblImportInfo.Text");
            // 
            // btnFileImport
            // 
            this.btnFileImport.Location = new System.Drawing.Point(3, 18);
            this.btnFileImport.Name = "btnFileImport";
            this.btnFileImport.Size = new System.Drawing.Size(120, 28);
            this.btnFileImport.TabIndex = 1;
            this.btnFileImport.Text = "Aus Datei (E-Mail) ...";
            this.btnFileImport.UseVisualStyleBackColor = true;
            this.btnFileImport.Click += new System.EventHandler(this.btnFileImport_Click);
            // 
            // btnClipboardImport
            // 
            this.btnClipboardImport.Location = new System.Drawing.Point(139, 18);
            this.btnClipboardImport.Name = "btnClipboardImport";
            this.btnClipboardImport.Size = new System.Drawing.Size(140, 28);
            this.btnClipboardImport.TabIndex = 2;
            this.btnClipboardImport.Text = "Text aus Zwischenablage";
            this.btnClipboardImport.UseVisualStyleBackColor = true;
            this.btnClipboardImport.Click += new System.EventHandler(this.btnClipboardImport_Click);
            // 
            // groupBoxImport
            // 
            this.groupBoxImport.Controls.Add(this.btnFileImport);
            this.groupBoxImport.Controls.Add(this.btnClipboardImport);
            this.groupBoxImport.Location = new System.Drawing.Point(1, 211);
            this.groupBoxImport.Name = "groupBoxImport";
            this.groupBoxImport.Size = new System.Drawing.Size(282, 55);
            this.groupBoxImport.TabIndex = 3;
            this.groupBoxImport.TabStop = false;
            this.groupBoxImport.Text = "Daten importieren";
            // 
            // importFileDialog
            // 
            this.importFileDialog.DefaultExt = "*.eml";
            this.importFileDialog.Filter = "E-Mail (*.eml)|*.eml|Textdateien (*.txt)|*.txt|Alle Dateien (*.*)|*.*";
            this.importFileDialog.InitialDirectory = "Environment.SpecialFolder.Downloads";
            this.importFileDialog.RestoreDirectory = true;
            // 
            // pictureBoxWinStore
            // 
            this.pictureBoxWinStore.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxWinStore.Image = global::TimeCalc.Properties.Resources.logo_windows_store;
            this.pictureBoxWinStore.Location = new System.Drawing.Point(140, 364);
            this.pictureBoxWinStore.Name = "pictureBoxWinStore";
            this.pictureBoxWinStore.Size = new System.Drawing.Size(140, 36);
            this.pictureBoxWinStore.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBoxWinStore.TabIndex = 6;
            this.pictureBoxWinStore.TabStop = false;
            this.pictureBoxWinStore.Click += new System.EventHandler(this.pictureBoxWinStore_Click);
            // 
            // pictureBoxPlayStore
            // 
            this.pictureBoxPlayStore.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxPlayStore.Image = global::TimeCalc.Properties.Resources.logo_google_play;
            this.pictureBoxPlayStore.Location = new System.Drawing.Point(140, 321);
            this.pictureBoxPlayStore.Name = "pictureBoxPlayStore";
            this.pictureBoxPlayStore.Size = new System.Drawing.Size(140, 42);
            this.pictureBoxPlayStore.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBoxPlayStore.TabIndex = 5;
            this.pictureBoxPlayStore.TabStop = false;
            this.pictureBoxPlayStore.Click += new System.EventHandler(this.pictureBoxPlayStore_Click);
            // 
            // pictureBoxAppStore
            // 
            this.pictureBoxAppStore.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxAppStore.Image = global::TimeCalc.Properties.Resources.logo_apple_app_store;
            this.pictureBoxAppStore.Location = new System.Drawing.Point(140, 279);
            this.pictureBoxAppStore.Name = "pictureBoxAppStore";
            this.pictureBoxAppStore.Size = new System.Drawing.Size(140, 41);
            this.pictureBoxAppStore.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBoxAppStore.TabIndex = 4;
            this.pictureBoxAppStore.TabStop = false;
            this.pictureBoxAppStore.Click += new System.EventHandler(this.pictureBoxAppStore_Click);
            // 
            // pictureBoxBMAS
            // 
            this.pictureBoxBMAS.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxBMAS.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxBMAS.Image")));
            this.pictureBoxBMAS.Location = new System.Drawing.Point(4, 279);
            this.pictureBoxBMAS.Name = "pictureBoxBMAS";
            this.pictureBoxBMAS.Size = new System.Drawing.Size(120, 120);
            this.pictureBoxBMAS.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBoxBMAS.TabIndex = 7;
            this.pictureBoxBMAS.TabStop = false;
            this.pictureBoxBMAS.Click += new System.EventHandler(this.pictureBoxBMAS_Click);
            // 
            // lblImportLogos
            // 
            this.lblImportLogos.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.lblImportLogos.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblImportLogos.Location = new System.Drawing.Point(0, 271);
            this.lblImportLogos.Name = "lblImportLogos";
            this.lblImportLogos.Size = new System.Drawing.Size(284, 133);
            this.lblImportLogos.TabIndex = 8;
            // 
            // frmImportIntro
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 404);
            this.Controls.Add(this.pictureBoxBMAS);
            this.Controls.Add(this.pictureBoxAppStore);
            this.Controls.Add(this.pictureBoxPlayStore);
            this.Controls.Add(this.pictureBoxWinStore);
            this.Controls.Add(this.lblImportLogos);
            this.Controls.Add(this.groupBoxImport);
            this.Controls.Add(this.lblImportInfo);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmImportIntro";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Datenimport";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmImportIntro_KeyDown);
            this.groupBoxImport.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxWinStore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPlayStore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxAppStore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxBMAS)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblImportInfo;
        private System.Windows.Forms.Button btnFileImport;
        private System.Windows.Forms.Button btnClipboardImport;
        private System.Windows.Forms.GroupBox groupBoxImport;
        private System.Windows.Forms.OpenFileDialog importFileDialog;
        private System.Windows.Forms.PictureBox pictureBoxAppStore;
        private System.Windows.Forms.PictureBox pictureBoxPlayStore;
        private System.Windows.Forms.PictureBox pictureBoxWinStore;
        private System.Windows.Forms.PictureBox pictureBoxBMAS;
        private System.Windows.Forms.Label lblImportLogos;
    }
}