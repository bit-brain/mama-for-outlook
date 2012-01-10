namespace OutlookAddInMama
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class FormRegionMama : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public FormRegionMama(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.labelPattern = new System.Windows.Forms.Label();
            this.labelLocation = new System.Windows.Forms.Label();
            this.checkBoxPreview = new System.Windows.Forms.CheckBox();
            this.buttonSend = new System.Windows.Forms.Button();
            this.textBoxPattern = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.textBoxLocation = new System.Windows.Forms.TextBox();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.labelCopyright = new System.Windows.Forms.Label();
            this.labelCopyright2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelPattern
            // 
            this.labelPattern.AutoSize = true;
            this.labelPattern.Location = new System.Drawing.Point(3, 8);
            this.labelPattern.Name = "labelPattern";
            this.labelPattern.Size = new System.Drawing.Size(41, 13);
            this.labelPattern.TabIndex = 0;
            this.labelPattern.Text = "Pattern";
            // 
            // labelLocation
            // 
            this.labelLocation.AutoSize = true;
            this.labelLocation.Location = new System.Drawing.Point(179, 8);
            this.labelLocation.Name = "labelLocation";
            this.labelLocation.Size = new System.Drawing.Size(48, 13);
            this.labelLocation.TabIndex = 2;
            this.labelLocation.Text = "Location";
            // 
            // checkBoxPreview
            // 
            this.checkBoxPreview.AutoSize = true;
            this.checkBoxPreview.Location = new System.Drawing.Point(444, 7);
            this.checkBoxPreview.Name = "checkBoxPreview";
            this.checkBoxPreview.Size = new System.Drawing.Size(64, 17);
            this.checkBoxPreview.TabIndex = 5;
            this.checkBoxPreview.Text = "Preview";
            this.toolTip.SetToolTip(this.checkBoxPreview, "Do not send directly but show each e-mail to be send. (Might be a lot!)");
            this.checkBoxPreview.UseVisualStyleBackColor = true;
            // 
            // buttonSend
            // 
            this.buttonSend.Location = new System.Drawing.Point(514, 3);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(75, 23);
            this.buttonSend.TabIndex = 6;
            this.buttonSend.Text = "Send";
            this.toolTip.SetToolTip(this.buttonSend, "Do it. DO IT!");
            this.buttonSend.UseVisualStyleBackColor = true;
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // textBoxPattern
            // 
            this.textBoxPattern.Location = new System.Drawing.Point(50, 5);
            this.textBoxPattern.Name = "textBoxPattern";
            this.textBoxPattern.Size = new System.Drawing.Size(123, 20);
            this.textBoxPattern.TabIndex = 1;
            this.toolTip.SetToolTip(this.textBoxPattern, "The regex pattern which files are sent and what is needed for replacement.");
            // 
            // folderBrowserDialog
            // 
            this.folderBrowserDialog.ShowNewFolderButton = false;
            // 
            // textBoxLocation
            // 
            this.textBoxLocation.Location = new System.Drawing.Point(234, 5);
            this.textBoxLocation.Name = "textBoxLocation";
            this.textBoxLocation.Size = new System.Drawing.Size(123, 20);
            this.textBoxLocation.TabIndex = 3;
            this.toolTip.SetToolTip(this.textBoxLocation, "Where to look for the attachement files.");
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(363, 3);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowse.TabIndex = 4;
            this.buttonBrowse.Text = "Browse...";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // labelCopyright
            // 
            this.labelCopyright.AutoSize = true;
            this.labelCopyright.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelCopyright.Location = new System.Drawing.Point(607, 2);
            this.labelCopyright.Name = "labelCopyright";
            this.labelCopyright.Size = new System.Drawing.Size(214, 13);
            this.labelCopyright.TabIndex = 8;
            this.labelCopyright.Text = "Mama - Mass Attachement Mailing Assistant";
            this.toolTip.SetToolTip(this.labelCopyright, "It\'s dedicated to all those mamas out there. Hey, you do a great job, madams!");
            // 
            // labelCopyright2
            // 
            this.labelCopyright2.AutoSize = true;
            this.labelCopyright2.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelCopyright2.Location = new System.Drawing.Point(607, 17);
            this.labelCopyright2.Name = "labelCopyright2";
            this.labelCopyright2.Size = new System.Drawing.Size(138, 13);
            this.labelCopyright2.TabIndex = 9;
            this.labelCopyright2.Text = "by Tim Mueller / bit-brain.de";
            this.toolTip.SetToolTip(this.labelCopyright2, "tim enjoyed building this piece of software. Hopefully you enjoy using it.");
            // 
            // FormRegionMama
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelCopyright2);
            this.Controls.Add(this.labelCopyright);
            this.Controls.Add(this.buttonBrowse);
            this.Controls.Add(this.textBoxLocation);
            this.Controls.Add(this.textBoxPattern);
            this.Controls.Add(this.buttonSend);
            this.Controls.Add(this.checkBoxPreview);
            this.Controls.Add(this.labelLocation);
            this.Controls.Add(this.labelPattern);
            this.Name = "FormRegionMama";
            this.Size = new System.Drawing.Size(825, 32);
            this.FormRegionShowing += new System.EventHandler(this.FormRegionMama_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.FormRegionMama_FormRegionClosed);
            this.Load += new System.EventHandler(this.FormRegionMama_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "Mama";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;
            manifest.ShowInspectorRead = false;
            manifest.ShowReadingPane = false;

        }

        #endregion

        private System.Windows.Forms.Label labelPattern;
        private System.Windows.Forms.Label labelLocation;
        private System.Windows.Forms.CheckBox checkBoxPreview;
        private System.Windows.Forms.Button buttonSend;
        private System.Windows.Forms.TextBox textBoxPattern;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.TextBox textBoxLocation;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Label labelCopyright;
        private System.Windows.Forms.Label labelCopyright2;

        public partial class FormRegionMamaFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public FormRegionMamaFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                FormRegionMama.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.FormRegionMamaFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                FormRegionMama form = new FormRegionMama(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal FormRegionMama FormRegionMama
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(FormRegionMama))
                        return (FormRegionMama)item;
                }
                return null;
            }
        }
    }
}
