using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ReleasePack.Engine;

namespace ReleasePack.AddIn.UI
{
    /// <summary>
    /// WinForms UserControl docked in the SolidWorks TaskPane.
    /// Provides the complete Release Pack configuration UI.
    /// </summary>
    [ComVisible(true)]
    [ProgId("ReleasePack.AddIn.UI.ReleasePackTaskPane")]
    [Guid("D4F2E6A1-8B3C-4D7E-A5F9-1C2B3D4E5F60")]
    public class ReleasePackTaskPane : UserControl, IProgressCallback
    {
        private ISldWorks _swApp;

        // ── UI Controls ──────────────────────────────
        private Panel _headerPanel;
        private Label _titleLabel;

        // Scope
        private GroupBox _scopeGroup;
        private RadioButton _rbCurrent;
        private RadioButton _rbChildren;
        private RadioButton _rbRemote;
        // Metadata Controls
        private TextBox _txtCompany;
        private TextBox _txtProject;
        private TextBox _txtDrawnBy;
        private TextBox _txtCheckedBy;
        private Button _btnBrowse;
        private TextBox _txtRemotePath;

        // Outputs
        private GroupBox _outputGroup;
        private CheckBox _chkDrawing;
        private CheckBox _chkPDF;
        private CheckBox _chkDXF;
        private CheckBox _chkSTEP;
        private CheckBox _chkParasolid;
        private CheckBox _chkBOM;
        private CheckBox _chkPreview;

        // Options
        private GroupBox _optionsGroup;
        private ComboBox _cmbSheetSize;
        private ComboBox _cmbViewStandard;
        private ComboBox _cmbDimMode;
        private ComboBox _cmbComplexity;
        private TextBox _txtBomTemplate;
        private Button _btnBrowseBom;

        // Output folder
        private GroupBox _folderGroup;
        private RadioButton _rbFolderAuto;
        private RadioButton _rbFolderCustom;
        private TextBox _txtCustomFolder;
        private Button _btnBrowseFolder;

        // Actions
        private Button _btnPreview;
        private Button _btnGenerate;
        private ProgressBar _progressBar;
        private Label _lblStatus;
        private RichTextBox _logBox;

        public ReleasePackTaskPane()
        {
            InitializeUI();
        }

        public void Initialize(ISldWorks swApp)
        {
            _swApp = swApp;

            // Auto-select based on active doc type
            if (_swApp.ActiveDoc is AssemblyDoc)
            {
                _rbChildren.Checked = true;
                _rbCurrent.Checked = false;
            }
            else
            {
                _rbCurrent.Checked = true;
                _rbChildren.Checked = false;
            }
        }

        #region UI Construction

        private void InitializeUI()
        {
            this.SuspendLayout();
            // Match SolidWorks PropertyManager native background
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.AutoScroll = true;
            this.Dock = DockStyle.Fill;
            this.Font = new Font("Segoe UI", 9F);

            int y = 0;

            // ── Header ──────────────────────────────────
            _headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.FromArgb(0, 122, 204), // SolidWorks/VS Accent Blue
                Padding = new Padding(10, 0, 0, 0)
            };

            _titleLabel = new Label
            {
                Text = "Release Pack Studio",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            _headerPanel.Controls.Add(_titleLabel);
            this.Controls.Add(_headerPanel);

            // Container for scrollable content
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(8)
            };

            y = 8;

            // ── Scope ───────────────────────────────────
            _scopeGroup = CreateGroupBox("Scope", y, 130);
            y += 4;

            _rbCurrent = new RadioButton
            {
                Text = "Current Document Only",
                Location = new Point(12, 22),
                Size = new Size(200, 20),
                Checked = true 
            };

            _rbChildren = new RadioButton
            {
                Text = "Current + All Children (Assembly)",
                Location = new Point(12, 44),
                Size = new Size(200, 20)
            };

            // ── Project Details Group ────────────────────────────────────────────────
            var grpProject = new GroupBox
            {
                Text = "Project Metadata (Title Block)",
                Location = new Point(10, 100),
                Size = new Size(280, 140)
            };
            this.Controls.Add(grpProject);

            int lblX = 10, txtX = 90;
            y = 20; // Reset y for this group
            int step = 28;

            // Company
            grpProject.Controls.Add(new Label { Text = "Company:", Location = new Point(lblX, y + 3), AutoSize = true });
            _txtCompany = new TextBox { Location = new Point(txtX, y), Size = new Size(180, 20) };
            grpProject.Controls.Add(_txtCompany);

            // Project
            y += step;
            grpProject.Controls.Add(new Label { Text = "Project:", Location = new Point(lblX, y + 3), AutoSize = true });
            _txtProject = new TextBox { Location = new Point(txtX, y), Size = new Size(180, 20) };
            grpProject.Controls.Add(_txtProject);

            // Drawn By
            y += step;
            grpProject.Controls.Add(new Label { Text = "Drawn By:", Location = new Point(lblX, y + 3), AutoSize = true });
            _txtDrawnBy = new TextBox { Location = new Point(txtX, y), Size = new Size(180, 20), Text = System.Environment.UserName };
            grpProject.Controls.Add(_txtDrawnBy);

            // Checked By
            y += step;
            grpProject.Controls.Add(new Label { Text = "Checked By:", Location = new Point(lblX, y + 3), AutoSize = true });
            _txtCheckedBy = new TextBox { Location = new Point(txtX, y), Size = new Size(180, 20) };
            grpProject.Controls.Add(_txtCheckedBy);


            // ── Output Options Group ─────────────────────────────────────────────────
            var grpOutput = new GroupBox
            {
                Text = "Output Options",
                Location = new Point(10, 250),
                Size = new Size(280, 100)
            };

            // ── Display Options Group ────────────────────────────────────────────────
            var grpDisplay = new GroupBox
            {
                Text = "Drawing Options",
                Location = new Point(10, 360),
                Size = new Size(280, 115)
            };
            this.Controls.Add(grpDisplay);

            _progressBar = new ProgressBar
            {
                Location = new Point(10, 500),
                Size = new Size(280, 15),
                Visible = false
            };
            this.Controls.Add(_progressBar);

            _lblStatus = new Label
            {
                Text = "Ready",
                Location = new Point(10, 520),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            this.Controls.Add(_lblStatus);

            _rbRemote = new RadioButton
            {
                Text = "Select Remote File",
                Location = new Point(12, 66),
                Size = new Size(200, 20)
            };
            _rbRemote.CheckedChanged += (s, e) =>
            {
                _txtRemotePath.Enabled = _rbRemote.Checked;
                _btnBrowse.Enabled = _rbRemote.Checked;
            };

            _txtRemotePath = new TextBox
            {
                Location = new Point(12, 90),
                Size = new Size(180, 23),
                Enabled = false,
                BorderStyle = BorderStyle.FixedSingle
            };

            _btnBrowse = new Button
            {
                Text = "...",
                Location = new Point(196, 89),
                Size = new Size(32, 23),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(245, 245, 245)
            };
            _btnBrowse.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            _btnBrowse.Click += BtnBrowseFile_Click;

            _scopeGroup.Controls.AddRange(new Control[]
                { _rbCurrent, _rbChildren, _rbRemote, _txtRemotePath, _btnBrowse });
            contentPanel.Controls.Add(_scopeGroup);

            // ── Output Types ────────────────────────────
            y = 146;
            _outputGroup = CreateGroupBox("Output Types", y, 120);

            _chkDrawing = CreateCheckBox("Drawing (.slddrw)", 22, true);
            _chkPDF = CreateCheckBox("PDF", 42, true);
            _chkDXF = CreateCheckBox("DXF", 62, true);
            _chkSTEP = CreateCheckBox("STEP (.stp)", 82, false);
            _chkParasolid = CreateCheckBox("Parasolid (.x_t)", 22, false, 130);
            _chkBOM = CreateCheckBox("BOM Excel (.xlsx)", 42, true, 130);
            _chkPreview = CreateCheckBox("Preview (.png)", 62, false, 130);

            _outputGroup.Controls.AddRange(new Control[]
                { _chkDrawing, _chkPDF, _chkDXF, _chkSTEP, _chkParasolid, _chkBOM, _chkPreview });
            contentPanel.Controls.Add(_outputGroup);

            // ── Options ─────────────────────────────────
            y = 274;
            _optionsGroup = CreateGroupBox("Drawing Options", y, 170);

            var lblSheet = new Label { Text = "Sheet Size:", Location = new Point(12, 24), Size = new Size(70, 20) };
            _cmbSheetSize = new ComboBox
            {
                Location = new Point(85, 22),
                Size = new Size(140, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbSheetSize.Items.AddRange(new[] { "Auto", "A4 Landscape", "A3 Landscape", "A2 Landscape", "A1 Landscape" });
            _cmbSheetSize.SelectedIndex = 0;

            var lblStd = new Label { Text = "Standard:", Location = new Point(12, 52), Size = new Size(70, 20) };
            _cmbViewStandard = new ComboBox
            {
                Location = new Point(85, 50),
                Size = new Size(140, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbViewStandard.Items.AddRange(new[] { "3rd Angle (ANSI)", "1st Angle (ISO)" });
            _cmbViewStandard.SelectedIndex = 0;

            var lblDimMode = new Label { Text = "Dim Mode:", Location = new Point(12, 80), Size = new Size(65, 20) };
            _cmbDimMode = new ComboBox
            {
                Location = new Point(80, 78),
                Size = new Size(145, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbDimMode.Items.AddRange(new[] { "Full Auto", "Model Dimensions", "Hybrid Auto" });
            _cmbDimMode.SelectedIndex = 0;

            var lblBom = new Label { Text = "BOM Tpl:", Location = new Point(12, 108), Size = new Size(60, 20) };
            _txtBomTemplate = new TextBox { Location = new Point(75, 106), Size = new Size(115, 23) };
            _btnBrowseBom = new Button { Text = "...", Location = new Point(195, 105), Size = new Size(25, 25) };
            _btnBrowseBom.Click += BtnBrowseBom_Click;

            var lblComplexity = new Label { Text = "Complexity:", Location = new Point(12, 136), Size = new Size(65, 20) };
            _cmbComplexity = new ComboBox
            {
                Location = new Point(80, 134),
                Size = new Size(140, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbComplexity.Items.AddRange(new[] { "Simple", "Moderate", "High-Density" });
            _cmbComplexity.SelectedIndex = 1;

            _optionsGroup.Controls.AddRange(new Control[] { lblSheet, _cmbSheetSize, lblStd, _cmbViewStandard, lblDimMode, _cmbDimMode, lblBom, _txtBomTemplate, _btnBrowseBom, lblComplexity, _cmbComplexity });
            contentPanel.Controls.Add(_optionsGroup);

            // ── Output Folder ───────────────────────────
            y = 455;
            _folderGroup = CreateGroupBox("Save Location", y, 90);

            _rbFolderAuto = new RadioButton
            {
                Text = "Auto (next to source file)",
                Location = new Point(12, 22),
                Size = new Size(220, 20),
                Checked = true
            };
            _rbFolderCustom = new RadioButton
            {
                Text = "Custom folder:",
                Location = new Point(12, 44),
                Size = new Size(120, 20)
            };
            _rbFolderCustom.CheckedChanged += (s, e) =>
            {
                _txtCustomFolder.Enabled = _rbFolderCustom.Checked;
                _btnBrowseFolder.Enabled = _rbFolderCustom.Checked;
            };

            _txtCustomFolder = new TextBox
            {
                Location = new Point(12, 66),
                Size = new Size(180, 23),
                Enabled = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            _btnBrowseFolder = new Button
            {
                Text = "...",
                Location = new Point(196, 65),
                Size = new Size(32, 23),
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(245, 245, 245)
            };
            _btnBrowseFolder.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            _btnBrowseFolder.Click += BtnBrowseFolder_Click;

            _folderGroup.Controls.AddRange(new Control[]
                { _rbFolderAuto, _rbFolderCustom, _txtCustomFolder, _btnBrowseFolder });
            contentPanel.Controls.Add(_folderGroup);

            // ── Generate & Preview Buttons ─────────────────────────
            y = 555;
            
            _btnPreview = new Button
            {
                Text = "Preview Layout",
                Location = new Point(8, y),
                Size = new Size(110, 36),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand
            };
            _btnPreview.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            
            _btnGenerate = new Button
            {
                Text = "Generate Tasks",
                Location = new Point(126, y),
                Size = new Size(114, 36),
                BackColor = Color.FromArgb(0, 122, 204), // Native Action Blue
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand
            };
            _btnGenerate.FlatAppearance.BorderSize = 0;
            _btnGenerate.FlatAppearance.MouseOverBackColor = Color.FromArgb(28, 151, 234);
            _btnGenerate.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 90, 158);
            _btnGenerate.Click += BtnGenerate_Click;
            
            contentPanel.Controls.Add(_btnPreview);
            contentPanel.Controls.Add(_btnGenerate);

            // ── Progress Bar ────────────────────────────
            y = 600;
            _progressBar = new ProgressBar
            {
                Location = new Point(8, y),
                Size = new Size(232, 16),
                Style = ProgressBarStyle.Continuous
            };
            contentPanel.Controls.Add(_progressBar);

            // ── Log Box ─────────────────────────────────
            y = 622;
            _logBox = new RichTextBox
            {
                Location = new Point(8, y),
                Size = new Size(232, 200),
                ReadOnly = true,
                BackColor = Color.FromArgb(30, 35, 45),
                ForeColor = Color.FromArgb(200, 210, 220),
                Font = new Font("Consolas", 8F),
                BorderStyle = BorderStyle.None
            };
            contentPanel.Controls.Add(_logBox);

            this.Controls.Add(contentPanel);
            contentPanel.BringToFront(); // Above header

            this.ResumeLayout(false);
        }

        private GroupBox CreateGroupBox(string title, int y, int height)
        {
            return new GroupBox
            {
                Text = title,
                Location = new Point(8, y),
                Size = new Size(232, height),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };
        }

        private CheckBox CreateCheckBox(string text, int y, bool isChecked, int x = 12)
        {
            return new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(115, 20),
                Checked = isChecked,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Black
            };
        }

        #endregion

        #region Event Handlers

        private void BtnBrowseFile_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "SolidWorks Files|*.sldprt;*.sldasm|Parts|*.sldprt|Assemblies|*.sldasm";
                dlg.Title = "Select SolidWorks File";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _txtRemotePath.Text = dlg.FileName;
                }
            }
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select output folder for Release Pack";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _txtCustomFolder.Text = dlg.SelectedPath;
                }
            }
        }

        private void BtnBrowseBom_Click(object sender, EventArgs e)
        {
             using (var dlg = new OpenFileDialog())
             {
                 dlg.Filter = "BOM Templates (*.sldbomtbt)|*.sldbomtbt|All Files (*.*)|*.*";
                 dlg.Title = "Select BOM Template";
                 if (dlg.ShowDialog() == DialogResult.OK)
                 {
                     _txtBomTemplate.Text = dlg.FileName;
                 }
             }
        }

        private async void BtnGenerate_Click(object sender, EventArgs e)
        {
            if (_swApp == null)
            {
                MessageBox.Show("SolidWorks is not connected.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate
            if (_swApp.ActiveDoc == null && !_rbRemote.Checked)
            {
                MessageBox.Show("No active document. Please open a file or select a remote file.",
                    "No Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Build options
            var options = BuildOptions();

            // ── Interactive Selection Check ──
            if (!_rbRemote.Checked && _swApp.ActiveDoc != null)
            {
                var dict = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                ModelDoc2 doc = (ModelDoc2)_swApp.ActiveDoc;
                SelectionMgr selMgr = (SelectionMgr)doc.SelectionManager;
                
                if (selMgr != null && selMgr.GetSelectedObjectCount2(-1) > 0)
                {
                    for (int i = 1; i <= selMgr.GetSelectedObjectCount2(-1); i++)
                    {
                        int type = selMgr.GetSelectedObjectType3(i, -1);
                        if (type == (int)swSelectType_e.swSelCOMPONENTS)
                        {
                            Component2 comp = (Component2)selMgr.GetSelectedObjectsComponent4(i, -1);
                            if (comp != null)
                            {
                                string path = comp.GetPathName();
                                if (!string.IsNullOrEmpty(path))
                                    dict.Add(path);
                            }
                        }
                        else 
                        {
                            // If user selected faces/edges, try to get the component they belong to
                            Component2 comp = (Component2)selMgr.GetSelectedObjectsComponent4(i, -1);
                            if (comp != null)
                            {
                                string path = comp.GetPathName();
                                if (!string.IsNullOrEmpty(path))
                                    dict.Add(path);
                            }
                            else if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                            {
                                // If inside a part document and body/face is selected, assume they want just this part
                                dict.Add(doc.GetPathName());
                            }
                        }
                    }

                    if (dict.Count > 0)
                    {
                        var result = MessageBox.Show(
                            $"You have {dict.Count} specific component(s) selected in the model.\n\n" +
                            "Would you like to generate the Release Pack ONLY for the selected components?\n\n" +
                            "• 'Yes': Export only selections\n" +
                            "• 'No': Export the entire assembly",
                            "Selective Export", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (result == DialogResult.Cancel)
                        {
                            return; // Abort entirely
                        }
                        else if (result == DialogResult.Yes)
                        {
                            options.SelectedComponentPaths = dict;
                        }
                    }
                }
            }

            // Disable UI
            _btnGenerate.Enabled = false;
            _btnGenerate.Text = "⏳ Generating...";
            _progressBar.Value = 0;
            _logBox.Clear();

            try
            {
                LogMessage("Starting Release Pack generation...\n");

                // Run on background thread (SolidWorks API must be called from main thread though)
                var pipeline = new ExportPipeline(_swApp, this);
                var results = pipeline.Execute(options);

                // Show summary
                if (results != null && results.Count > 0)
                {
                    int success = 0, failed = 0;
                    foreach (var r in results)
                    {
                        if (r.Success) success++; else failed++;
                    }

                    MessageBox.Show(
                        $"Release Pack Complete!\n\n" +
                        $"Successful exports: {success}\n" +
                        $"Failed exports: {failed}\n\n" +
                        $"Output folder: {(options.UseCustomFolder ? options.OutputFolder : "(auto)")}",
                        "Release Pack", MessageBoxButtons.OK,
                        failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LogError($"Fatal error: {ex.Message}");
                MessageBox.Show($"Error: {ex.Message}", "Release Pack Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnGenerate.Enabled = true;
                _btnGenerate.Text = "Generate Tasks";
            }
        }

        #endregion

        #region Build Options

        private ExportOptions BuildOptions()
        {
            var options = new ExportOptions
            {
                // Scope
                Scope = _rbChildren.Checked ? ExportScope.CurrentAndChildren :
                        _rbRemote.Checked ? ExportScope.RemoteFile :
                        ExportScope.CurrentDocument,

                // Metadata
                CompanyName = _txtCompany.Text,
                ProjectName = _txtProject.Text,
                DrawnBy = _txtDrawnBy.Text,
                CheckedBy = _txtCheckedBy.Text,

                RemoteFilePath = _rbRemote.Checked ? _txtRemotePath.Text : null,
                GenerateDrawing = _chkDrawing.Checked,
                ExportPDF = _chkPDF.Checked,
                ExportDXF = _chkDXF.Checked,
                ExportSTEP = _chkSTEP.Checked,
                ExportParasolid = _chkParasolid.Checked,
                ExportBOM = _chkBOM.Checked,
                ExportPreviewImage = _chkPreview.Checked,

                // Drawing options
                SheetSize = MapSheetSize(_cmbSheetSize.SelectedIndex),
                ViewStandard = _cmbViewStandard.SelectedIndex == 0 ? ViewStandard.ThirdAngle : ViewStandard.FirstAngle,
                DimensionMode = _cmbDimMode.SelectedIndex == 1 ? DimensionMode.ModelDimensions :
                                _cmbDimMode.SelectedIndex == 2 ? DimensionMode.HybridAuto :
                                DimensionMode.FullAuto,
                BomTemplatePath = _txtBomTemplate.Text,

                // Folder
                UseCustomFolder = _rbFolderCustom.Checked,
                OutputFolder = _rbFolderCustom.Checked ? _txtCustomFolder.Text : null
            };

            return options;
        }

        private SheetSizeOption MapSheetSize(int index)
        {
            switch (index)
            {
                case 0: return SheetSizeOption.Auto;
                case 1: return SheetSizeOption.A4_Landscape;
                case 2: return SheetSizeOption.A3_Landscape;
                case 3: return SheetSizeOption.A2_Landscape;
                case 4: return SheetSizeOption.A1_Landscape;
                default: return SheetSizeOption.Auto;
            }
        }

        #endregion

        #region IProgressCallback

        public void ReportProgress(int percent, string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => ReportProgress(percent, message)));
                return;
            }

            _progressBar.Value = Math.Min(percent, 100);
            LogMessage(message);
        }

        public void LogMessage(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => LogMessage(message)));
                return;
            }

            _logBox.SelectionColor = Color.FromArgb(170, 210, 240);
            _logBox.AppendText(message + "\n");
            _logBox.ScrollToCaret();
        }

        public void LogWarning(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => LogWarning(message)));
                return;
            }

            _logBox.SelectionColor = Color.FromArgb(255, 200, 80);
            _logBox.AppendText("⚠ " + message + "\n");
            _logBox.ScrollToCaret();
        }

        public void LogError(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => LogError(message)));
                return;
            }

            _logBox.SelectionColor = Color.FromArgb(255, 100, 100);
            _logBox.AppendText("✗ " + message + "\n");
            _logBox.ScrollToCaret();
        }

        #endregion
    }
}
