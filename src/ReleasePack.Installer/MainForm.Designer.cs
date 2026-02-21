namespace ReleasePack.Installer
{
    partial class MainForm
    {
        // Field removed to satisfy compiler warning CS0414
        private System.Windows.Forms.Label _lblStatus;
        private System.Windows.Forms.Button _btnInstall;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.PictureBox _bannerBox;
        private System.Windows.Forms.TextBox _txtSwPath;
        private System.Windows.Forms.Button _btnBrowse;
        private System.Windows.Forms.Label _lblPath;

        private void InitializeComponent()
        {
            this._lblStatus = new System.Windows.Forms.Label();
            this._btnInstall = new System.Windows.Forms.Button();
            this._progressBar = new System.Windows.Forms.ProgressBar();
            this._bannerBox = new System.Windows.Forms.PictureBox();
            this._txtSwPath = new System.Windows.Forms.TextBox();
            this._btnBrowse = new System.Windows.Forms.Button();
            this._lblPath = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(this._bannerBox)).BeginInit();
            this.SuspendLayout();
            
            // _bannerBox
            this._bannerBox.Dock = System.Windows.Forms.DockStyle.Top;
            this._bannerBox.Location = new System.Drawing.Point(0, 0);
            this._bannerBox.Name = "_bannerBox";
            this._bannerBox.Size = new System.Drawing.Size(584, 100);
            this._bannerBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this._bannerBox.TabStop = false;
            string bannerPath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Banner.png");
            if (System.IO.File.Exists(bannerPath)) this._bannerBox.Image = System.Drawing.Image.FromFile(bannerPath);
            else this._bannerBox.BackColor = System.Drawing.Color.FromArgb(40, 60, 80);

            // _lblPath
            this._lblPath.AutoSize = true;
            this._lblPath.Location = new System.Drawing.Point(12, 110);
            this._lblPath.Name = "_lblPath";
            this._lblPath.Size = new System.Drawing.Size(150, 13);
            this._lblPath.Text = "SolidWorks Executable Path:";

            // _txtSwPath
            this._txtSwPath.Location = new System.Drawing.Point(15, 130);
            this._txtSwPath.Name = "_txtSwPath";
            this._txtSwPath.Size = new System.Drawing.Size(490, 20);
            this._txtSwPath.TabIndex = 4;

            // _btnBrowse
            this._btnBrowse.Location = new System.Drawing.Point(515, 129);
            this._btnBrowse.Name = "_btnBrowse";
            this._btnBrowse.Size = new System.Drawing.Size(50, 22);
            this._btnBrowse.TabIndex = 5;
            this._btnBrowse.Text = "...";
            this._btnBrowse.UseVisualStyleBackColor = true;
            this._btnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);

            // _lblStatus
            this._lblStatus.AutoSize = true;
            this._lblStatus.Location = new System.Drawing.Point(12, 165);
            this._lblStatus.Name = "_lblStatus";
            this._lblStatus.Size = new System.Drawing.Size(350, 13);
            this._lblStatus.TabIndex = 0;
            this._lblStatus.Text = "Click Install to register SolidWorks Release Pack Add-in.";
            
            // _btnInstall
            this._btnInstall.Location = new System.Drawing.Point(15, 190);
            this._btnInstall.Name = "_btnInstall";
            this._btnInstall.Size = new System.Drawing.Size(550, 40);
            this._btnInstall.TabIndex = 1;
            this._btnInstall.Text = "INSTALL RELEASE PACK";
            this._btnInstall.UseVisualStyleBackColor = true;
            this._btnInstall.Click += new System.EventHandler(this.BtnInstall_Click);
            
            // _progressBar
            this._progressBar.Location = new System.Drawing.Point(15, 240);
            this._progressBar.Name = "_progressBar";
            this._progressBar.Size = new System.Drawing.Size(550, 23);
            this._progressBar.TabIndex = 2;
            
            // MainForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 280);
            this.Controls.Add(this._lblPath);
            this.Controls.Add(this._txtSwPath);
            this.Controls.Add(this._btnBrowse);
            this.Controls.Add(this._bannerBox);
            this.Controls.Add(this._progressBar);
            this.Controls.Add(this._btnInstall);
            this.Controls.Add(this._lblStatus);
            this.Name = "MainForm";
            this.Text = "SolidWorks Release Pack Installer";
            ((System.ComponentModel.ISupportInitialize)(this._bannerBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
