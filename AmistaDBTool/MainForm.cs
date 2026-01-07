using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;

namespace AmistaDBTool
{
    public partial class MainForm : Form
    {
        private TextBox txtLog;
        private Button btnExecute;
        private Button btnCopyToClipboard;
        private ProgressBar progressBar;
        private CheckBox chkJob1;

        private SapConnector _sapConnector;
        private JobManager _jobManager;

        public MainForm()
        {
            InitializeComponent();
            InitializeLogic();
        }

        private void InitializeLogic()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                
                IConfiguration config = builder.Build();

                _sapConnector = new SapConnector(config, Log);
                _jobManager = new JobManager(_sapConnector, Log, UpdateProgress);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing configuration: {ex.Message}");
                Log($"Initialization Error: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(Log), message);
                return;
            }

            txtLog.AppendText($"{DateTime.Now:HH:mm:ss}: {message}{Environment.NewLine}");
        }

        private void UpdateProgress(int percent)
        {
             if (this.InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateProgress), percent);
                return;
            }
            if (percent < 0) percent = 0;
            if (percent > 100) percent = 100;
            progressBar.Value = percent;
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            if (chkJob1.Checked)
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Excel Files|*.xls;*.xlsx";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        btnExecute.Enabled = false;
                        try
                        {
                            _jobManager.ExecuteJob1(ofd.FileName);
                            MessageBox.Show("Job Completed Successfully.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error during execution: {ex.Message}");
                        }
                        finally
                        {
                            btnExecute.Enabled = true;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a job to execute.");
            }
        }

        private void BtnCopyToClipboard_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtLog.Text))
            {
                Clipboard.SetText(txtLog.Text);
                MessageBox.Show("Log copied to clipboard.");
            }
        }

        private void InitializeComponent()
        {
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.btnCopyToClipboard = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.chkJob1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            
            // 
            // txtLog
            // 
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Location = new System.Drawing.Point(12, 107);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(776, 331);
            this.txtLog.TabIndex = 0;
            
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(12, 59);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(120, 23);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.BtnExecute_Click);
            
            // 
            // btnCopyToClipboard
            // 
            this.btnCopyToClipboard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyToClipboard.Location = new System.Drawing.Point(668, 444); // Adjusted for visibility
            this.btnCopyToClipboard.Name = "btnCopyToClipboard";
            this.btnCopyToClipboard.Size = new System.Drawing.Size(120, 23);
            this.btnCopyToClipboard.TabIndex = 2;
            this.btnCopyToClipboard.Text = "Copy to Clipboard";
            this.btnCopyToClipboard.UseVisualStyleBackColor = true;
            this.btnCopyToClipboard.Click += new System.EventHandler(this.BtnCopyToClipboard_Click);

            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(12, 88);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(776, 13);
            this.progressBar.TabIndex = 3;

            // 
            // chkJob1
            // 
            this.chkJob1.AutoSize = true;
            this.chkJob1.Location = new System.Drawing.Point(12, 12);
            this.chkJob1.Name = "chkJob1";
            this.chkJob1.Size = new System.Drawing.Size(225, 19);
            this.chkJob1.TabIndex = 4;
            this.chkJob1.Text = "Job 1 - Delete BP Catalogs from XLS";
            this.chkJob1.UseVisualStyleBackColor = true;

            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 480);
            this.Controls.Add(this.chkJob1);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnCopyToClipboard);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtLog);
            this.Name = "MainForm";
            this.Text = "Amista DB Tool";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
