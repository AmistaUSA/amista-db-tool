using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using System.Security.Cryptography;
using Microsoft.Extensions.Configuration;

namespace AmistaDBTool
{
    public partial class MainForm : Form
    {
        // Existing controls
        private TextBox txtLog;
        private Button btnExecute;
        private Button btnCopyToClipboard;
        private ProgressBar progressBar;
        private CheckBox chkJob1;

        // Credential controls
        private GroupBox grpCredentials;
        private TextBox txtServer;
        private ComboBox cmbDbServerType;
        private TextBox txtDBUserName;
        private TextBox txtDBPassword;
        private TextBox txtCompanyDB;
        private TextBox txtUserName;
        private TextBox txtPassword;
        private TextBox txtLicenseServer;
        private TextBox txtSLDServer;
        private Button btnTestConnection;


        private SapConnector _sapConnector;
        private JobManager _jobManager;
        private IConfiguration _config;

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
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _config = builder.Build();

                _sapConnector = new SapConnector(_config, Log);
                _jobManager = new JobManager(_sapConnector, Log, UpdateProgress);

                LoadCredentials();
            }
            catch (Exception ex)
            {
                SecureLogger.LogError($"Initialization error: {ex.Message}");
                Log("Initialization error occurred. Check logs for details.");
                MessageBox.Show("Failed to initialize application. Please check the configuration file.",
                    "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadCredentials()
        {
            try
            {
                var sapConfig = _config.GetSection("SapConnection");
                txtServer.Text = sapConfig["Server"] ?? "";
                txtDBUserName.Text = sapConfig["DBUserName"] ?? "";
                txtDBPassword.Text = UnprotectPassword(sapConfig["DBPassword"] ?? "");
                txtCompanyDB.Text = sapConfig["CompanyDB"] ?? "";
                txtUserName.Text = sapConfig["UserName"] ?? "";
                txtPassword.Text = UnprotectPassword(sapConfig["Password"] ?? "");
                txtLicenseServer.Text = sapConfig["LicenseServer"] ?? "";
                txtSLDServer.Text = sapConfig["SLDServer"] ?? "";

                var dbServerType = sapConfig["DbServerType"] ?? "dst_HANADB";
                var index = cmbDbServerType.Items.IndexOf(dbServerType);
                cmbDbServerType.SelectedIndex = index >= 0 ? index : 0;
            }
            catch (Exception ex)
            {
                Log($"Error loading credentials: {ex.Message}");
            }
        }

        private void SaveCredentials()
        {
            try
            {
                var jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                var json = File.ReadAllText(jsonPath);

                using var doc = JsonDocument.Parse(json);
                var loggingSection = doc.RootElement.TryGetProperty("Logging", out var logging)
                    ? logging.Clone()
                    : default;

                var settings = new
                {
                    SapConnection = new
                    {
                        Server = txtServer.Text,
                        DbServerType = cmbDbServerType.SelectedItem?.ToString() ?? "dst_HANADB",
                        DBUserName = txtDBUserName.Text,
                        DBPassword = ProtectPassword(txtDBPassword.Text),
                        CompanyDB = txtCompanyDB.Text,
                        UserName = txtUserName.Text,
                        Password = ProtectPassword(txtPassword.Text),
                        LicenseServer = txtLicenseServer.Text,
                        SLDServer = txtSLDServer.Text
                    },
                    Logging = loggingSection.ValueKind != JsonValueKind.Undefined
                        ? JsonSerializer.Deserialize<object>(loggingSection.GetRawText())
                        : new { LogLevel = new { Default = "Information", Microsoft_Hosting_Lifetime = "Information" } }
                };

                var options = new JsonSerializerOptions { WriteIndented = true };
                var newJson = JsonSerializer.Serialize(settings, options);
                File.WriteAllText(jsonPath, newJson);

                // Reload configuration
                var builder = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                _config = builder.Build();

                Log("Credentials saved successfully.");
            }
            catch (Exception ex)
            {
                SecureLogger.LogError($"Error saving credentials: {ex.Message}");
                Log("Error saving credentials. Check logs for details.");
                MessageBox.Show("Failed to save credentials. Please check file permissions.",
                    "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string ProtectPassword(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return plainText;
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(plainText);
                byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            catch
            {
                return plainText;
            }
        }

        private string UnprotectPassword(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText)) return encryptedText;
            try
            {
                byte[] encrypted = Convert.FromBase64String(encryptedText);
                byte[] decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch
            {
                // Not encrypted yet (plain text) - return as is
                return encryptedText;
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
                            SecureLogger.LogError($"Job execution error: {ex.Message}");
                            MessageBox.Show("An error occurred during job execution. Check the log for details.",
                                "Execution Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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



        private void BtnTestConnection_Click(object sender, EventArgs e)
        {
            btnTestConnection.Enabled = false;
            Log("Testing connection...");

            SapConnector testConnector = null;
            try
            {
                // Save credentials first to ensure config is updated
                SaveCredentials();

                // Create a separate test connector (don't affect the main one)
                testConnector = new SapConnector(_config, Log);

                // Attempt connection
                var company = testConnector.GetCompany();

                if (company != null && company.Connected)
                {
                    MessageBox.Show("Connection successful!", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Log("Connection test successful.");
                }
                else
                {
                    MessageBox.Show("Connection failed. Check the log for details.", "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                SecureLogger.LogError($"Connection test failed: {ex.Message}");
                Log($"Connection test failed: {ex.Message}");
                MessageBox.Show("Connection failed. Please verify your credentials and server settings.",
                    "Test Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Always disconnect and clean up test connector
                testConnector?.Disconnect();
                btnTestConnection.Enabled = true;

                // Recreate main connector and job manager with updated config
                _sapConnector = new SapConnector(_config, Log);
                _jobManager = new JobManager(_sapConnector, Log, UpdateProgress);
            }
        }

        private void InitializeComponent()
        {
            this.grpCredentials = new System.Windows.Forms.GroupBox();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.cmbDbServerType = new System.Windows.Forms.ComboBox();
            this.txtDBUserName = new System.Windows.Forms.TextBox();
            this.txtDBPassword = new System.Windows.Forms.TextBox();
            this.txtCompanyDB = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtLicenseServer = new System.Windows.Forms.TextBox();
            this.txtSLDServer = new System.Windows.Forms.TextBox();
            this.btnTestConnection = new System.Windows.Forms.Button();

            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.btnCopyToClipboard = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.chkJob1 = new System.Windows.Forms.CheckBox();
            this.grpCredentials.SuspendLayout();
            this.SuspendLayout();

            //
            // grpCredentials
            //
            this.grpCredentials.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpCredentials.Location = new System.Drawing.Point(12, 12);
            this.grpCredentials.Name = "grpCredentials";
            this.grpCredentials.Size = new System.Drawing.Size(776, 200);
            this.grpCredentials.TabIndex = 0;
            this.grpCredentials.TabStop = false;
            this.grpCredentials.Text = "SAP Connection Settings";

            // Layout constants
            int col1X = 15;
            int col2X = 270;
            int col3X = 525;
            int labelWidth = 100;
            int textWidth = 145;
            int row1Y = 25;
            int row2Y = 55;
            int row3Y = 85;
            int row4Y = 115;
            int row5Y = 145;

            // Row 1: Server, DB Server Type, License Server
            var lblServer = new Label { Text = "Server:", Location = new System.Drawing.Point(col1X, row1Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtServer.Location = new System.Drawing.Point(col1X + labelWidth, row1Y);
            this.txtServer.Size = new System.Drawing.Size(textWidth, 23);

            var lblDbServerType = new Label { Text = "DB Type:", Location = new System.Drawing.Point(col2X, row1Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.cmbDbServerType.Location = new System.Drawing.Point(col2X + labelWidth, row1Y);
            this.cmbDbServerType.Size = new System.Drawing.Size(textWidth, 23);
            this.cmbDbServerType.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbDbServerType.Items.AddRange(new object[] {
                "dst_MSSQL", "dst_DB_2", "dst_SYBASE", "dst_MSSQL2005", "dst_MAXDB",
                "dst_MSSQL2008", "dst_MSSQL2012", "dst_MSSQL2014", "dst_HANADB",
                "dst_MSSQL2016", "dst_MSSQL2017", "dst_MSSQL2019"
            });

            var lblLicenseServer = new Label { Text = "License Server:", Location = new System.Drawing.Point(col3X, row1Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtLicenseServer.Location = new System.Drawing.Point(col3X + labelWidth, row1Y);
            this.txtLicenseServer.Size = new System.Drawing.Size(textWidth, 23);

            // Row 2: Company DB, DB User Name, SLD Server
            var lblCompanyDB = new Label { Text = "Company DB:", Location = new System.Drawing.Point(col1X, row2Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtCompanyDB.Location = new System.Drawing.Point(col1X + labelWidth, row2Y);
            this.txtCompanyDB.Size = new System.Drawing.Size(textWidth, 23);

            var lblDBUserName = new Label { Text = "DB User:", Location = new System.Drawing.Point(col2X, row2Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtDBUserName.Location = new System.Drawing.Point(col2X + labelWidth, row2Y);
            this.txtDBUserName.Size = new System.Drawing.Size(textWidth, 23);

            var lblSLDServer = new Label { Text = "SLD Server:", Location = new System.Drawing.Point(col3X, row2Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtSLDServer.Location = new System.Drawing.Point(col3X + labelWidth, row2Y);
            this.txtSLDServer.Size = new System.Drawing.Size(textWidth, 23);

            // Row 3: SAP User Name, DB Password
            var lblUserName = new Label { Text = "SAP User:", Location = new System.Drawing.Point(col1X, row3Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtUserName.Location = new System.Drawing.Point(col1X + labelWidth, row3Y);
            this.txtUserName.Size = new System.Drawing.Size(textWidth, 23);

            var lblDBPassword = new Label { Text = "DB Password:", Location = new System.Drawing.Point(col2X, row3Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtDBPassword.Location = new System.Drawing.Point(col2X + labelWidth, row3Y);
            this.txtDBPassword.Size = new System.Drawing.Size(textWidth, 23);
            this.txtDBPassword.UseSystemPasswordChar = true;

            // Row 4: SAP Password
            var lblPassword = new Label { Text = "SAP Password:", Location = new System.Drawing.Point(col1X, row4Y), Size = new System.Drawing.Size(labelWidth, 20) };
            this.txtPassword.Location = new System.Drawing.Point(col1X + labelWidth, row4Y);
            this.txtPassword.Size = new System.Drawing.Size(textWidth, 23);
            this.txtPassword.UseSystemPasswordChar = true;

            // Row 5: Buttons
            this.btnTestConnection.Location = new System.Drawing.Point(col1X, row5Y);
            this.btnTestConnection.Size = new System.Drawing.Size(180, 30);
            this.btnTestConnection.Text = "Save and Test Credentials";
            this.btnTestConnection.UseVisualStyleBackColor = true;
            this.btnTestConnection.Click += new System.EventHandler(this.BtnTestConnection_Click);



            // Add controls to GroupBox
            this.grpCredentials.Controls.Add(lblServer);
            this.grpCredentials.Controls.Add(this.txtServer);
            this.grpCredentials.Controls.Add(lblDbServerType);
            this.grpCredentials.Controls.Add(this.cmbDbServerType);
            this.grpCredentials.Controls.Add(lblLicenseServer);
            this.grpCredentials.Controls.Add(this.txtLicenseServer);
            this.grpCredentials.Controls.Add(lblCompanyDB);
            this.grpCredentials.Controls.Add(this.txtCompanyDB);
            this.grpCredentials.Controls.Add(lblDBUserName);
            this.grpCredentials.Controls.Add(this.txtDBUserName);
            this.grpCredentials.Controls.Add(lblSLDServer);
            this.grpCredentials.Controls.Add(this.txtSLDServer);
            this.grpCredentials.Controls.Add(lblUserName);
            this.grpCredentials.Controls.Add(this.txtUserName);
            this.grpCredentials.Controls.Add(lblDBPassword);
            this.grpCredentials.Controls.Add(this.txtDBPassword);
            this.grpCredentials.Controls.Add(lblPassword);
            this.grpCredentials.Controls.Add(this.txtPassword);
            this.grpCredentials.Controls.Add(this.btnTestConnection);


            //
            // chkJob1 (moved down)
            //
            this.chkJob1.AutoSize = true;
            this.chkJob1.Location = new System.Drawing.Point(12, 222);
            this.chkJob1.Name = "chkJob1";
            this.chkJob1.Size = new System.Drawing.Size(225, 19);
            this.chkJob1.TabIndex = 1;
            this.chkJob1.Text = "Job 1 - Delete BP Catalogs from XLS";
            this.chkJob1.UseVisualStyleBackColor = true;

            //
            // btnExecute (moved down)
            //
            this.btnExecute.Location = new System.Drawing.Point(12, 249);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(120, 23);
            this.btnExecute.TabIndex = 2;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.BtnExecute_Click);

            //
            // progressBar (moved down)
            //
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(12, 278);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(776, 13);
            this.progressBar.TabIndex = 3;

            //
            // txtLog (moved down)
            //
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Location = new System.Drawing.Point(12, 297);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(776, 331);
            this.txtLog.TabIndex = 4;

            //
            // btnCopyToClipboard
            //
            this.btnCopyToClipboard.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyToClipboard.Location = new System.Drawing.Point(668, 634);
            this.btnCopyToClipboard.Name = "btnCopyToClipboard";
            this.btnCopyToClipboard.Size = new System.Drawing.Size(120, 23);
            this.btnCopyToClipboard.TabIndex = 5;
            this.btnCopyToClipboard.Text = "Copy to Clipboard";
            this.btnCopyToClipboard.UseVisualStyleBackColor = true;
            this.btnCopyToClipboard.Click += new System.EventHandler(this.BtnCopyToClipboard_Click);

            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 670);
            this.Controls.Add(this.grpCredentials);
            this.Controls.Add(this.chkJob1);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnCopyToClipboard);
            this.Name = "MainForm";
            this.Text = "Amista DB Tool";
            this.grpCredentials.ResumeLayout(false);
            this.grpCredentials.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
