using System;
using System.Text;
using System.Security.Cryptography;
using SAPbobsCOM;
using Microsoft.Extensions.Configuration;

namespace AmistaDBTool
{
    public class SapConnector
    {
        private Company _company;
        private readonly IConfiguration _config;
        private readonly Action<string> _logger;

        public SapConnector(IConfiguration config, Action<string> logger)
        {
            _config = config;
            _logger = logger;
        }

        public Company GetCompany()
        {
            if (_company != null && _company.Connected)
                return _company;

            _company = new Company();
            var sapConfig = _config.GetSection("SapConnection");

            _company.Server = sapConfig["Server"];
            _company.DbServerType = Enum.Parse<BoDataServerTypes>(sapConfig["DbServerType"] ?? "dst_HANADB");
            _company.DbUserName = sapConfig["DBUserName"];
            _company.DbPassword = UnprotectPassword(sapConfig["DBPassword"]);
            _company.CompanyDB = sapConfig["CompanyDB"];
            _company.UserName = sapConfig["UserName"];
            _company.Password = UnprotectPassword(sapConfig["Password"]);
            _company.LicenseServer = sapConfig["LicenseServer"];
            _company.SLDServer = sapConfig["SLDServer"];
            _company.UseTrusted = false; 

            try
            {
                SecureLogger.LogDebug("Attempting SAP connection...");

                int ret = _company.Connect();
                SecureLogger.LogDebug($"Connect() returned: {ret}");

                if (ret != 0)
                {
                    _company.GetLastError(out int errCode, out string errMsg);
                    string error = $"SAP connection failed: {errCode} - {errMsg}";
                    SecureLogger.LogError(error);
                    _logger(error);
                    throw new Exception($"SAP Connection Failed: {errMsg}");
                }

                _logger("Connected to SAP Company.");
                SecureLogger.LogDebug("SAP connection successful.");
                return _company;
            }
            catch (Exception ex)
            {
                SecureLogger.LogError($"Exception in Connect: {ex.Message}");
                throw;
            }
        }

        public void Disconnect()
        {
            if (_company != null && _company.Connected)
            {
                _company.Disconnect();
                _logger("Disconnected from SAP.");
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
    }
}
