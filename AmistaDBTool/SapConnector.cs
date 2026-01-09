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

                    // Log full details securely (for admin troubleshooting)
                    SecureLogger.LogError($"SAP connection failed: Code={errCode}, Message={errMsg}");

                    // Show sanitized message to user (no internal details)
                    string userMessage = GetSanitizedErrorMessage(errCode);
                    _logger($"Connection failed: {userMessage}");

                    throw new SapConnectionException(userMessage, errCode);
                }

                _logger("Connected to SAP Company.");
                SecureLogger.LogDebug("SAP connection successful.");
                return _company;
            }
            catch (SapConnectionException)
            {
                throw; // Re-throw our sanitized exception
            }
            catch (Exception ex)
            {
                SecureLogger.LogError($"Exception in Connect: {ex.Message}");
                _logger("Connection failed due to an unexpected error. Check logs for details.");
                throw new SapConnectionException("An unexpected error occurred while connecting to SAP.", ex);
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

        /// <summary>
        /// Returns a user-friendly error message based on SAP error code.
        /// Does not expose internal system details.
        /// </summary>
        private static string GetSanitizedErrorMessage(int errorCode)
        {
            // Map common SAP error codes to user-friendly messages
            // Full details are logged securely, not shown to user
            return errorCode switch
            {
                -1 => "Unable to connect to the server. Please verify the server address.",
                -2 => "Invalid database type. Please check your configuration.",
                -10 => "License server connection failed. Please verify the license server address.",
                -100 => "Invalid credentials. Please verify your username and password.",
                -101 => "User account is locked or disabled.",
                -102 => "Database access denied. Please verify database credentials.",
                -103 => "Company database not found. Please verify the company database name.",
                -1000 => "Server is not available. Please check network connectivity.",
                _ => $"Connection failed (Error code: {errorCode}). Please check logs for details."
            };
        }

        private const string EncryptionPrefix = "ENC:";

        private string UnprotectPassword(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText)) return encryptedText;

            // Check for encryption prefix
            if (encryptedText.StartsWith(EncryptionPrefix))
            {
                try
                {
                    string base64 = encryptedText.Substring(EncryptionPrefix.Length);
                    byte[] encrypted = Convert.FromBase64String(base64);
                    byte[] decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                    return Encoding.UTF8.GetString(decrypted);
                }
                catch (Exception ex)
                {
                    SecureLogger.LogError($"Password decryption failed: {ex.Message}");
                    throw new CryptographicException(
                        "Failed to decrypt password. The password may have been encrypted by a different user account.", ex);
                }
            }

            // Legacy plaintext password detected - log warning and return as-is for migration
            // The password will be encrypted when credentials are next saved via MainForm
            SecureLogger.LogWarning("Plaintext password detected in configuration. Save credentials to encrypt.");
            return encryptedText;
        }
    }
}
