using System;
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
            _company.DbPassword = sapConfig["DBPassword"];
            _company.CompanyDB = sapConfig["CompanyDB"];
            _company.UserName = sapConfig["UserName"];
            _company.Password = sapConfig["Password"];
            _company.LicenseServer = sapConfig["LicenseServer"];
            _company.SLDServer = sapConfig["SLDServer"];
            _company.UseTrusted = false; 

            try
            {
                File.AppendAllText("debug.log", $"[{DateTime.Now}] Attempting to connect...\n");
                
                File.AppendAllText("debug.log", $"[{DateTime.Now}] Company object created? {_company != null}\n");
                
                int ret = _company.Connect();
                File.AppendAllText("debug.log", $"[{DateTime.Now}] Connect() returned: {ret}\n");

                if (ret != 0)
                {
                    _company.GetLastError(out int errCode, out string errMsg);
                    string error = $"Error connecting to SAP: {errCode} - {errMsg}";
                    File.AppendAllText("debug.log", $"[{DateTime.Now}] {error}\n");
                    _logger(error);
                    throw new Exception($"SAP Connection Failed: {errMsg}");
                }

                File.AppendAllText("debug.log", $"[{DateTime.Now}] Connection check passed. Logging success...\n");
                
                _logger($"Connected to SAP Company.");
                File.AppendAllText("debug.log", $"[{DateTime.Now}] Connected successfully.\n");
                return _company;
            }
            catch (Exception ex)
            {
                File.AppendAllText("debug.log", $"[{DateTime.Now}] EXCEPTION in Connect: {ex.Message}\n{ex.StackTrace}\n");
                throw;
            }
            return _company;
        }

        public void Disconnect()
        {
            if (_company != null && _company.Connected)
            {
                _company.Disconnect();
                _logger("Disconnected from SAP.");
            }
        }
    }
}
