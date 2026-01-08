using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Security;
using System.Text.RegularExpressions;
using ExcelDataReader;
using SAPbobsCOM;
using System.Text;

namespace AmistaDBTool
{
    public class JobManager
    {
        private readonly Action<string> _logger;
        private readonly Action<int> _progress;
        private readonly SapConnector _sapConnector;

        // Pattern for valid SAP codes: alphanumeric, hyphens, underscores, dots, spaces
        private static readonly Regex ValidCodePattern = new Regex(@"^[a-zA-Z0-9\-_\.\s]+$", RegexOptions.Compiled);
        private const int MaxCodeLength = 100;

        // File validation constants
        private static readonly string[] AllowedExtensions = { ".xls", ".xlsx" };
        private const long MaxFileSizeBytes = 50 * 1024 * 1024; // 50 MB

        public JobManager(SapConnector sapConnector, Action<string> logger, Action<int> progress)
        {
            _sapConnector = sapConnector;
            _logger = logger;
            _progress = progress;
        }

        /// <summary>
        /// Validates that a code contains only safe characters and is within length limits.
        /// </summary>
        private bool IsValidCode(string code)
        {
            if (string.IsNullOrEmpty(code)) return false;
            if (code.Length > MaxCodeLength) return false;
            return ValidCodePattern.IsMatch(code);
        }

        /// <summary>
        /// Escapes single quotes for SQL (defense in depth).
        /// </summary>
        private string EscapeSql(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Replace("'", "''");
        }

        /// <summary>
        /// Validates file path for security: extension, size, path traversal.
        /// </summary>
        private void ValidateFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be empty.");

            // Check for path traversal attempts
            if (filePath.Contains("..") || filePath.Contains("..\\") || filePath.Contains("../"))
                throw new SecurityException("Invalid file path: path traversal detected.");

            // Get full path to resolve any relative paths
            string fullPath = Path.GetFullPath(filePath);

            // Check file exists
            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"File not found: {fullPath}");

            // Validate extension
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (!AllowedExtensions.Contains(extension))
                throw new ArgumentException($"Invalid file type. Allowed: {string.Join(", ", AllowedExtensions)}");

            // Check file size
            var fileInfo = new FileInfo(fullPath);
            if (fileInfo.Length > MaxFileSizeBytes)
                throw new ArgumentException($"File too large. Maximum size: {MaxFileSizeBytes / (1024 * 1024)} MB");

            // Check file size is not zero
            if (fileInfo.Length == 0)
                throw new ArgumentException("File is empty.");
        }

        public void ExecuteJob1(string filePath)
        {
            // Validate file path before processing
            ValidateFilePath(filePath);

            _logger($"Starting Job 1 with file: {filePath}");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        var table = result.Tables[0];
                        int totalRows = table.Rows.Count;
                        _logger($"Found {totalRows} rows in Excel file.");

                        var company = _sapConnector.GetCompany();
                        Recordset rs = null;
                        dynamic bpCatalog = null;

                        try
                        {
                            rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            bpCatalog = company.GetBusinessObject(BoObjectTypes.oAlternateCatNum);

                            int processed = 0;
                            foreach (DataRow row in table.Rows)
                            {
                                try
                                {
                                    string cardCode = row[0]?.ToString()?.Trim();
                                    string itemCode = row[1]?.ToString()?.Trim();

                                    if (string.IsNullOrEmpty(cardCode) || string.IsNullOrEmpty(itemCode))
                                    {
                                        _logger($"Row {processed + 1}: Skipped due to missing data.");
                                        continue;
                                    }

                                    // Validate input to prevent SQL injection
                                    if (!IsValidCode(cardCode) || !IsValidCode(itemCode))
                                    {
                                        _logger($"Row {processed + 1}: Skipped due to invalid characters in CardCode or ItemCode.");
                                        continue;
                                    }

                                    // Escape SQL as defense in depth
                                    string safeCardCode = EscapeSql(cardCode);
                                    string safeItemCode = EscapeSql(itemCode);

                                    string query = $"SELECT \"Substitute\" FROM \"OSCN\" WHERE \"CardCode\" = '{safeCardCode}' AND \"ItemCode\" = '{safeItemCode}'";
                                    rs.DoQuery(query);

                                    if (rs.RecordCount > 0)
                                    {
                                        string Substitute = rs.Fields.Item("Substitute").Value.ToString();
                                        if (bpCatalog.GetByKey(itemCode, cardCode, Substitute))
                                        {
                                            int ret = bpCatalog.Remove();
                                            if (ret != 0)
                                            {
                                                company.GetLastError(out int err, out string msg);
                                                _logger($"Row {processed + 1}: Error deleting BP Catalog for {cardCode}-{itemCode}. Error: {msg}");
                                            }
                                            else
                                            {
                                                _logger($"Row {processed + 1}: Successfully deleted BP Catalog for {cardCode}-{itemCode}.");
                                            }
                                        }
                                        else
                                        {
                                            _logger($"Row {processed + 1}: Could not retrieve object for {cardCode}-{itemCode}");
                                        }
                                    }
                                    else
                                    {
                                        _logger($"Row {processed + 1}: No BP Catalog found for {cardCode}-{itemCode}.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger($"Row {processed + 1}: Error processing row. {ex.Message}");
                                }
                                finally
                                {
                                    processed++;
                                    int pct = (int)((double)processed / totalRows * 100);
                                    _progress(pct);
                                }

                                // Keep UI responsive
                                System.Windows.Forms.Application.DoEvents();
                            }
                        }
                        finally
                        {
                            // Always release COM objects to prevent memory leaks
                            if (rs != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                            if (bpCatalog != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(bpCatalog);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"Critical Error in Job 1: {ex.Message}");
                throw;
            }
            finally
            {
                _sapConnector.Disconnect();
            }
            
            _logger("Job 1 Completed.");
        }
    }
}
