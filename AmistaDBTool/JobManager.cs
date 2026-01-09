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

        // Pattern for valid SAP codes: alphanumeric, hyphens, underscores only
        // Removed dots and spaces to reduce SQL injection attack surface
        private static readonly Regex ValidCodePattern = new Regex(@"^[a-zA-Z0-9\-_]+$", RegexOptions.Compiled);
        private const int MaxCodeLength = 50; // SAP B1 codes are typically max 15-20 chars

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
        /// Escapes and sanitizes value for SQL (defense in depth).
        /// Handles multiple SQL injection vectors including quotes, comments, and control characters.
        /// </summary>
        private string EscapeSql(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            var sb = new StringBuilder(value.Length);
            foreach (char c in value)
            {
                switch (c)
                {
                    case '\'': sb.Append("''"); break;      // Escape single quotes
                    case '\\': sb.Append("\\\\"); break;    // Escape backslashes
                    case '\0': break;                        // Remove null bytes
                    case '\r': break;                        // Remove carriage returns
                    case '\n': break;                        // Remove newlines
                    case '\x1a': break;                      // Remove substitute char (EOF in some contexts)
                    default:
                        // Only allow printable ASCII characters
                        if (c >= 32 && c <= 126)
                            sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Validates and sanitizes a SAP code for safe SQL usage.
        /// Returns null if the code is invalid.
        /// </summary>
        private string ValidateAndSanitizeCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return null;

            // Trim and normalize
            string trimmed = code.Trim();

            // Check length
            if (trimmed.Length == 0 || trimmed.Length > MaxCodeLength) return null;

            // Validate against whitelist pattern
            if (!ValidCodePattern.IsMatch(trimmed)) return null;

            // Apply SQL escaping as defense in depth
            return EscapeSql(trimmed);
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
                                    string rawCardCode = row[0]?.ToString();
                                    string rawItemCode = row[1]?.ToString();

                                    // Validate and sanitize inputs - returns null if invalid
                                    string safeCardCode = ValidateAndSanitizeCode(rawCardCode);
                                    string safeItemCode = ValidateAndSanitizeCode(rawItemCode);

                                    if (safeCardCode == null || safeItemCode == null)
                                    {
                                        _logger($"Row {processed + 1}: Skipped - invalid or missing CardCode/ItemCode.");
                                        SecureLogger.LogWarning($"Row {processed + 1}: Input validation failed for potential SQL injection attempt.");
                                        continue;
                                    }

                                    // Store original trimmed values for SAP API calls (after validation)
                                    string cardCode = rawCardCode.Trim();
                                    string itemCode = rawItemCode.Trim();

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
