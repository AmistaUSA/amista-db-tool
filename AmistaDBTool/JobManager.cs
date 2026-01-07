using System;
using System.IO;
using System.Data;
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

        public JobManager(SapConnector sapConnector, Action<string> logger, Action<int> progress)
        {
            _sapConnector = sapConnector;
            _logger = logger;
            _progress = progress;
        }

        public void ExecuteJob1(string filePath)
        {
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
                        Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        dynamic bpCatalog = company.GetBusinessObject(BoObjectTypes.oAlternateCatNum);

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

                                string query = $"SELECT \"Substitute\" FROM \"OSCN\" WHERE \"CardCode\" = '{cardCode}' AND \"ItemCode\" = '{itemCode}'";
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
                        
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(bpCatalog);
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
