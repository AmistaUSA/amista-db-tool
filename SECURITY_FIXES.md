# Security Fixes Todo List

## Critical Priority

- [x] **Fix SQL injection in JobManager.cs:65** - Input validation + SQL escaping (SAP DI API doesn't support parameterized queries)
- [x] **Remove hardcoded credentials from appsettings.json** - DPAPI encryption implemented for passwords
- [x] **Add appsettings.json to .gitignore** - Created .gitignore and appsettings.template.json

## High Priority

- [x] **Implement secure logging in Program.cs** - Created SecureLogger with %LOCALAPPDATA% storage and log rotation
- [x] **Add file path validation in JobManager.cs:31** - Extension whitelist, 50MB size limit, path traversal prevention
- [x] **Add input validation for cardCode/itemCode** - Implemented regex validation and length limits

## Medium Priority

- [x] **Sanitize error messages in MainForm.cs** - Generic messages to users, details logged via SecureLogger
- [x] **Fix COM object resource leak in JobManager.cs** - Wrapped COM objects in try/finally block
- [x] **Update ExcelDataReader to latest version** - Updated from 3.6.0 to 3.7.0
