# Customizing the Reporter Service for Your Environment

This guide provides step-by-step instructions for implementing the Reporter Service in your specific environment. Follow these instructions carefully to ensure proper deployment and configuration.

## Prerequisites

### System Requirements
- **Operating System**: Windows Server 2008 R2 or later
- **Framework**: .NET Framework 4.0 or later
- **Database**: SQL Server 2008 or later
- **Memory**: Minimum 2GB RAM (4GB recommended)
- **Disk Space**: 1GB free space for logs and temporary files

### Required Software Components
1. **Crystal Reports Runtime** (version compatible with your reports)
2. **Spire.PDF Library** (included in packages)
3. **log4net Framework** (included in packages)
4. **SQL Server Client Components**

## Step 1: Database Setup

### 1.1 Create the Database
Execute the provided `ReportingService_Database_Schema.sql` script:

```sql
-- Run this script as a database administrator
sqlcmd -S YourServerName -i ReportingService_Database_Schema.sql
```

### 1.2 Configure Database Permissions
Create a dedicated service account for the Reporter Service:

```sql
-- Create login
CREATE LOGIN [ReporterService] WITH PASSWORD = 'YourSecurePassword123!'

-- Create user in ReportingService database
USE [ReportingService]
CREATE USER [ReporterService] FOR LOGIN [ReporterService]

-- Grant necessary permissions
ALTER ROLE [ReportWriter] ADD MEMBER [ReporterService]
ALTER ROLE [ReportReader] ADD MEMBER [ReporterService]
```

### 1.3 Update Connection Strings
Modify the following files with your database connection details:

**In `ReporterService.vb`** (lines 169-171):
```vb
objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "UserName", "ReporterService")
objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Password", "YourSecurePassword123!")
objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBName", "ReportingService")
objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBServer", "YourServerName\YourInstanceName")
```

## Step 2: Web Service Configuration

### 2.1 Update Web Service References
Replace the web service references in your project:

**In `ReporterService.vb`** (lines 310-311):
```vb
Dim ReporterWS As New com.yourcompany.reporting.Service
Dim BasicWS As New com.yourcompany.basic.Service
```

### 2.2 Configure Web Service Endpoints
Update the `app.config` file:

```xml
<applicationSettings>
    <ReporterService.My.MySettings>
        <setting name="ReporterService_com_yourcompany_reporting_Service"
            serializeAs="String">
            <value>https://yourcompany.com/reporting/service.asmx</value>
        </setting>
        <setting name="ReporterService_com_yourcompany_basic_Service"
            serializeAs="String">
            <value>https://yourcompany.com/basic/service.asmx</value>
        </setting>
    </ReporterService.My.MySettings>
</applicationSettings>
```

### 2.3 Update Service URLs
**In `ReporterService.vb`** (line 567):
```vb
ResultString = http.geturl("http://yourcompany.com/reporting/service.asmx/ExecReport?sXML=" & wp, reportServiceUrl, 80, "", "")
```

## Step 3: Network Configuration

### 3.1 Update Network Settings
**In `app.config`**:
```xml
<appSettings>
    <add key="reportServiceUrl" value="192.168.1.100"/>
</appSettings>
```

### 3.2 Configure Logging Network
**In `app.config`**:
```xml
<appender name="RemoteSyslogAppender" type="log4net.Appender.RemoteSyslogAppender">
    <remoteAddress value="192.168.1.200" />
</appender>
```

## Step 4: File System Configuration

### 4.1 Create Required Directories
Create the following directories on your server:

```batch
mkdir C:\Logs
mkdir C:\Reports\Output
mkdir C:\Reports\Temp
mkdir C:\gsbatchprint64
```

### 4.2 Set Directory Permissions
Grant the service account full control to these directories:

```batch
icacls "C:\Logs" /grant "YourDomain\ReporterServiceAccount":F
icacls "C:\Reports\Output" /grant "YourDomain\ReporterServiceAccount":F
icacls "C:\Reports\Temp" /grant "YourDomain\ReporterServiceAccount":F
```

### 4.3 Update File Paths
**In `ReporterService.vb`** (line 221):
```vb
path = "C:\Logs\"
```

## Step 5: Email Configuration

### 5.1 SMTP Server Setup
Configure your SMTP server settings in the report destinations:

```sql
INSERT INTO [dbo].[REPORT_DESTINATIONS] 
([ROW_ID], [NAME], [TYPE_CD], [HOSTNAME], [USERNAME], [PASSWORD]) 
VALUES 
('DEST_EMAIL_001', 'Company Email Server', 'EMAIL', 'smtp.yourcompany.com', 'reports@yourcompany.com', 'YourEmailPassword')
```

### 5.2 Update Email Addresses
**In `ReporterService.vb`**, replace all instances of:
```vb
"admin@yourcompany.com"
```
with your actual admin email address.

## Step 6: Printer Configuration

### 6.1 Install Print Drivers
Ensure the following are installed on your server:
- PDF printing drivers
- Network printer drivers
- Ghostscript (if using gsbatchprint64)

### 6.2 Configure Printer Paths
**In `ReporterService.vb`** (line 1145):
```vb
lvbProcessInfo.FileName = "C:\gsbatchprint64\gsbatchprintc.exe "
```

### 6.3 Update Printer Queue Names
Configure printer names in your `LIST_OF_VALUES` table:

```sql
INSERT INTO [dbo].[LIST_OF_VALUES] 
([ROW_ID], [TYPE], [VAL], [DISP_VAL]) 
VALUES 
('PRINTER_001', 'PRINTER_QUEUES', 'YourNetworkPrinter', 'Main Office Printer')
```

## Step 7: Service Installation

### 7.1 Build the Service
1. Open the solution in Visual Studio
2. Update all references to your company's web services
3. Build the solution in Release mode
4. Verify all dependencies are included

### 7.2 Install the Service
```batch
# Copy files to installation directory
xcopy "bin\Release\*" "C:\ReporterService\" /E /Y

# Install the service
sc create "ReporterService" binPath= "C:\ReporterService\ReporterService.exe" start= auto
sc description "ReporterService" "Automated Report Generation and Distribution Service"
```

### 7.3 Configure Service Account
```batch
# Set service to run under dedicated account
sc config "ReporterService" obj= "YourDomain\ReporterServiceAccount" password= "ServiceAccountPassword"
```

## Step 8: Initial Configuration

### 8.1 Set Registry Values
The service will create default registry values on first run, but you can pre-configure them:

```batch
reg add "HKLM\Software\ReporterService" /v "MyInterval" /t REG_SZ /d "5" /f
reg add "HKLM\Software\ReporterService" /v "UserName" /t REG_SZ /d "ReporterService" /f
reg add "HKLM\Software\ReporterService" /v "Password" /t REG_SZ /d "YourSecurePassword123!" /f
reg add "HKLM\Software\ReporterService" /v "DBName" /t REG_SZ /d "ReportingService" /f
reg add "HKLM\Software\ReporterService" /v "DBServer" /t REG_SZ /d "YourServerName\YourInstanceName" /f
reg add "HKLM\Software\ReporterService" /v "Debug" /t REG_SZ /d "Y" /f
reg add "HKLM\Software\ReporterService" /v "Logging" /t REG_SZ /d "Y" /f
```

### 8.2 Test Database Connectivity
Create a test script to verify database connectivity:

```sql
-- Test connection
SELECT GETDATE() AS CurrentTime, @@SERVERNAME AS ServerName
```

## Step 9: Sample Data Setup

### 9.1 Create Sample Reports
```sql
-- Insert sample report
INSERT INTO [dbo].[REPORTS] 
([ROW_ID], [NAME], [FILENAME], [DESCRIPTION], [SQL_REP]) 
VALUES 
('RPT_001', 'User Activity Report', 'UserActivity.rpt', 'Daily user activity summary', 'SELECT * FROM CONTACTS WHERE CREATED >= GETDATE() - 1')

-- Insert sample report entity
INSERT INTO [dbo].[REPORT_ENTITIES] 
([ROW_ID], [REPORT_ID], [REP_DEST_ID], [FORMAT]) 
VALUES 
('ENT_001', 'RPT_001', 'DEST_001', 'PDF')
```

### 9.2 Create Test Schedule
```sql
-- Insert test schedule
INSERT INTO [dbo].[REPORT_EXECUTION_SCHEDULE] 
([ROW_ID], [TRAN_ID], [ENT_ID], [STATUS]) 
VALUES 
('SCHED_001', 'TEST_001', 'ENT_001', 'PENDING')
```

## Step 10: Testing and Validation

### 10.1 Start the Service
```batch
sc start "ReporterService"
```

### 10.2 Monitor Logs
Check the following locations for logs:
- Windows Event Log: `Application` log, source `ReporterService`
- File logs: `C:\Logs\ReporterService.log`
- Remote syslog: Your configured syslog server

### 10.3 Verify Processing
```sql
-- Check processing status
SELECT * FROM [dbo].[REPORT_EXECUTION_SCHEDULE] WHERE STATUS = 'PENDING'
SELECT * FROM [dbo].[PROCESSING_QUEUE] WHERE STATUS = 'PROCESSING'
```

## Step 11: Production Configuration

### 11.1 Disable Debug Mode
```batch
reg add "HKLM\Software\ReporterService" /v "Debug" /t REG_SZ /d "N" /f
```

### 11.2 Optimize Performance
- Adjust `MyInterval` based on your processing needs
- Configure appropriate batch sizes
- Set up monitoring and alerting

### 11.3 Security Hardening
- Use strong passwords for all accounts
- Enable Windows Firewall rules for required ports
- Configure SSL/TLS for web service communications
- Regular security updates

## Troubleshooting Common Issues

### Issue 1: Service Won't Start
**Symptoms**: Service fails to start with error messages
**Solutions**:
1. Check Windows Event Log for detailed error messages
2. Verify database connectivity
3. Ensure all required directories exist
4. Check service account permissions

### Issue 2: Reports Not Processing
**Symptoms**: Queue has pending records but no processing occurs
**Solutions**:
1. Verify web service connectivity
2. Check database locks (look for records with `LOCKED_BY` set)
3. Review web service endpoint configuration
4. Check network connectivity to report service

### Issue 3: Delivery Failures
**Symptoms**: Reports generate but fail to deliver
**Solutions**:
1. Verify SMTP server configuration
2. Check network printer connectivity
3. Validate file system permissions
4. Review FTP server settings

### Issue 4: Performance Issues
**Symptoms**: Slow processing or high resource usage
**Solutions**:
1. Adjust batch processing size
2. Optimize database queries
3. Review system resources
4. Check for database locks

## Maintenance Procedures

### Daily Tasks
- Monitor service status
- Review error logs
- Check queue processing
- Verify delivery success rates

### Weekly Tasks
- Review performance metrics
- Clean up old log files
- Update report schedules
- Backup configuration

### Monthly Tasks
- Security updates
- Performance optimization
- Capacity planning
- Disaster recovery testing

## Support and Escalation

### Log Collection
When reporting issues, collect the following:
1. Windows Event Log entries
2. Service log files
3. Database error logs
4. Network connectivity tests

### Contact Information
For technical support:
- Internal IT Support: [Your IT Support Contact]
- Database Administrator: [Your DBA Contact]
- Network Administrator: [Your Network Admin Contact]

---

*This customization guide provides comprehensive instructions for implementing the Reporter Service in your environment. Follow each step carefully and test thoroughly before deploying to production.*
