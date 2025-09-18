Option Explicit On 

Imports System.ServiceProcess
Imports System.Xml
Imports System.Text
Imports System.IO
Imports System.Diagnostics
Imports System.Data
Imports System.Web
Imports System.Web.Mail
Imports Microsoft.VisualBasic.ControlChars
Imports System.Timers
Imports System.Net
Imports System.Data.SqlClient
Imports System.Collections
Imports Microsoft.Win32
Imports log4net
Imports Spire.Pdf
Imports System.Reflection
Imports System.Configuration

Public Class ReporterService
    Inherits System.ServiceProcess.ServiceBase
    Private t As Timer
    Private log As EventLog

    ' Declare parameters
    Private MyInterval As String        ' Timer interval
    Private SelectNum As Integer        ' Select amount
    Private UserName As String          ' Database username
    Private PassWord As String          ' Database password
    Private DBName As String            ' Database name
    Private DBServer As String          ' Database server name
    Private Debug As String             ' Debug flag - "Y" or "N"
    Private Logging As String           ' Logging flag - "Y" or "N"
    Private reportServiceUrl As String       ' IP address of the report service endpoint

    ' Registry variables
    Private regKey As RegistryKey
    Private regSubKey As RegistryKey

    ' Misc variables
    Private NoInterval As Double
    Private InProcess As Boolean

    ' Logging declarations
    Private logfile As String
    Private errmsg As String
    Private myeventlog As log4net.ILog
    Private mydebuglog As log4net.ILog
    Private MachineName As String

    ' Enumerate object used for validation
    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
    End Enum

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '
        'If (System.Diagnostics.Debugger.IsAttached) Then
        'System.Diagnostics.Debugger.Break()
        'End If

        'If Environment.UserInteractive Then
        'Dim serviceToRun As ReporterService = New ReporterService
        'serviceToRun.OnStart(Nothing)
        'Console.WriteLine("Press any key to exit...")
        'Console.ReadLine()
        'serviceToRun.OnStop()
        'Else
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New ReporterService()}
        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
        'End If

    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'ReporterService
        '
        Me.ServiceName = "ReporterService"

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' This is executed by the service on startup
        ' In this instance, it creates a time that executes the interval that is specified
        ' The timer executes an event when it fires.  That event is where the interval
        ' code should be placed.

        Dim objReg As New CRegistry
        Dim RegValue As Object
        Dim TestKey As RegistryKey
        'Debugger.Break()

        ' ============================================
        ' Variable setup
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        MachineName = System.Environment.MachineName.ToString

        ' ============================================
        ' Get configuration variables
        Try
            reportServiceUrl = System.Configuration.ConfigurationManager.AppSettings.Get("reportServiceUrl")
        Catch ex As Exception
            Try
                Dim oEventLog As EventLog = New EventLog("Application")
                If Not Diagnostics.EventLog.SourceExists("ReporterService") Then
                    Diagnostics.EventLog.CreateEventSource("ReporterService", "Application")
                End If
                Diagnostics.EventLog.WriteEntry("ReporterService", ex.Message, System.Diagnostics.EventLogEntryType.Error)
                myeventlog.Error("ReporterService error " & ex.Message)
            Catch e As Exception
            End Try
            Throw ex
        End Try

        ' ============================================
        ' Retrieve registry entries and start timer
        Try
            ' ============================================
            ' If the key can be created, then do so and set default values
            TestKey = Registry.LocalMachine.OpenSubKey("Software\ReporterService")
            If TestKey Is Nothing Then
                objReg.CreateSubKey(objReg.HKeyLocalMachine, "Software\ReporterService")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "MyInterval", "5")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "UserName", "sa")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Password", "YourPasswordHere")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBName", "YourDatabaseName")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBServer", "YourServerName\YourInstanceName")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Debug", "Y")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Logging", "Y")
            End If

            ' ============================================
            ' Read parameters from Registry
            RegValue = ""
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "MyInterval", RegValue)
            MyInterval = RegValue.ToString
            SelectNum = Int(MyInterval)
            SelectNum = Int(SelectNum / 2)
            If SelectNum < 1 Then SelectNum = 1
            NoInterval = Val(MyInterval) * 1000
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "UserName", RegValue)
            UserName = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Password", RegValue)
            PassWord = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBName", RegValue)
            DBName = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "DBServer", RegValue)
            DBServer = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Debug", RegValue)
            Debug = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReporterService", "Logging", RegValue)
            Logging = RegValue.ToString

            ' ============================================
            ' Setup timer
            t = New Timer()
            AddHandler t.Elapsed, AddressOf TimerFired
            With t
                .Interval = NoInterval
                .AutoReset = True
                .Enabled = True
                .Start()
            End With
            InProcess = False

            ' ============================================
            ' Log to event viewer log
            Diagnostics.EventLog.WriteEntry("ReporterService", "Service Starting")

            ' ============================================
            ' Open debug log file if applicable
            If Debug = "Y" Or Logging = "Y" Then
                Dim versionNumber As Version
                versionNumber = Assembly.GetExecutingAssembly().GetName().Version

                Dim path As String
                path = "C:\Logs\"
                logfile = path & "ReporterService.log"
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("ReporterService Trace Log Started " & Format(Now))
                mydebuglog.Debug("Version: " & versionNumber.ToString)
                If Debug = "Y" Then
                    mydebuglog.Debug("PARAMETERS")
                    mydebuglog.Debug("  Debug: " & Debug)
                    mydebuglog.Debug("  Logging: " & Logging)
                    mydebuglog.Debug("  reportServiceUrl: " & reportServiceUrl)
                    mydebuglog.Debug("  MyInterval: " & MyInterval)
                    mydebuglog.Debug("  UserName: " & UserName)
                    mydebuglog.Debug("  PassWord: " & PassWord)
                    mydebuglog.Debug("  DBName: " & DBName)
                    mydebuglog.Debug("  DBServer: " & DBServer)
                End If
                If reportServiceUrl = "" Then reportServiceUrl = "192.168.1.100"
            End If

            ' Log start to syslog
            myeventlog.Info("ReportService Started")

        Catch obug As Exception
            Try
                Dim oEventLog As EventLog = New EventLog("Application")
                If Not Diagnostics.EventLog.SourceExists("ReporterService") Then
                    Diagnostics.EventLog.CreateEventSource("ReporterService", "Application")
                End If
                Diagnostics.EventLog.WriteEntry("ReporterService", obug.Message, System.Diagnostics.EventLogEntryType.Error)
                myeventlog.Error("ReporterService error " & obug.Message)
            Catch e As Exception
            End Try
            Throw obug
        End Try

    End Sub

    Protected Overrides Sub OnStop()
        ' This is executed when the service stops
        Try
            ' ============================================
            ' Close diagnostic logging
            If Debug = "Y" Or Logging = "Y" Then
                mydebuglog.Debug(vbCrLf & "Service Results: " & Trim(errmsg))
                mydebuglog.Debug("ReporterService Trace Log Ended " & Format(Now))
                mydebuglog.Debug("----------------------------------")
                mydebuglog = Nothing
            End If

            ' ============================================
            ' Stop timer
            Diagnostics.EventLog.WriteEntry("ReporterService", "Service Stopping: " & EventLogEntryType.Information)
            myeventlog.Info("ReportService Stopped")
            If Not t Is Nothing Then
                Try
                    t.Stop()
                    t.Dispose()
                Catch ex As Exception
                    Diagnostics.EventLog.WriteEntry("ReporterService", "Unable to stop timer")
                    myeventlog.Error("ReporterService unable to stop timer")
                End Try
            End If
        Catch obug As Exception
            Diagnostics.EventLog.WriteEntry("ReporterService", obug.Message)
            myeventlog.Error("ReporterService " & obug.Message)
        End Try

        myeventlog = Nothing
    End Sub

    Private Sub TimerFired(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        ' The timer fired.  Call CheckReporterQueue
        'If Debug = "Y" Then mydebuglog.Debug("   ~ timer fired. InProcess: " & InProcess.ToString)
        If Not InProcess Then
            InProcess = True
            CheckReporterQueue()
        End If
    End Sub

    Private Sub CheckReporterQueue()

        'If Debug = "Y" Then mydebuglog.Debug("   ~ in CheckReporterQueue")

        ' Web service declarations
        Dim response As HttpWebResponse = Nothing
        Dim Completed As String
        Dim LineNumber As Integer
        Dim ReporterWS As New com.yourcompany.reporting.Service
        Dim BasicWS As New com.yourcompany.basic.Service
        Dim http As New simplehttp()
        Dim ResultString As String

        ' Database declarations
        Dim con As SqlConnection
        Dim con2 As SqlConnection
        Dim cmd As SqlCommand
        Dim cmd2 As SqlCommand
        Dim dr As SqlDataReader
        Dim dr2 As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        'Dim svcresults As XmlElement        
        Dim returnv As Integer

        ' Report Queue Variables
        Dim NUMBER_EXECUTED, FREQUENCY_COUNT As Integer
        Dim JOB_ID, ENT_ID, FREQUENCY_TYPE, REPORT_FORMAT, REPORT_ID, NOTIFY_FLG, CON_ID, DETAIL_ID, SQL_LOG As String
        Dim START_DT, END_DT, PARAMETER, RESCHED_ERROR_FLG, REP_DEST_ID, REP_ES_ID, PRINT_QUEUE, PROCESS_QUEUE_ID, PRINTER_ID As String
        Dim wp As String

        ' Print related variables
        Dim d_dsize, d_item_name, d_ext, BinaryFile, perror As String
        Dim startIndex As Long = 0
        Dim appath As String = AppDomain.CurrentDomain.BaseDirectory
        Dim retval As Long
        Dim outbyte(1000) As Byte
        Dim bfs As FileStream
        Dim bw As BinaryWriter
        Dim printed As Boolean        

        'If Debug = "Y" Then mydebuglog.Debug("   ~ declarations completed")
        Try
            Completed = ""
            ResultString = ""
            LineNumber = 0
            InProcess = True
            printed = False
            perror = ""
            REP_ES_ID = ""
            PRINT_QUEUE = ""
            PROCESS_QUEUE_ID = ""
            PRINTER_ID = ""
            SQL_LOG = ""

            ' ============================================
            ' Open database connection 
            ConnS = "server=" & DBServer & ";uid=" & UserName & ";pwd=" & PassWord & ";database=" & DBName
            'If Debug = "Y" Then mydebuglog.Debug("   ~ opening database " & ConnS)
            errmsg = OpenDBConnection(ConnS, con, cmd)
            errmsg = OpenDBConnection(ConnS, con2, cmd2)
            If errmsg <> "" Then
                mydebuglog.Debug("Unable to open database connection: " & errmsg)
                myeventlog.Error("ReporterService : Unable to open database connection: " & errmsg)
                GoTo CloseOut
            End If
            'If Debug = "Y" Then mydebuglog.Debug("   ~ database opened")

            ' ============================================
            ' Lock the queue entries to process
            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                "SET LOCKED_BY = '" & MachineName & "' " & _
                "WHERE ROW_ID IN " & _
                "(SELECT TOP " & SelectNum.ToString.Trim() & " ROW_ID " & _
                "FROM YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                "WHERE NEXT_EXECUTE <= GETDATE() AND (LOCKED_BY IS NULL OR LOCKED_BY=''))"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & " ...." & vbCrLf & vbCrLf & "Locking queue entries at " & Format(Now))
                mydebuglog.Debug("Query Locking queue: " & vbCrLf & SqlS & vbCrLf)
            End If
            Try
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then
                    mydebuglog.Debug("No records found in the queue")
                    GoTo CloseOut
                End If
            Catch ex2 As Exception
                mydebuglog.Debug("Unable to Lock the queue entries: " & ex2.ToString)
                GoTo CloseOut
            End Try

            ' ============================================
            ' Get the queue entries to process 
            SqlS = "SELECT S.ROW_ID, S.ENT_ID, S.NUMBER_EXECUTED, S.FREQUENCY_COUNT, S.FREQUENCY_TYPE, " & _
                "S.REPORT_FORMAT, S.REPORT_ID, S.NOTIFY_FLG, S.CON_ID, CONVERT(VARCHAR, S.START_DT, 101), " & _
                "CONVERT(VARCHAR, S.END_DT, 101), S.PARAMETER, S.RESCHED_ERROR_FLG, E.REP_DEST_ID, S.REP_ES_ID, " & _
                "P.DESC_TEXT, Q.ROW_ID, R.SQL_LOG, S.PRINTER_ID " & _
                "FROM YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE S " & _
                "LEFT OUTER JOIN YourDatabase.dbo.REPORT_ENTITIES E ON E.ROW_ID=S.ENT_ID " & _
                "LEFT OUTER JOIN YourDatabase.dbo.LIST_OF_VALUES P ON P.ROW_ID=S.PRINTER_ID " & _
                "LEFT OUTER JOIN YourDatabase.dbo.PROCESSING_QUEUE Q ON Q.REP_ES_ID=S.REP_ES_ID " & _
                "LEFT OUTER JOIN YourDatabase.dbo.REPORTS R ON R.ROW_ID=PROCESS_REPORT_ID " & _
                "WHERE S.LOCKED_BY = '" & MachineName & "'"
            If Debug = "Y" Then
                mydebuglog.Debug("Checking for queue entries at " & Format(Now))
                mydebuglog.Debug("Query for reports to process: " & vbCrLf & SqlS & vbCrLf)
            End If
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then

                ' Go through records
                Try
                    While dr.Read()
                        LineNumber = LineNumber + 1

                        ' Retrieve data from the table
                        JOB_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        ENT_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                        NUMBER_EXECUTED = CheckDBNull(dr(2), enumObjectType.IntType)
                        FREQUENCY_COUNT = CheckDBNull(dr(3), enumObjectType.IntType)
                        FREQUENCY_TYPE = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                        REPORT_FORMAT = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                        REPORT_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                        NOTIFY_FLG = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                        CON_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType)).ToString
                        START_DT = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                        END_DT = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                        PARAMETER = Trim(CheckDBNull(dr(11), enumObjectType.StrType)).ToString
                        RESCHED_ERROR_FLG = Trim(CheckDBNull(dr(12), enumObjectType.StrType)).ToString
                        REP_DEST_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType)).ToString
                        REP_ES_ID = Trim(CheckDBNull(dr(14), enumObjectType.StrType)).ToString
                        PRINT_QUEUE = Trim(CheckDBNull(dr(15), enumObjectType.StrType)).ToString
                        PROCESS_QUEUE_ID = Trim(CheckDBNull(dr(16), enumObjectType.StrType)).ToString
                        SQL_LOG = Trim(CheckDBNull(dr(17), enumObjectType.StrType)).ToString
                        PRINTER_ID = Trim(CheckDBNull(dr(18), enumObjectType.StrType))

                        If Debug = "Y" Then
                            mydebuglog.Debug("Entry found # " & LineNumber.ToString)
                            mydebuglog.Debug("  >JOB_ID: " & JOB_ID)
                            mydebuglog.Debug("  >ENT_ID: " & ENT_ID)
                            mydebuglog.Debug("  >CON_ID: " & CON_ID)
                            mydebuglog.Debug("  >REP_DEST_ID: " & REP_DEST_ID)
                            mydebuglog.Debug("  >REP_ES_ID: " & REP_ES_ID)
                            mydebuglog.Debug("  >PRINT_QUEUE: " & PRINT_QUEUE)
                            mydebuglog.Debug("  >PRINTER_ID: " & PRINTER_ID)
                            mydebuglog.Debug("  >FREQUENCY_TYPE: " & FREQUENCY_TYPE)
                            mydebuglog.Debug("  >FREQUENCY_COUNT: " & FREQUENCY_COUNT.ToString)
                            mydebuglog.Debug("  >RESCHED_ERROR_FLG: " & RESCHED_ERROR_FLG)
                            mydebuglog.Debug("  >NUMBER_EXECUTED: " & NUMBER_EXECUTED.ToString & vbCrLf)
                        End If

                        ' If a CX_REP_ENT_SCHED.ROW_ID is supplied, reference this when calling the ExecReport service instead
                        ' of constructing a new record
                        If REP_ES_ID <> "" And FREQUENCY_TYPE = "" Then

                            ' Generate the service string
                            wp = "<ReportList>" & _
                                "<ReportItem>" & _
                                "<Debug>" & Debug & "</Debug>" & _
                                "<Database></Database>" & _
                                "<ReportId>" & REP_ES_ID & "</ReportId>" & _
                                "<REPORT></REPORT>" & _
                                "<SCREEN></SCREEN>" & _
                                "<DT_FLAG>N</DT_FLAG>" & _
                                "<REP_FILENAME></REP_FILENAME>" & _
                                "<REP_ID></REP_ID>" & _
                                "<FORMAT></FORMAT>" & _
                                "<FORMAT_CODE></FORMAT_CODE>" & _
                                "<FORMAT_EXTENSION></FORMAT_EXTENSION>" & _
                                "<PRICE_LIST></PRICE_LIST>" & _
                                "<START_DATE></START_DATE>" & _
                                "<END_DATE></END_DATE>" & _
                                "<PARAMETER></PARAMETER>" & _
                                "<DEFAULT_EMAIL></DEFAULT_EMAIL>" & _
                                "<NOTIFY_FLG></NOTIFY_FLG>" & _
                                "<CONTACT_ID></CONTACT_ID>" & _
                                "<ACCOUNT_ID></ACCOUNT_ID>" & _
                                "<FST_NAME></FST_NAME>" & _
                                "<LAST_NAME></LAST_NAME>" & _
                                "<EMAIL_ADDR></EMAIL_ADDR>" & _
                                "<DESTINATION></DESTINATION>" & _
                                "<REP_DESC></REP_DESC>" & _
                                "<ENT_ID></ENT_ID>" & _
                                "<REP_DEST_ID></REP_DEST_ID>" & _
                                "<SUB_ID></SUB_ID>" & _
                                "<SQL_REP></SQL_REP>" & _
                                "<SQL_MOD></SQL_MOD>" & _
                                "<USER_ID></USER_ID>" & _
                                "<SESSION_ID></SESSION_ID>" & _
                                "<TRAN_ID></TRAN_ID>" & _
                                "<DMS_FLAG></DMS_FLAG>" & _
                                "<SAVE_PORTAL></SAVE_PORTAL>" & _
                                "</ReportItem>" & _
                                "</ReportList>"
                        Else
                            ' Generate the service string
                            wp = "<ReportList>" & _
                                "<ReportItem>" & _
                                "<Debug>" & Debug & "</Debug>" & _
                                "<Database></Database>" & _
                                "<ReportId></ReportId>" & _
                                "<REPORT></REPORT>" & _
                                "<SCREEN></SCREEN>" & _
                                "<DT_FLAG>N</DT_FLAG>" & _
                                "<REP_FILENAME></REP_FILENAME>" & _
                                "<REP_ID>" & REPORT_ID & "</REP_ID>" & _
                                "<FORMAT>" & REPORT_FORMAT & "</FORMAT>" & _
                                "<FORMAT_CODE></FORMAT_CODE>" & _
                                "<FORMAT_EXTENSION></FORMAT_EXTENSION>" & _
                                "<PRICE_LIST></PRICE_LIST>" & _
                                "<START_DATE>" & START_DT & "</START_DATE>" & _
                                "<END_DATE>" & END_DT & "</END_DATE>" & _
                                "<PARAMETER>" & PARAMETER & "</PARAMETER>" & _
                                "<DEFAULT_EMAIL></DEFAULT_EMAIL>" & _
                                "<NOTIFY_FLG>" & NOTIFY_FLG & "</NOTIFY_FLG>" & _
                                "<CONTACT_ID>" & CON_ID & "</CONTACT_ID>" & _
                                "<ACCOUNT_ID></ACCOUNT_ID>" & _
                                "<FST_NAME></FST_NAME>" & _
                                "<LAST_NAME></LAST_NAME>" & _
                                "<EMAIL_ADDR></EMAIL_ADDR>" & _
                                "<DESTINATION></DESTINATION>" & _
                                "<REP_DESC></REP_DESC>" & _
                                "<ENT_ID>" & ENT_ID & "</ENT_ID>" & _
                                "<REP_DEST_ID>" & REP_DEST_ID & "</REP_DEST_ID>" & _
                                "<SUB_ID></SUB_ID>" & _
                                "<SQL_REP></SQL_REP>" & _
                                "<SQL_MOD></SQL_MOD>" & _
                                "<USER_ID></USER_ID>" & _
                                "<SESSION_ID></SESSION_ID>" & _
                                "<TRAN_ID></TRAN_ID>" & _
                                "<DMS_FLAG></DMS_FLAG>" & _
                                "<SAVE_PORTAL></SAVE_PORTAL>" & _
                                "</ReportItem>" & _
                                "</ReportList>"
                        End If

                        ' ============================================
                        ' Prepare SQL_LOG query
                        '  Substitute out SessionId, WorkshopId, TrainerId, ParticipantId, Parameter, Param
                        If SQL_LOG <> "" Then
                            SQL_LOG = Replace(SQL_LOG, "[SessionId]", SQL_LOG)
                            SQL_LOG = Replace(SQL_LOG, "[WorkshopId]", SQL_LOG)
                            SQL_LOG = Replace(SQL_LOG, "[TrainerId]", SQL_LOG)
                            SQL_LOG = Replace(SQL_LOG, "[ParticipantId]", SQL_LOG)
                            SQL_LOG = Replace(SQL_LOG, "[PARAMETER]", SQL_LOG)
                            SQL_LOG = Replace(SQL_LOG, "[Param]", SQL_LOG)
                        End If

                        ' ============================================
                        ' Generate the report
                        Try
                            Completed = "Failure"
                            Try
                                ' Send the web request
                                If Debug = "Y" Then mydebuglog.Debug("Writing to ExecReport at " & Now.ToString & " -  XML: " & wp)
                                Try
                                    'svcresults = ReporterWS.ExecReport(wp)                                    
                                    'If Debug = "Y" Then mydebuglog.Debug("  > raw results: " & svcresults.InnerText)
                                    'If svcresults.InnerText.Contains("Success") Or svcresults.InnerText.Contains("No records were found for this query.") Then
                                    '    Completed = "Success"
                                    'Else
                                    '    Completed = "Failure"
                                    'End If
                                    ResultString = http.geturl("http://yourcompany.com/reporting/service.asmx/ExecReport?sXML=" & wp, reportServiceUrl, 80, "", "")
                                    If Debug = "Y" Then mydebuglog.Debug("  > raw results: " & ResultString)
                                    If ResultString.Contains("Success") Or ResultString.Contains("No records were found for this query.") Then
                                        Completed = "Success"
                                    Else
                                        Completed = "Failure"
                                    End If
                                Catch ex4 As Exception
                                    If Debug = "Y" Then mydebuglog.Debug("Unable to execute the web service: " & ex4.ToString)
                                    Completed = "Failure"
                                    BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "25", "Unable to execute the web service: " & ex4.ToString, "admin@yourcompany.com", Debug)
                                End Try

                                ' Parse the results
                                If Completed = "Success" And ResultString <> "" Then
                                    'Dim xitems As System.Xml.XmlNodeList = svcresults.GetElementsByTagName("Report")
                                    'For Each item As XmlNode In xitems
                                    'REP_ES_ID = item.SelectSingleNode("ReportId").InnerText
                                    'Next
                                    Try
                                        Dim svcdoc As New XmlDocument
                                        svcdoc.LoadXml(ResultString)
                                        Dim xitems As System.Xml.XmlNodeList
                                        xitems = svcdoc.GetElementsByTagName("Report")
                                        For Each item As XmlNode In xitems
                                            REP_ES_ID = item.SelectSingleNode("ReportId").InnerText
                                        Next
                                    Catch ex4 As Exception
                                        If Debug = "Y" Then mydebuglog.Debug("Unable to parse service results: " & ex4.ToString)
                                        Completed = "Failure"
                                        BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "25", "Unable to execute the web service: " & ex4.ToString, "admin@yourcompany.com", Debug)
                                    End Try
                                End If
                                If Debug = "Y" Then mydebuglog.Debug("  > REP_ES_ID: " & REP_ES_ID & vbCrLf)

                            Catch ex3 As Exception
                                If Debug = "Y" Then mydebuglog.Debug("Unable to execute the web service: " & ex3.ToString)
                                myeventlog.Error("ReporterService Unable to execute the web service: " & ex3.Message)
                                GoTo CloseOut
                            End Try
                            myeventlog.Info("ReporterService executed Job Id " & JOB_ID & " for Entitlement Id " & ENT_ID & ", Results: " & Completed)
                            If Debug = "Y" Then mydebuglog.Debug("GenerateReport results at " & Now.ToString & ": " & Completed & vbCrLf)

                            ' Report Succeeded
                            If Completed = "Success" Then

                                ' If PRINT_QUEUE supplied, send the report to the specified queue
                                If PRINT_QUEUE <> "" And REP_ES_ID <> "" Then

                                    ' Retrieve the report image and save locally
                                    appath = appath & "\temp\"
                                    BinaryFile = ""
                                    SqlS = "SELECT TOP 1 D.ROW_ID, D.DIMAGE, D.DSIZE, D.FORMAT " & _
                                     "FROM YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE D " & _
                                     "WHERE D.ROW_ID='" & REP_ES_ID & "'"
                                    If Debug = "Y" Then mydebuglog.Debug("   > Get report document: " & SqlS)
                                    Try
                                        cmd2.CommandText = SqlS
                                        dr2 = cmd2.ExecuteReader()
                                        If Not dr2 Is Nothing Then
                                            While dr2.Read()
                                                d_item_name = Trim(CheckDBNull(dr2(0), enumObjectType.StrType))
                                                d_dsize = Trim(CheckDBNull(dr2(2), enumObjectType.StrType))
                                                d_ext = Trim(CheckDBNull(dr2(3), enumObjectType.StrType)).ToLower
                                                If d_dsize <> "" Then
                                                    If Debug = "Y" Then mydebuglog.Debug("   > Saving report document")
                                                    ReDim outbyte(Val(d_dsize) - 1)
                                                    startIndex = 0
                                                    Try
                                                        retval = dr2.GetBytes(1, 0, outbyte, 0, d_dsize)
                                                    Catch ex4 As Exception
                                                        perror = "Unable to retrieve the report document: " & ex4.ToString & vbCrLf
                                                        If Debug = "Y" Then mydebuglog.Debug("Unable to retrieve the report document: " & ex4.ToString)
                                                        myeventlog.Error("Unable to retrieve the report document: " & ex4.Message)
                                                    End Try

                                                    If Debug = "Y" Then
                                                        mydebuglog.Debug("   > Document Information.  Extension: " & d_ext)
                                                        mydebuglog.Debug("                            Item Name: " & d_item_name)
                                                        mydebuglog.Debug("                            Size: " & d_dsize)
                                                        mydebuglog.Debug("                            Bytes retrieved: " & retval.ToString)
                                                    End If

                                                    If perror = "" Then

                                                        ' Create temp file name
                                                        If d_ext <> "" Then
                                                            If Right(appath.Trim, 1) <> "\" Then appath = appath.Trim & "\"
                                                            BinaryFile = appath & REP_ES_ID & "." & d_ext
                                                            If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   > Computed attachment filename: " & BinaryFile)
                                                        End If

                                                        ' Retrieve item and store-to/replace-in the filesystem if found
                                                        If retval > 0 And BinaryFile <> "" Then
                                                            If (My.Computer.FileSystem.FileExists(BinaryFile)) Then Kill(BinaryFile)
                                                            Try
                                                                bfs = New FileStream(BinaryFile, FileMode.Create, FileAccess.Write)
                                                                bw = New BinaryWriter(bfs)
                                                                bw.Write(outbyte)
                                                                bw.Flush()
                                                                bw.Close()
                                                            Catch ex4 As Exception
                                                                perror = perror & "Unable to write the file to a temp file: " & ex4.ToString & vbCrLf
                                                                If Debug = "Y" Then mydebuglog.Debug("Unable to write the file to a temp file: " & ex4.ToString)
                                                                myeventlog.Error("Unable to write the file to a temp file: " & ex4.Message)
                                                                retval = 0
                                                            End Try

                                                            ' Close binary file
                                                            Try
                                                                bfs = Nothing
                                                                bw = Nothing
                                                                outbyte = Nothing
                                                            Catch ex As Exception
                                                            End Try

                                                            ' If the file is found print the file
                                                            If My.Computer.FileSystem.FileExists(BinaryFile) Then

                                                                ' SPIRE.PDF
                                                                Dim pdfdoc As New PdfDocument
                                                                Dim printdocument As New System.Drawing.Printing.PrintDocument
                                                                Dim PageCount As Integer
                                                                PageCount = 0
                                                                Try
                                                                    pdfdoc.LoadFromFile(BinaryFile)
                                                                    PageCount = pdfdoc.Pages.Count
                                                                Catch ex5 As Exception
                                                                    perror = perror & "Unable to load document " & BinaryFile & ": " & ex5.ToString & vbCrLf
                                                                    myeventlog.Error("Unable to load document " & BinaryFile & ": " & ex5.Message)
                                                                End Try
                                                                If Debug = "Y" Then mydebuglog.Debug("   > Pages: " & PageCount.ToString)
                                                                If PageCount > 0 Then
                                                                    If Debug = "Y" Then mydebuglog.Debug("   > Printing " & BinaryFile & " to: " & PRINT_QUEUE)
                                                                    Try
                                                                        pdfdoc.PageSettings.Size = PdfPageSize.Letter
                                                                        pdfdoc.PrintDocument.DefaultPageSettings.Landscape = False
                                                                        pdfdoc.PageScaling = PdfPrintPageScaling.FitSize
                                                                        pdfdoc.PrintFromPage = 1
                                                                        pdfdoc.PrintToPage = PageCount
                                                                        pdfdoc.PrinterName = PRINT_QUEUE
                                                                        printdocument = pdfdoc.PrintDocument
                                                                        printdocument.PrintController = New System.Drawing.Printing.StandardPrintController
                                                                        printdocument.Print()
                                                                        printed = True                                                                        
                                                                    Catch ex6 As Exception
                                                                        perror = perror & "Unable to print: " & ex6.ToString & vbCrLf
                                                                        myeventlog.Error("Unable to print: " & ex6.Message)
                                                                        BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "24", perror, "admin@yourcompany.com", Debug)
                                                                    End Try
                                                                Else
                                                                    If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "   > Document " & d_item_name & " has no pages to print")
                                                                End If

                                                                ' Delete temp file
                                                                Try
                                                                    If Debug = "Y" Then mydebuglog.Debug("   > Deleting " & BinaryFile)
                                                                    Kill(BinaryFile)
                                                                    pdfdoc = Nothing
                                                                    printdocument = Nothing
                                                                Catch ex4 As Exception
                                                                    perror = perror & "Unable to delete the temp file: " & ex4.ToString & vbCrLf
                                                                    If Debug = "Y" Then mydebuglog.Debug("Unable to delete the temp file: " & ex4.ToString)
                                                                    myeventlog.Error("Unable to delete the temp file: " & ex4.Message)
                                                                End Try
                                                            End If
                                                        End If
                                                    End If

                                                End If
                                            End While

                                            ' Close datareader
                                            Try
                                                dr2.Close()
                                            Catch ex As Exception
                                            End Try
                                        End If

                                    Catch ex3 As Exception
                                        perror = perror & "Unable to print the report document: " & ex3.ToString & vbCrLf
                                        If Debug = "Y" Then mydebuglog.Debug("Unable to print the report document: " & ex3.ToString)
                                        myeventlog.Error("Unable to print the report document: " & ex3.Message)
                                    End Try
                                End If

                                ' If SQL_LOG is supplied, process it
                                If SQL_LOG <> "" Then
                                    If Debug = "Y" Then mydebuglog.Debug("Processing SQL_LOG: " & vbCrLf & SQL_LOG & vbCrLf)
                                    Try
                                        cmd2.CommandText = SQL_LOG
                                        returnv = cmd2.ExecuteNonQuery()
                                        If returnv = 0 Then
                                            If Debug = "Y" Then mydebuglog.Debug("Unable to process SQL_LOG" & vbCrLf)
                                            perror = perror & "Unable to process SQL_LOG" & vbCrLf
                                        End If
                                    Catch ex3 As Exception
                                        perror = perror & "Unable to process SQL_LOG: " & ex3.ToString & vbCrLf
                                        If Debug = "Y" Then mydebuglog.Debug("Unable to process SQL_LOG: " & ex3.ToString)
                                        BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "26", ex3.ToString, "admin@yourcompany.com", Debug)
                                        GoTo CloseOut
                                    End Try

                                End If

                                ' If REP_ES_ID supplied, update the related CX_PROCESS_QUEUE record
                                If PROCESS_QUEUE_ID <> "" Then
                                    If printed Then
                                        SqlS = "UPDATE YourDatabase.dbo.PROCESSING_QUEUE " & _
                                            "SET STATUS_CD='Printed', PRINTER_ID='" & PRINTER_ID & "', EXECUTED_DT=GETDATE(), LAST_UPD=GETDATE() " & _
                                            "WHERE ROW_ID='" & PROCESS_QUEUE_ID & "'"
                                        perror = "Printed report to " & PRINT_QUEUE
                                    Else
                                        SqlS = "UPDATE YourDatabase.dbo.PROCESSING_QUEUE " & _
                                            "SET STATUS_CD='Error', EXECUTED_DT=GETDATE(), LAST_UPD=GETDATE() " & _
                                            "WHERE ROW_ID='" & PROCESS_QUEUE_ID & "'"
                                        perror = "Unable to print report to " & PRINT_QUEUE & vbCrLf & perror
                                    End If
                                    If Debug = "Y" Then mydebuglog.Debug("Update Process Queue record: " & vbCrLf & SqlS & vbCrLf)
                                    Try
                                        cmd2.CommandText = SqlS
                                        returnv = cmd2.ExecuteNonQuery()
                                        If returnv = 0 Then
                                            If Debug = "Y" Then mydebuglog.Debug("Unable to update Process Queue record - not found" & vbCrLf)
                                            perror = perror & "Unable to update Process Queue record" & vbCrLf
                                        End If
                                    Catch ex3 As Exception
                                        perror = perror & "Unable to update Process Queue record: " & ex3.ToString & vbCrLf
                                        If Debug = "Y" Then mydebuglog.Debug("Unable to update Process Queue record: " & ex3.ToString)
                                        BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "27", ex3.ToString, "admin@yourcompany.com", Debug)
                                        GoTo CloseOut
                                    End Try

                                    ' Log to CX_PROCESS_QUEUE_DETAIL
                                    perror = Left(ChkString(perror), 500)
                                    DETAIL_ID = BasicWS.GenerateRecordId("PROCESSING_QUEUE_DETAIL", "N", Debug)
                                    SqlS = "INSERT YourDatabase.dbo.PROCESSING_QUEUE_DETAIL " & _
                                        "(CONFLICT_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, MODIFICATION_NUM, " & _
                                        "ROW_ID, QUEUE_ID, MACHINE, MESSAGE) " & _
                                        "VALUES(0, GETDATE(), '1-1QV1U', GETDATE(), '1-1QV1U', 0, " & _
                                        "'" & DETAIL_ID & "','" & PROCESS_QUEUE_ID & "','" & MachineName & "','" & perror & "')"
                                    If Debug = "Y" Then mydebuglog.Debug("Insert Process Queue Detail record: " & vbCrLf & SqlS & vbCrLf)
                                    Try
                                        cmd2.CommandText = SqlS
                                        returnv = cmd2.ExecuteNonQuery()
                                        If returnv = 0 Then
                                            mydebuglog.Debug("Unable to Insert Process Queue Detail record" & vbCrLf)
                                        End If
                                    Catch ex3 As Exception
                                        mydebuglog.Debug("Unable to update Process Queue: " & ex3.ToString)
                                        GoTo CloseOut
                                    End Try

                                End If

                                ' Generate query to update the queue entry to reschedule as applicable
                                SqlS = ""

                                If NUMBER_EXECUTED < FREQUENCY_COUNT Or FREQUENCY_COUNT = 0 Then
                                    ' Reschedule job
                                    If Debug = "Y" Then mydebuglog.Debug("Rescheduling job")
                                    Select Case FREQUENCY_TYPE
                                        Case "Daily"
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "START_DT = DATEADD(dd,1,START_DT), " & _
                                                "END_DT = DATEADD(dd,1,END_DT), " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = DATEADD(dd,1,NEXT_EXECUTE), " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                        Case "Weekly"
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "START_DT = DATEADD(dd,7,START_DT), " & _
                                                "END_DT = DATEADD(dd,7,END_DT), " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = DATEADD(dd,7,NEXT_EXECUTE), " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                        Case "Monthly"
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "START_DT = DATEADD(mm,1,START_DT), " & _
                                                "END_DT = DATEADD(mm,1,END_DT), " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = DATEADD(mm,1,NEXT_EXECUTE), " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                        Case "Quarterly"
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "START_DT = DATEADD(mm,3,START_DT), " & _
                                                "END_DT = DATEADD(mm,3,END_DT), " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = DATEADD(mm,3,NEXT_EXECUTE), " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                        Case "Yearly"
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "START_DT = DATEADD(mm,12,START_DT), " & _
                                                "END_DT = DATEADD(mm,12,END_DT), " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = DATEADD(mm,12,NEXT_EXECUTE), " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                        Case Else
                                            SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                                "SET LOCKED_BY = NULL,  " & _
                                                "LAST_STATUS = 'Success', " & _
                                                "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                                "LAST_EXECUTED = GETDATE(), " & _
                                                "NEXT_EXECUTE = NULL, " & _
                                                "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                                "WHERE ROW_ID='" & JOB_ID & "'"
                                    End Select

                                Else
                                    ' Do not reschedule job - finished count
                                    If Debug = "Y" Then mydebuglog.Debug("Finalizing job")
                                    SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                        "SET LOCKED_BY = NULL,  " & _
                                        "LAST_STATUS = 'Success', " & _
                                        "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                        "LAST_EXECUTED = GETDATE(), " & _
                                        "NEXT_EXECUTE = NULL, " & _
                                        "REP_ES_ID = '" & REP_ES_ID & "' " & _
                                        "WHERE ROW_ID='" & JOB_ID & "'"
                                End If
                            Else
                                ' Failed
                                SqlS = ""
                                If RESCHED_ERROR_FLG = "Y" Then
                                    If Debug = "Y" Then mydebuglog.Debug("Rescheduling job due to failure")
                                    SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                       "SET LOCKED_BY = NULL,  " & _
                                       "LAST_STATUS = 'Failed', " & _
                                       "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                       "REP_ES_ID = '" & REP_ES_ID & "', " & _
                                       "LAST_EXECUTED = GETDATE(), " & _
                                       "NEXT_EXECUTE = DATEADD(mi,10,GETDATE()) " & _
                                       "WHERE ROW_ID='" & JOB_ID & "'"
                                Else
                                    If Debug = "Y" Then mydebuglog.Debug("Clearing job lock due to failure")
                                    SqlS = "UPDATE YourDatabase.dbo.REPORT_EXECUTION_SCHEDULE " & _
                                       "SET LOCKED_BY = NULL,  " & _
                                       "LAST_STATUS = 'Failed', " & _
                                       "NUMBER_EXECUTED = NUMBER_EXECUTED + 1, " & _
                                       "REP_ES_ID = '" & REP_ES_ID & "', " & _
                                       "LAST_EXECUTED = GETDATE(), " & _
                                       "NEXT_EXECUTE = NULL " & _
                                       "WHERE ROW_ID='" & JOB_ID & "'"
                                End If
                            End If

                            ' Report Execute the update query
                            If SqlS <> "" Then
                                If Debug = "Y" Then mydebuglog.Debug("Query Update queue: " & vbCrLf & SqlS & vbCrLf)
                                Try
                                    cmd2.CommandText = SqlS
                                    returnv = cmd2.ExecuteNonQuery()
                                    If returnv = 0 Then
                                        mydebuglog.Debug("Unable to update entry - not found" & vbCrLf)
                                    End If
                                Catch ex3 As Exception
                                    mydebuglog.Debug("Unable to update entry: " & ex3.ToString)
                                        BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "28", ex3.ToString, "admin@yourcompany.com", Debug)
                                End Try
                            End If

                        Catch ex2 As Exception
                            If Debug = "Y" Or Logging = "Y" Then
                                mydebuglog.Debug("Exception raised " & Str(ex2))
                                myeventlog.Error("ReporterService exception raised: " & vbCrLf & ex2.Message)
                                Diagnostics.EventLog.WriteEntry("ReporterService", "Exception raised " & ex2.Message)
                            End If
                            BasicWS.ErrorNotice(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, "29", ex2.ToString, "admin@yourcompany.com", Debug)
                            Completed = "Error"
                        End Try

                    End While
                Catch ex As Exception
                    errmsg = errmsg & "Error reading record. " & ex.ToString
                    GoTo CloseOut
                End Try
            Else
                dr.Close()
                errmsg = errmsg & "The contact was not found."
                GoTo CloseOut
            End If


CloseOut:
            ' ============================================
            ' Close database connections
            Try
                If Not con Is Nothing Then con.Close()
                If Not dr Is Nothing Then dr.Close()
                con.Dispose()
                con = Nothing
                cmd.Dispose()
                cmd = Nothing
                dr = Nothing
                ReporterWS = Nothing
            Catch ex As Exception
                Diagnostics.EventLog.WriteEntry("ReporterService", "Unable to close database connection " & ex.Message)
                If Debug = "Y" Then mydebuglog.Debug("Unable to close database connection " & ex.Message)
            End Try

            ' ============================================
            ' Debug output of results
            If Debug = "Y" And LineNumber > 0 Then
                mydebuglog.Debug("Total entries processed: " & Str(LineNumber) & vbCrLf)
                If errmsg <> "" Then
                    mydebuglog.Debug("Error messages recported: " & vbCrLf & errmsg)
                End If
            End If
            InProcess = False

        Catch wex As WebException
            If Not wex.Response Is Nothing Then
                Dim errorResponse As HttpWebResponse = Nothing
                Try
                    errorResponse = DirectCast(wex.Response, HttpWebResponse)
                    ' Save error description for log
                    errmsg = errorResponse.StatusDescription
                    Diagnostics.EventLog.WriteEntry("ReporterService", "WEB SERVICE ERROR RECEIVED:  " & errmsg)
                Finally
                    If Not errorResponse Is Nothing Then errorResponse.Close()
                End Try
            End If
            Diagnostics.EventLog.WriteEntry("ReporterService", errmsg)
            If Debug = "Y" Then mydebuglog.Debug("Web service error: " & errmsg & vbCrLf)

            ' ============================================
            ' Close database connections
            Try
                If Not con Is Nothing Then con.Close()
                If Not dr Is Nothing Then dr.Close()
                con.Dispose()
                con = Nothing
                cmd.Dispose()
                cmd = Nothing
                dr = Nothing
                ReporterWS = Nothing
            Catch ex As Exception
                Diagnostics.EventLog.WriteEntry("ReporterService", "Unable to close database connection " & ex.Message)
                If Debug = "Y" Then mydebuglog.Debug("Unable to close database connection " & ex.Message)
            End Try
            ReporterWS = Nothing
            InProcess = False
        End Try
    End Sub

    Public Shared Sub LogInfo(ByVal sMessage As String)
        ' Write info into the event viewer
        Try
            Dim oEventLog As EventLog = New EventLog("Application")
            If Not Diagnostics.EventLog.SourceExists("ReporterService") Then
                Diagnostics.EventLog.CreateEventSource("ReporterService", "Application")
            End If
            Diagnostics.EventLog.WriteEntry("ReporterService", sMessage, System.Diagnostics.EventLogEntryType.Information)
        Catch e As Exception
        End Try
    End Sub

    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Private Function GetNodeValue(ByVal sNodeName As String, ByVal oParentNode As XmlNode) As String
        ' Generic function to return the value of a node in an XML document
        Dim oNode As XmlNode = oParentNode.SelectSingleNode(".//" + sNodeName)
        If oNode Is Nothing Then
            Return String.Empty
        Else
            Return oNode.InnerText
        End If
    End Function

    Function ChkString(ByVal Instring As String) As String
        ' Generic function to create a string that can be used in a SQL INSERT statement
        Dim temp, outstring As String
        Dim i As Integer
        temp = Instring
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        ChkString = outstring
    End Function

    Function ClearString(ByVal Instring As String) As String
        ' Generic function to remove CRLFs from a string
        Dim temp, outstring, tstr As String
        Dim i As Integer
        temp = Instring
        outstring = ""
        For i = 1 To Len(temp$)
            tstr = Mid(temp, i, 1)
            Select Case tstr
                Case Chr(9)
                    outstring = outstring
                Case Chr(10)
                    outstring = outstring
                Case Chr(13)
                    outstring = outstring
                Case Else
                    outstring = outstring & tstr
            End Select
        Next
        ClearString = outstring
    End Function

    Public Function CheckDBNull(ByVal obj As Object, Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
        ' Checks an object to determine if its null, and if so sets it to a not-null empty value
        Dim objReturn As Object
        objReturn = obj
        If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
            objReturn = ""
        ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
            objReturn = 0
        ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
            objReturn = 0.0
        End If
        Return objReturn
    End Function

    Function CheckNull(ByVal Instring As String) As String
        ' Check to see if a string is null
        If Instring Is Nothing Then
            CheckNull = ""
        Else
            CheckNull = Instring
        End If
    End Function

    ' =================================================
    ' PRINT FUNCTIONS
    Private Function UseLvbPrint(ByVal oPrinter As String, fileName As String, portrait As Boolean, sTray As String) As String

        Dim lvbArguments As String
        Dim lvbProcessInfo As ProcessStartInfo
        Dim lvbProcess As Process

        Try
            If portrait Then
                ' -I ""{2}""
                lvbArguments = String.Format(" -P ""{0}"" -O Port -F ""{1}"" ", oPrinter, fileName)
            Else
                lvbArguments = String.Format(" -P ""{0}"" -O Land -F ""{1}"" ", oPrinter, fileName)
            End If
            If Debug = "Y" Then mydebuglog.Debug("   > lvbArguments: " & lvbArguments)

            lvbProcessInfo = New ProcessStartInfo()
            lvbProcessInfo.WindowStyle = ProcessWindowStyle.Hidden

            ' location of gsbatchprintc.exe
            lvbProcessInfo.FileName = "C:\gsbatchprint64\gsbatchprintc.exe "
            lvbProcessInfo.Arguments = lvbArguments
            lvbProcessInfo.UseShellExecute = False
            lvbProcessInfo.RedirectStandardOutput = True
            lvbProcessInfo.RedirectStandardError = True
            lvbProcessInfo.CreateNoWindow = True
            Try
                lvbProcess = Process.Start(lvbProcessInfo)
            Catch ex As Exception
                Return "Error"
            End Try

            '
            ' Read in all the text from the process with the StreamReader.
            '
            Using reader As StreamReader = lvbProcess.StandardOutput
                Dim result As String = reader.ReadToEnd()
            End Using

            Using readerErr As StreamReader = lvbProcess.StandardError
                Dim resultErr As String = readerErr.ReadToEnd()
                If resultErr.Trim() > "" Then
                    lvbProcess.Close()
                    Return resultErr
                End If
            End Using

            If lvbProcess.HasExited = False Then
                lvbProcess.WaitForExit(3000)
            End If

            lvbProcess.Close()

            Return ""

        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function

    ' Class for reading and writing the Windows Registry 
    ' overcoming the restrictions imposed by 
    ' GetSetting y SaveSetting, which only allow you to 
    ' read and write from  HKEY_CURRENT_USER\Software\VB and VBA
    ' Programming: Sinhu� B�ez
    ' Date: FEB/25/2003

    Public Class CRegistry

        '---------public properties of the class----------------------
        Public ReadOnly Property HKeyLocalMachine() As RegistryKey
            Get
                Return Registry.LocalMachine
            End Get
        End Property

        Public ReadOnly Property HkeyClassesRoot() As RegistryKey
            Get
                Return Registry.ClassesRoot
            End Get
        End Property

        Public ReadOnly Property HKeyCurrentUser() As RegistryKey
            Get
                Return Registry.CurrentUser
            End Get
        End Property

        Public ReadOnly Property HKeyUsers() As RegistryKey
            Get
                Return Registry.Users
            End Get
        End Property

        Public ReadOnly Property HKeyCurrentConfig() As RegistryKey
            Get
                Return Registry.CurrentConfig
            End Get
        End Property

        '--------------------------------------------------------------------
        'Description: Writes a value in the Registry
        'Parameters: 
        '   ParentKey: a RegistryKey that represents any of the six partent keys
        '              where you want to write.
        '   SubKey: a String with the name of the subkey(or nested subkeys)
        '           where you want to write. This Subkey or subkeys may exist 
        '           or not. If a subkey(or subkeys) doesn't exist 
        '           this method will create it.
        '   ValueName: a String with the name of the value to be created.
        '   Value: an Object with the value to be stored.
        'Returns: True if the method succeded, otherwise False.
        'Date: FEB/25/2003
        'Programming: Sinhu� B�ez
        '--------------------------------------------------------------------
        Public Function WriteValue(ByVal ParentKey As RegistryKey, _
                                   ByVal SubKey As String, _
                                   ByVal ValueName As String, _
                                   ByVal Value As Object) As Boolean

            Dim Key As RegistryKey

            Try
                'Opens the given subkey
                Key = ParentKey.OpenSubKey(SubKey, True)
                If Key Is Nothing Then 'if Key doesn't exist then create it
                    Key = ParentKey.CreateSubKey(SubKey)
                End If

                'sets the value
                Key.SetValue(ValueName, Value)

                Return True
            Catch e As Exception
                Return False
            End Try
        End Function

        '--------------------------------------------------------------------
        'Description: Reads a value from the Registry
        'Parameters: 
        '   ParentKey: a RegistryKey that represents any of the six partent keys
        '              where you want to read.
        '   SubKey: a String with the name of the subkey(or nested subkeys)
        '           where you want to read.
        '   ValueName: a String with the name of the value to be read.
        '   Value: an Object with the value to be read.
        'Returns: True if the method succeded, otherwise False.
        'Date: FEB/25/2003
        'Programming: Sinhu� B�ez
        '--------------------------------------------------------------------
        Public Function ReadValue(ByVal ParentKey As RegistryKey, _
                                  ByVal SubKey As String, _
                                  ByVal ValueName As String, _
                                  ByRef Value As Object) As Boolean
            Dim Key As RegistryKey

            Try
                'opens the given subkey
                Key = ParentKey.OpenSubKey(SubKey, True)
                If Key Is Nothing Then 'it Key doesn't exist then throw an exception
                    Throw New Exception("The key does not exist")
                End If

                'Gets the value
                Value = Key.GetValue(ValueName)

                Return True
            Catch e As Exception
                Return False
            End Try
        End Function

        '--------------------------------------------------------------------
        'Description: Deletes a value from the Registry
        'Parameters: 
        '   ParentKey: a RegistryKey that represents any of the six partent keys
        '              where you want to delete a value from.
        '   SubKey: a String with the name of the subkey(or nested subkeys)
        '           where you want to delete a value from.
        '   ValueName:  a String with the name of the value to be deleted.
        'Returns: True if the method succeded, otherwise False.
        'Date: MAR/3/2003
        'Programming: Sinhu� B�ez
        '--------------------------------------------------------------------
        Public Function DeleteValue(ByVal ParentKey As RegistryKey, _
                                    ByVal SubKey As String, _
                                    ByVal ValueName As String) As Boolean
            Dim Key As RegistryKey

            Try
                'opens the given subkey
                Key = ParentKey.OpenSubKey(SubKey, True)

                'deletes the value
                If Not Key Is Nothing Then
                    Key.DeleteValue(ValueName)
                    Return True
                Else
                    Return False
                End If
            Catch e As Exception
                Return False
            End Try

        End Function

        '--------------------------------------------------------------------
        'Description: Deletes a key from the Registry
        'Parameters: 
        '   ParentKey: a RegistryKey that represents any of the six partent keys
        '              where you want to delete a subkey from.
        '   SubKey: a String with the name of the subkey to be deleted.
        'Returns: True if the method succeded, otherwise False.
        'Date: MAR/3/2003
        'Programming: Sinhu� B�ez
        '--------------------------------------------------------------------
        Public Function DeleteSubKey(ByVal ParentKey As RegistryKey, _
                                    ByVal SubKey As String) As Boolean

            Try
                'deletes the subkey and returns a False if Subkey doesn't exist
                ParentKey.DeleteSubKey(SubKey, False)
                Return True
            Catch e As Exception
                Return False
            End Try
        End Function

        '--------------------------------------------------------------------
        'Description: Create a registry key
        'Parameters: 
        '   ParentKey: a RegistryKey that represents any of the six partent keys
        '              where you want to create a subkey.
        '   SubKey:  a String with the name of the subkey to be created.
        'Returns: True if the method succeded, otherwise False.
        'Date: MAR/3/2003
        'Programming: Sinhu� B�ez
        '--------------------------------------------------------------------
        Public Function CreateSubKey(ByVal ParentKey As RegistryKey, _
                                   ByVal SubKey As String) As Boolean

            Try
                'creates the given subkey
                ParentKey.CreateSubKey(SubKey)
                Return True
            Catch e As Exception
                Return False
            End Try
        End Function

        Public Function GetAllSubKeys(ByVal ParentKey As RegistryKey, _
                                         ByVal SubKey As String, _
                                         ByRef strSubKeys() As String) As Boolean
            Dim Key As RegistryKey

            Try
                'opens the given subkey
                Key = ParentKey.OpenSubKey(SubKey)
                If Not Key Is Nothing Then
                    'get all the subKeys (child subkeys)
                    strSubKeys = Key.GetSubKeyNames
                    Return True
                Else
                    Return False
                End If
            Catch e As Exception
                Return False
            End Try
        End Function
    End Class

    ' =================================================
    ' HTTP PROXY CLASS
    Class simplehttp
        Public Function geturl(ByVal url As String, ByVal proxyip As String, ByVal port As Integer, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Referer = ""
            req.Headers.[Set]("Accept-Language", "en,en-us")
            Dim stream_in As StreamReader

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don�t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim response As String = ""
            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                stream_in = New StreamReader(resp.GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function

        Public Function getposturl(ByVal url As String, ByVal postdata As String, ByVal proxyip As String, ByVal port As Short, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.ContentLength = postdata.Length
            req.Referer = ""

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don�t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim stream_out As New StreamWriter(req.GetRequestStream(), System.Text.Encoding.ASCII)
            stream_out.Write(postdata)
            stream_out.Close()
            Dim response As String = ""

            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                Dim resStream As Stream = resp.GetResponseStream()
                Dim stream_in As New StreamReader(req.GetResponse().GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function
    End Class

End Class
