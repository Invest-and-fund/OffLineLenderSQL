Imports System.Text
Imports System.IO
Imports System.Configuration

Public Class GenFinances

    Public Shared Function GetServerName() As String
        Dim s, strConn As String

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString.ToUpper

        s = ""
        If strConn.Contains("DEV") Then
            s = "Development Server"
            Return s
        End If

        If s = "" Then
            If strConn.Contains("UAT") Then
                s = "UAT Server"
                Return s
            End If
        End If

        If s = "" Then
            If strConn.Contains("TESTING") Then
                s = "Testing Server"
                Return s
            End If
        End If
        If s = "" Then
            If strConn.Contains("MAIN2") Then
                s = "Shadow Server"
                Return s
            End If
        End If

        If s = "" Then
            If strConn.Contains("FINALTEST") Then
                s = "User Test Server"
                Return s
            End If
        End If

        If s = "" Then
            If strConn.Contains("MAIN") Then
                s = "Live Server"
                Return s
            End If
        End If

        GetServerName = s
    End Function

    Public Shared Function LogEntry(ByVal sEntry As String) As Integer
        Dim mydocpath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim sb As StringBuilder = New StringBuilder()
        Dim sFile As String

        sb.Append(Now.ToShortDateString & "=" & Now.ToLongTimeString & "=" & sEntry)
        sb.AppendLine()

        sFile = Directory.GetCurrentDirectory & "\logs\Finances-" & Now.ToString("yyyyMMdd") & ".log"
        Using outfile As StreamWriter = New StreamWriter(sFile, True)
            outfile.Write(sb.ToString())
        End Using

        LogEntry = 0
    End Function



End Class
