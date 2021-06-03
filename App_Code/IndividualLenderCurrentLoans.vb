Imports System.Data.SqlClient
Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Globalization
Imports System.Linq
Imports System.Text
Public Class IndividualLenderCurrentLoans
    Public Shared _IndividualLenderLastLentDetails As String
    Public Shared _IndividualLenderLastLentDetailsLoanID As String
    Public Shared _IndividualLenderLastLentDetailsLoanName As String
    Public Shared _IndividualLenderLastLentDetailsLoanDateTime As String
    Public Shared _IndividualLenderLastLentDetailsNextMaturityDateTime As String
    Public Shared _IndividualLenderLastLentDetailsLoanAmount As String
    Public Shared _IndividualLenderLastLentDetailsEarned As String
    Public Shared _IndividualLenderCurrentLoansDataTable As DataTable
    Public Shared _IndividualLenderCurrentLoansHasRows As Boolean


    Public Shared Function GetIndividualLenderCurrentLoansDT(ByVal searchValueFromSelectedByUser As String)
        Dim dt As New DataTable()
        _IndividualLenderLastLentDetailsLoanID = String.Empty
        _IndividualLenderLastLentDetailsLoanName = String.Empty
        _IndividualLenderLastLentDetailsLoanDateTime = String.Empty
        _IndividualLenderLastLentDetailsNextMaturityDateTime = String.Empty
        _IndividualLenderLastLentDetailsLoanAmount = "£0.00"


        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Using cmd As New SqlCommand("GetIndividualLenderCurrentLoansByAccountID")
                    With cmd.Parameters
                        .Add(New SqlParameter("@AccountID", searchValueFromSelectedByUser))
                    End With
                    cmd.Connection = con
                    cmd.CommandTimeout = 0
                    cmd.CommandType = CommandType.StoredProcedure
                    Using sda As New SqlDataAdapter(cmd)
                        sda.Fill(dt)
                    End Using
                End Using
            Catch ex As Exception
            Finally

            End Try
        End Using



        If dt.Rows.Count > 0 Then
            _IndividualLenderCurrentLoansDataTable = dt
            _IndividualLenderCurrentLoansHasRows = True
            Dim lastLentDetails = dt.AsEnumerable()
            Dim lastLentLoanAmount = lastLentDetails.Max(Function(l) l.Field(Of DateTime)("StartDate"))
            _IndividualLenderLastLentDetailsLoanAmount = lastLentLoanAmount.ToString()
            Dim lastLentLoanRef As Integer = lastLentDetails.Max(Function(l) l.Field(Of Integer)("TheLoan"))
            _IndividualLenderLastLentDetailsLoanID = lastLentLoanRef.ToString()

            Dim lastLentLoanDate = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of DateTime)("StartDate")).FirstOrDefault().ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture)
            _IndividualLenderLastLentDetailsLoanDateTime = lastLentLoanDate.ToString()

            Dim lastLentLoanName = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of String)("CompanyName")).FirstOrDefault()
            _IndividualLenderLastLentDetailsLoanName = lastLentLoanName.ToString()
            Dim lastLentAmount = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of Double)("outstanding")).FirstOrDefault()
            _IndividualLenderLastLentDetailsLoanAmount = GenDB.PenceToCurrencyStringPounds(lastLentAmount * 100)


            Dim LoanId = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of Integer)("TheLoan")).FirstOrDefault()

            Dim therate = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of Double)("therate")).FirstOrDefault()


            Dim lastLoanDate = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of DateTime)("LastDate")).FirstOrDefault().ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture)

            Dim nextMaturity = (From c In lastLentDetails Order By c.Field(Of DateTime)("LastDate") Where c.Field(Of DateTime)("LastDate") >= DateTime.Today Select c.Field(Of DateTime)("LastDate")).FirstOrDefault().ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture)
            nextMaturity = If(nextMaturity.Contains("01/01/0001 00:00:00"), "N/A see Current Loans Tab", nextMaturity)
            _IndividualLenderLastLentDetailsNextMaturityDateTime = nextMaturity.ToString()


            Dim LoanType = (From c In lastLentDetails Where c.Field(Of Integer)("TheLoan") = lastLentLoanRef Select c.Field(Of Integer)("LoanType")).FirstOrDefault()


        End If
        Return dt
    End Function
End Class
