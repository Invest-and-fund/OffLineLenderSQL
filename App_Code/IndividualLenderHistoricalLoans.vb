Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class IndividualLenderHistoricalLoans
    Public Shared _IndividualLenderHistoricalLoansDataTable As DataTable
    Public Shared _IndividualLenderHistoricalLoansHasRows As Boolean

    Public Shared Function GetIndividualLenderHistoricalLoans(ByVal accountIdSelectedByUser As String) As DataTable

        Dim result As DataTable = New DataTable()

        Dim sSQL As String

        Dim strSqlIndividualLenderHistoricalLoans As StringBuilder = New StringBuilder()

        Try
            If Not String.IsNullOrWhiteSpace(accountIdSelectedByUser) Then
                sSQL = "SELECT l.loanid                     AS The_Loan,  
                          Trim(l.business_name)              AS Company_Name,  
                           convert(date,(date_of_last_payment) )          AS dd_lastdate,  
                          Cast(Sum(f.amount) AS FLOAT)/100 AS Interest_Received,  
                          Trim(l.nature_of_business)         AS Description  
              FROM   loans l,  
                          lh_id_loan li,  
                          fin_trans f,  
                          orders o  
                  WHERE  f.accountid =  @I_ACCOUNT_ID  
                          AND f.orderid = o.orderid  
                          AND f.transtype IN ( 1303, 1308 )  
                          AND o.lh_id = li.lh_id  
                          AND l.loanid = li.loanid  
                   GROUP  BY l.business_name,  
                             l.loanid,  
                             l.nature_of_business,  
                             l.date_of_last_payment  
                    ORDER  BY dd_lastdate DESC "

                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        With cmd.Parameters
                            .Add(New SqlParameter("@I_ACCOUNT_ID", IndividualLenderProfileDetails._LenderAccountId))
                        End With
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using



                _IndividualLenderHistoricalLoansHasRows = False

                If result.Rows.Count > 0 Then
                    _IndividualLenderHistoricalLoansDataTable = result
                    _IndividualLenderHistoricalLoansHasRows = True
                End If
            Else
                result = Nothing
            End If



        Catch ex As Exception

            Throw
        Finally



            strSqlIndividualLenderHistoricalLoans.Length = 0
        End Try

        Return result

    End Function

    'Function ExportToCsv() As StringBuilder
    '    If Not _IndividualLenderHistoricalLoansDataTable.Equals(Nothing) Then
    '        Dim table As DataTable = _IndividualLenderHistoricalLoansDataTable
    '        Return IAFDataManager.Provider.ExportDataTableToCSV(table, String.Empty, String.Empty, False)
    '    End If

    '    Return Nothing
    'End Function

End Class
