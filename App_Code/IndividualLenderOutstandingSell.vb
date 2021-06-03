
Imports System
Imports System.Data
Imports System.Text
Imports System.Data.SqlClient

Public Class IndividualLenderOutstandingSell

    Public Shared _IndividualLenderOutstandingSellDataTable As DataTable
    Public Shared _IndividualLenderOutstandingSellHasRows As Boolean

    Public Shared Function GetIndividualLenderOutstandingSell(ByVal accountIdSelectedByUser As String) As DataTable

        Dim result As DataTable = New DataTable()

        Dim sSQL As String
        Dim strSqlIndividualLenderOutstandingSell As StringBuilder = New StringBuilder()

        Try

            If Not String.IsNullOrWhiteSpace(accountIdSelectedByUser) Then

                sSQL = "SELECT Trim(l.business_name)  AS COMPANY_NAME,  
                            Trim(l.description)  As DESCRIPTION,  
                            (l.loanid)   AS THE_LOAN,  
                            l.ipo_end  AS IPO_END,
                            Cast(l.maxloanamount As FLOAT) / 100  As MAX_LOAN_AMOUNT,
                            Trim(l.security_guarantees)      As SECURITY_GURANTEES,  
                            (l.term)      AS TERM,  
                            Trim(l.company_risk_as_stars)        As COMPANY_RISK_STARS,  
                            o.orderid               AS THE_ORDER,
                            o.accountid           As ACCOUNT_ID,  
                            Cast(o.principal_outstanding As FLOAT) /100     As OUTSTANDING_AMOUNT,  
                            otp.nummonths    AS MONTHS_LEFT,
                            (Cast(o.amount As FLOAT) / 100)   As SALE_AMOUNT,
                            Cast(o.premium As FLOAT) / 100    As PREMIUM_PERCENT,
                            (Cast(o.premium As FLOAT) / 10000) * (o.principal_outstanding / 100)  As PREMIUM,
                            (Cast(o.principal_outstanding As FLOAT) / 100) + ((Cast(o.premium As FLOAT) / 10000) * (o.principal_outstanding / 100))          As AMOUNT_FULL,
                            Cast(o.rate As FLOAT) / 100         As NEW_RATE,
                            Trim(l.anythingelse)   As ANYTHING_ELSE,  
                            Trim(l.purpose_of_loan)     As PURPOSE_OF_LOAN,  
                            Trim(l.background)      As BACKGROUND,  
                            l.people     AS PEOPLE  
                     From orders o,
                            loans l,
                            loan_holdings lh,
                            outstandingrepayments otp  
                     Where o.ordertype = @orSell
                            And o.status = 0 ---- 0 = open  
                           And o.accountid =   @I_ACCOUNT_ID
                            And lh.loan_holdings_id = o.lh_id  
                            And otp.loanid = lh.loanid  
                            And l.loanid = lh.loanid"

                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        With cmd.Parameters
                            .Add(New SqlParameter("@I_ACCOUNT_ID", IndividualLenderProfileDetails._LenderAccountId))
                            .Add(New SqlParameter("@orSell", GenDB.orSell))
                        End With
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using


                _IndividualLenderOutstandingSellHasRows = False

                If result.Rows.Count > 0 Then
                    _IndividualLenderOutstandingSellDataTable = result
                    _IndividualLenderOutstandingSellHasRows = True
                End If
            Else
                result = Nothing
            End If


        Catch ex As Exception

            Throw
        Finally



            strSqlIndividualLenderOutstandingSell.Length = 0
        End Try

        Return result

    End Function

    'Function ExportToCsv() As StringBuilder
    '    If Not _IndividualLenderOutstandingSellDataTable.Equals(Nothing) Then
    '        Dim table As DataTable = _IndividualLenderOutstandingSellDataTable
    '        Return IAFDataManager.Provider.ExportDataTableToCSV(table, "", "", False)
    '    End If

    '    Return Nothing
    'End Function

End Class
