
Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Text
Imports System.Data.SqlClient

Public Class IndividualLenderTradingHistory
    Public Shared _IndividualLenderTradingHistoryLastLentDate As DateTime
    Public Shared _IndividualLenderTradingHistoryDataTable As DataTable
    Public Shared _IndividualLenderTradingHistorysHasRows As Boolean


    Public Shared Function GetIndividualLenderTradingHistory(ByVal accountIdSelectedByUser As String) As DataTable

        Dim result As DataTable = New DataTable()


        Dim strSqlIndividualLenderTradingHistory As StringBuilder = New StringBuilder()

        Try

            If Not String.IsNullOrWhiteSpace(accountIdSelectedByUser) Then
                Dim sSQL As String
                sSQL = "         SELECT o.ordertransdate as ORDER_TRANS_DATE,  
                         lh.loanid as LOAN_ID,  
                         Trim(Replace(l.trading_name, ',', ' ')) AS Loan_Name,  
                         o.lh_id as LH_ID,  
                         Concat(Trim(u1.lastname)  
                        , ', '  
                         ,Trim(u1.firstname)   
                         ,' ('  
                         , a1.accountid  
                       , ')'   )                         AS THE_BUYER,  
                         Concat(Trim(u2.lastname)  
                         , ', '  
                         , Trim(u2.firstname)  
                         , ' ('  
                         , a2.accountid  
                          , ')'  )                          AS THE_SELLER,  
                         Cast(o.orderamount AS FLOAT) / 100 AS ORDER_AMOUNT,  
                         os.premium as PREMIUM  
                   FROM   ordertrans_accounts o,  
                         users u1,  
                         users u2,  
                         accounts a1,  
                         accounts a2,  
                         orders os,  
                         loan_holdings lh,  
                         loans l  
               WHERE  o.ordertypebuy in (2,3) -- 2 = Buy and 3= Sell  
                         AND a1.userid = u1.userid  
                         AND a1.accountid = o.accountidbuy  
                         AND a2.userid = u2.userid  
                         AND a2.accountid = o.accountidsell  
                         AND os.orderid = o.orderidsell  
                         AND lh.loan_holdings_id = os.lh_id  

                          AND l.loanid = lh.loanid

						  AND  (o.accountidsell  IN  (@I_ACCOUNT_ID )
						     or   o.accountidbuy    IN (@I_ACCOUNT_ID))
                         order by o.ordertransdate desc  "




                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        With cmd.Parameters
                            .Add(New SqlParameter("@I_ACCOUNT_ID", accountIdSelectedByUser))

                        End With
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using



                _IndividualLenderTradingHistorysHasRows = False

                If result.Rows.Count > 0 Then
                    _IndividualLenderTradingHistoryDataTable = result
                    _IndividualLenderTradingHistorysHasRows = True
                End If
            Else
                result = Nothing
            End If



        Catch ex As Exception

            Throw
        Finally



            strSqlIndividualLenderTradingHistory.Length = 0
        End Try

        Return result

    End Function



    Function ExportToCsv() As StringBuilder
        'If Not _IndividualLenderTradingHistoryDataTable.Equals(Nothing) Then
        '    Dim table As DataTable = _IndividualLenderTradingHistoryDataTable
        '    Return IAFDataManager.Provider.ExportDataTableToCSV(table, String.Empty, String.Empty, False)
        'End If

        Return Nothing
    End Function

End Class
