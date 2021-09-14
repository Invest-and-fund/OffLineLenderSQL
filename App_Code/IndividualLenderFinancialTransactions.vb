Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Data.SqlClient

Public Class IndividualLenderFinancialTransactions
    Public Shared _IndividualLenderFinancialTransactionsDataTable As DataTable
    Public Shared _IndividualLenderFinancialTransactionsHasRows As Boolean
    Private Const STR_SystemString As String = "System.String"
    Public Shared _IndividualLenderTransactionDateFrom As String
    Public Shared _IndividualLenderTransactionDateTo As String
    Public Shared _IndividualLenderTransactionsReportDefault As Boolean
    Public Shared _IndividualLenderTransactionsCount As Integer


    Public Shared Function GetLoanName(ByVal dr As DataRow) As String
        Dim arr1ToGetLoanDesc As Integer() = {1101, 1103, 1200, 1201, 1205, 1206, 1207, 1208, 1209, 1210, 1300, 1301, 1302, 1303, 1304, 1305, 1306, 1307, 1308}
        Dim arr2ToGetLoanDescPlusBidId As Integer() = {1211, 1212, 1213, 1214}
        Dim arr3ToGetLoanDescPlusFTOrderId As Integer() = {1400, 1401, 1402, 1403, 1404, 1405, 1406, 1407, 1408, 1409, 1410, 1411, 1412, 1413}
        Dim arr4ToGetEmptyLoanDesc As Integer() = {1100, 1102}
        Dim intIndexOfTransTypeInLoanDescription As Integer = 0
        Dim intTransType As Integer = 0
        Dim strLoanName As String = Nothing
        Dim sLoanDesc As String = Convert.ToString(GenDB.fnDBStringField(dr("LoanDesc"))).Trim()

        If String.IsNullOrEmpty(sLoanDesc) Then
            sLoanDesc = Convert.ToString(GenDB.fnDBStringField(dr("LoanDesc2"))).Trim()
        End If

        intTransType = GenDB.fnDBIntField(dr("TransType"))
        strLoanName = String.Empty
        intIndexOfTransTypeInLoanDescription = Array.IndexOf(arr1ToGetLoanDesc, intTransType)

        If intIndexOfTransTypeInLoanDescription > -1 Then
            strLoanName = sLoanDesc
        End If

        If String.IsNullOrEmpty(strLoanName) Then
            intIndexOfTransTypeInLoanDescription = Array.IndexOf(arr2ToGetLoanDescPlusBidId, intTransType)

            If intIndexOfTransTypeInLoanDescription > -1 Then
                strLoanName = String.Format("{0} Ref: {1}", sLoanDesc, GenDB.fnDBStringField(dr("bidid")))
            End If
        End If

        If String.IsNullOrEmpty(strLoanName) Then
            intIndexOfTransTypeInLoanDescription = Array.IndexOf(arr3ToGetLoanDescPlusFTOrderId, intTransType)

            If intIndexOfTransTypeInLoanDescription > -1 Then
                strLoanName = String.Format("{0} Ref: {1}", sLoanDesc, GenDB.fnDBStringField(dr("ftorderid")))
            End If
        End If

        intIndexOfTransTypeInLoanDescription = Array.IndexOf(arr4ToGetEmptyLoanDesc, intTransType)

        If intIndexOfTransTypeInLoanDescription > -1 Then
            strLoanName = String.Empty
        End If

        Return strLoanName.Trim()
    End Function

    Public Shared Function GetIndividualLenderAnnualStatementData() As DataSet

        Dim strLoanName As String = Nothing
        Dim strLastDesc As String = Nothing
        Dim ds As DataSet = New DataSet()
        Dim dr As DataRow = Nothing
        Dim dColTransDesc As DataColumn = New DataColumn("TransDesc", Type.[GetType](STR_SystemString))
        Dim dColAmountDeposit As DataColumn = New DataColumn("AmountDeposit", Type.[GetType](STR_SystemString))
        Dim dColTheActualBalance As DataColumn = New DataColumn("TheActualBalance", Type.[GetType](STR_SystemString))
        Dim dColAmountWithdraw As DataColumn = New DataColumn("AmountWithdraw", Type.[GetType](STR_SystemString))
        Dim dColStatementBalanceAmount As DataColumn = New DataColumn("StatementBalanceAmount", Type.[GetType](STR_SystemString))
        Dim intCounter As Integer = 0
        Dim intAmount As Integer = 0
        Dim intActualBalance As Integer = 0
        Dim intTransType As Integer = 0
        Dim fAmount As Double = 0
        Dim strAmount As String = Nothing
        Dim strActualBalance As String = Nothing
        Dim intNegativeTransTypes As Integer() = {1001, 1002, 1024, 1102, 1103, 1200, 1204, 1206, 1209, 1212, 1214, 1300, 1302, 1304, 1306, 1401, 1402, 1405, 1406, 1408, 1410, 1412}



        Dim sSQL As String
        _IndividualLenderFinancialTransactionsHasRows = False
        _IndividualLenderTransactionsCount = 0
        _IndividualLenderFinancialTransactionsDataTable = Nothing

        Try
            sSQL = "SELECT ft.fin_transid,  
                            ft.datecreated       AS thedate,  
                            ft.transtype,  
                            ft.header_id,  
                            ft.orderid           AS ftorderid,  
                            coalesce (Trim(ln.business_name) , Trim(ln1.business_name) )    AS LoanDesc,  
                            Trim(lh.business_name)     AS LoanDesc2,  
                             (ft.amount)            AS amount1,  
                             (ft.balance_available) AS TheBalance1,  
                             (ft.balance_actual)    AS TheActualBalance1,  
                            bidid,  
                             
                            o.lh_id,
							coalesce (o.loanid, lc.loanid) as loanid
							, lc.LOANID as lcloanid
							,ft.accountid
                     FROM   fin_trans ft  
                            LEFT OUTER JOIN fin_bals fb  
                                         ON ft.fin_transid = fb.fin_transid  
                            LEFT OUTER JOIN fin_bals_suspense fbs  
                                         ON ft.fin_transid = fbs.fin_transid  
                            LEFT OUTER JOIN orders o  
                                         ON ft.orderid = o.orderid  
                            LEFT OUTER JOIN loans ln  
                                         ON ln.loanid = o.loanid  
                            LEFT OUTER JOIN bids bd  
                                         ON bd.orderid = o.orderid  
                            LEFT OUTER JOIN lh_id_loan lh   
                                         ON lh.lh_id = o.lh_id 
							LEFT OUTER JOIN LOANCO_LH_BALS lC 
                                         ON  ft.amount = lc.amount  and FORMAT(ft.datecreated, 'yyyy-MM-dd HH') = FORMAT(lc.datecreated, 'yyyy-MM-dd HH')
							LEFT OUTER JOIN loans ln1  
                                         ON ln1.loanid = lc.loanid
                     WHERE  ft.accountid = 3711
                            AND ft.isactive = 0  
                            AND ft.transtype NOT IN ( 1200, 1201, 1203, 1204,  
                                                      1211, 1212, 1213, 1214 )"


            ''''''''''''''Need to check
            'If Not _IndividualLenderTransactionsReportDefault Then
            '    sSQL = sSQL & " AND ft.datecreated between @D_StartDate and @I_EndDate"
            'End If

            sSQL = sSQL & "ORDER  BY ft.datecreated DESC, ft.fin_transid DESC"
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                    Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                    With cmd.Parameters
                        .Add(New SqlParameter("@I_ACCOUNT_ID", IndividualLenderProfileDetails._LenderAccountId))
                        '.Add(New SqlParameter("@D_StartDate", IndividualLenderFinancialTransactions._IndividualLenderTransactionDateFrom))
                        '.Add(New SqlParameter("@I_EndDate", IndividualLenderFinancialTransactions._IndividualLenderTransactionDateTo))
                    End With
                    con.Open()

                    adapter.SelectCommand = cmd

                    adapter.Fill(ds)


                Catch ex As Exception
                Finally

                End Try
            End Using


            Dim table As DataTable = New DataTable()
            table = ds.Tables(0)
            intCounter = table.Rows.Count
            Dim t3 = Task.Run(Sub() table.Columns.Add(dColTransDesc))
            t3.Wait()
            Dim t4 = Task.Run(Sub() table.Columns.Add(dColAmountDeposit))
            t4.Wait()
            Dim t5 = Task.Run(Sub() table.Columns.Add(dColTheActualBalance))
            t5.Wait()
            Dim t6 = Task.Run(Sub() table.Columns.Add(dColAmountWithdraw))
            t6.Wait()
            Dim t7 = Task.Run(Sub() table.Columns.Add(dColStatementBalanceAmount))
            t7.Wait()
            strLastDesc = String.Empty
            _IndividualLenderTransactionsCount = intCounter

            If intCounter > 0 Then

                If _IndividualLenderTransactionsReportDefault Then
                    Dim t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function() table.AsEnumerable())
                    t.Wait()
                    Dim t1 = Task(Of DateTime).Factory.StartNew(Function() t.Result.Min(Function(l) l.Field(Of DateTime)("thedate")))
                    t1.Wait()
                    Dim t2 = Task(Of DateTime).Factory.StartNew(Function() t.Result.Max(Function(l) l.Field(Of DateTime)("thedate")))
                    t2.Wait()
                    _IndividualLenderTransactionDateFrom = t1.Result.ToString("dd/MM/yyyy")
                    _IndividualLenderTransactionDateTo = t2.Result.ToString("dd/MM/yyyy")
                End If

                For i = 0 To intCounter - 1
                    _IndividualLenderFinancialTransactionsHasRows = True
                    Dim drCurrentRow = table.Rows(i)
                    dr = drCurrentRow
                    intTransType = GenDB.fnDBIntField(dr("TransType"))
                    strLoanName = GetLoanName(dr)
                    intAmount = GenDB.fnDBIntField(dr("amount1"))
                    intActualBalance = GenDB.fnDBIntField(dr("TheActualBalance1"))
                    strActualBalance = GenDB.PenceToCurrencyStringPounds(intActualBalance)
                    fAmount = intAmount / 100.0R

                    If fAmount < 0 Then
                        fAmount = fAmount * -1
                    End If

                    strAmount = GenDB.PenceToCurrencyStringPounds(intAmount)

                    If intNegativeTransTypes.Contains(intTransType) Then
                        dr("amountwithdraw") = String.Format("({0})", strAmount)
                        dr("statementbalanceamount") = Convert.ToString(Convert.ToDecimal(strActualBalance.Replace("£", String.Empty)) + Convert.ToDecimal(strAmount.Replace("£", String.Empty)))
                    Else
                        dr("amountdeposit") = strAmount
                        dr("statementbalanceamount") = Convert.ToString(Convert.ToDecimal(strActualBalance.Replace("£", String.Empty)) - Convert.ToDecimal(strAmount.Replace("£", String.Empty)))
                    End If

                    dr("TheActualBalance") = strActualBalance

                    If (intTransType <> 1100) And (intTransType <> 1102) Then

                        If String.IsNullOrEmpty(strLoanName) Then
                            strLoanName = strLastDesc
                        End If
                    End If

                    strLastDesc = strLoanName
                    drCurrentRow("TransDesc") = GenDB.GetFinTrans(intTransType)
                    drCurrentRow("LoanDesc") = strLoanName

                    If intTransType = 1100 Then
                        Dim intFinTransID As Integer = Convert.ToInt32(drCurrentRow("Fin_transid"))

                        If (intFinTransID >= 11889) AndAlso (intFinTransID <= 11913) Then
                            drCurrentRow("TransDesc") = "Fee adjustment"
                        End If
                    End If
                Next

                table.AcceptChanges()
            End If

            _IndividualLenderFinancialTransactionsDataTable = ds.Tables(0)
            Return ds



        Catch ex As Exception

            Throw
        Finally

        End Try

        Return Nothing

    End Function

    Function ExportToCsv() As StringBuilder
        'If Not _IndividualLenderFinancialTransactionsDataTable.Equals(Nothing) Then
        '    Dim table As DataTable = _IndividualLenderFinancialTransactionsDataTable
        '    Return IAFDataManager.Provider.ExportDataTableToCSV(table, String.Empty, String.Empty, False)
        'End If

        'Return Nothing
    End Function
End Class
