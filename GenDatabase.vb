Public Class GenDatabase

    Public Shared dsPreBids As DataSet
    Public Shared dsMandates As DataSet
    Public Shared dsLoans As DataSet
    Public Shared dsFinBals As DataSet
    Public Shared dsLoanSets As DataSet
    Public Shared dsLenders As DataSet
    Public Shared dsHoldings As DataSet
    Public Shared dsListLoanSets As DataSet
    Public Shared dsLoanSetiD As DataSet
    Public Shared dsCurrentBids As DataSet

    Public Shared LatestDate As DateTime
    Public Shared NumPreLoans As Integer

    Public Shared Sub LoanLoans()
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i, iCounter As Integer
        Dim dr As DataRow
        Dim CurDate As DateTime

        sSQL = "select *
                 from loans l, accounts a, users u
                 where a.accountid = l.accountid
                    and l.isactive = 0
                    and l.LoanStatus = 1
                    and maxloanamount > 0
                    and u.userid = a.userid
                 order by LoanID"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(dsLoans)

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing

        sSQL = "select *
                 from loan_sets
                where isactive = 0
                order by loanSetID"

        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(dsLoanSets)

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Sub


    Public Shared Function CurrentBidTotal(LoanID As Integer) As Integer
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim m As Integer
        Dim j As Integer
        Dim k As Integer
        Dim dr As DataRow

        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        dsCurrentBids = New DataSet


        sSQL = "select sum(b.amount) as BidTotal
                from bids b, orders o, loans l
                where l.loanid = @loanid
                 and o.loanid = l.loanid
                 and b.orderid = o.orderid
                 and b.isactive = 0
                 and o.isactive = 0"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@loanid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = LoanID

        dsCurrentBids.Clear()
        Adaptor.Fill(dsCurrentBids)

        Try
            dr = dsCurrentBids.Tables(0).Rows(0)
            j = dsCurrentBids.Tables(0).Rows.Count
            k = GenDB.fnDBIntField(dr("BidTotal"))

        Catch ex As Exception
            k = 0
        End Try

        CurrentBidTotal = k

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing



    End Function

    Public Shared Sub AddMandate(AccountID As Integer, InterestRate As Integer, MinTenor As Integer, MaxTenor As Integer,
                                MaxLDGV As Integer, MaxLTV As Integer, A_Percent As Integer, A_Amount As Integer,
                                 B_Amount As Integer, Concentration As Integer, isactive As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim sSQL As String

        sSQL = "update mandates
               set isactive = 1
              where accountid = @accountid
                 and mandateid in
             (select max(mandateid) from mandates
               group by accountid)"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()
        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("ACCOUNTID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = AccountID
        Cmd.ExecuteNonQuery()

        MyConn.Close()


        sSQL = "insert into Mandates (ACCOUNTID, ISACTIVE, TENOR_MAX, MIN_INTERESTRATE, MAX_LTV, TENOR_MIN, MAX_LGDV, A_PERCENT_OF_DRAWDOWN, 
                                      A_MAX_AMOUNT, B_AMOUNT, CONCENTRATION ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("ACCOUNTID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = AccountID
        Cmd.Parameters.Add("ISACTIVE", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = isactive
        Cmd.Parameters.Add("TENOR_MAX", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxTenor
        Cmd.Parameters.Add("MIN_INTERESTRATE", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = InterestRate
        Cmd.Parameters.Add("MAX_LTV", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxLTV
        Cmd.Parameters.Add("TENOR_MIN", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MinTenor
        Cmd.Parameters.Add("MAX_LGDV", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxLDGV
        Cmd.Parameters.Add("A_PERCENT_OF_DRAWDOWN", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = A_Percent
        Cmd.Parameters.Add("A_MAX_AMOUNT", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = A_Amount
        Cmd.Parameters.Add("B_AMOUNT", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = B_Amount
        Cmd.Parameters.Add("CONCENTRATION", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = Concentration
       

        Cmd.ExecuteNonQuery()

        MyConn.Close()
        Cmd = Nothing
        MyConn = Nothing
    End Sub

    'Public Shared Sub UpdateMandate(MandateID As Integer, InterestRate As Integer, MinTenor As Integer,
    '                                 MaxTenor As Integer, MaxLDGV As Integer, MaxLTV As Integer,
    '                                 A_Percent As Integer, A_Amount As Integer, B_Amount As Integer,
    '                                 Concentration As Integer)
    '    Dim strConn As String
    '    Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
    '    Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
    '    Dim sSQL As String
    '    Dim i As Integer

    '    sSQL = "update Mandates 
    '            set TENOR_MIN = @1,
    '                MIN_INTERESTRATE = @2,
    '                TENOR_MAX = @3,
    '                MAX_LTV = @4,
    '                MAX_LGDV = @5,
    '                A_PERCENT_OF_DRAWDOWN = @6,
    '                A_MAX_AMOUNT = @7,
    '                B_AMOUNT = @8,
    '                CONCENTRATION = @9 
    '            where MandateID = @mid"

    '    strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
    '    MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

    '    MyConn.Open()

    '    Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
    '    Cmd.Connection = MyConn
    '    Cmd.CommandText = sSQL
    '    Cmd.Parameters.Clear()
    '    Cmd.Parameters.Add("1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MinTenor
    '    Cmd.Parameters.Add("2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = InterestRate
    '    Cmd.Parameters.Add("3", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxTenor
    '    Cmd.Parameters.Add("4", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxLTV
    '    Cmd.Parameters.Add("5", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MaxLDGV
    '    Cmd.Parameters.Add("6", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = A_Percent
    '    Cmd.Parameters.Add("7", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = A_Amount
    '    Cmd.Parameters.Add("8", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = B_Amount
    '    Cmd.Parameters.Add("9", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = Concentration
    '    Cmd.Parameters.Add("mid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MandateID

    '    Try
    '        i = Cmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '        i = -1
    '    End Try

    '    MyConn.Close()
    '    Cmd = Nothing
    '    MyConn = Nothing
    'End Sub

    'Public Shared Sub DeactivateMandate(MandateID As Integer)
    '    Dim strConn As String
    '    Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
    '    Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
    '    Dim sSQL As String
    '    Dim i As Integer

    '    sSQL = "update Mandates 
    '            set ISACTIVE = 1
    '            where MandateID = @mid"

    '    strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
    '    MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

    '    MyConn.Open()

    '    Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
    '    Cmd.Connection = MyConn
    '    Cmd.CommandText = sSQL
    '    Cmd.Parameters.Clear()
    '    Cmd.Parameters.Add("mid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = MandateID

    '    Try
    '        i = Cmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '        i = -1
    '    End Try

    '    MyConn.Close()
    '    Cmd = Nothing
    '    MyConn = Nothing
    'End Sub

    Public Shared Sub AddPreBid(AccountID As Integer, LoanID As Integer, Amount As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim sSQL As String

        sSQL = "insert into PreBids (ACCOUNTID, LOANID, AMOUNT) VALUES (?, ?, ?)"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("ACCOUNTID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = AccountID
        Cmd.Parameters.Add("LOANID", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = LoanID
        Cmd.Parameters.Add("AMOUNT", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = Amount

        Cmd.ExecuteNonQuery()

        MyConn.Close()
        Cmd = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Sub UpdatePreBids(PreBidID As Integer, Amount As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim sSQL As String
        Dim i As Integer

        sSQL = "update Prebids
                set AMOUNT = @1
                where PrebidID = @mid"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = Amount
        Cmd.Parameters.Add("mid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = PreBidID

        Try
            i = Cmd.ExecuteNonQuery()
        Catch ex As Exception
            i = -1
        End Try

        MyConn.Close()
        Cmd = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Sub DeletePreBids(PreBidID As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand
        Dim sSQL As String
        Dim i As Integer

        sSQL = "update Prebids
                set isActive  = 1
                where PrebidID = @mid"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("mid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = PreBidID

        Try
            i = Cmd.ExecuteNonQuery()
        Catch ex As Exception
            i = -1
        End Try

        MyConn.Close()
        Cmd = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Sub LoadPreLoans()
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i, iCounter As Integer
        Dim dr As DataRow
        Dim CurDate As DateTime

        Dim dateNow As DateTime = DateTime.Now

        'dateNow = dateNow.AddDays(6)

        Dim dateBlockIfNotReg As DateTime = New DateTime(2019, 12, 9, 0, 0, 1)



        sSQL = "select m.*, (trim(u.Firstname) || ' ' || trim(u.Lastname)) as Username, t.description as accountType, t.account_Type_ID as accounttypeid
                 from mandates m, Users u, Accounts  a, account_types t
                 where m.isactive = 0
                   and a.AccountID = m.AccountID
                   and u.UserID = a.UserID
                   and a.accounttype = t.account_type_id 
                   and u.activated = 5
                   and a.activated_bank = 5
                   and u.activated_cert = 5
                   and current_timestamp < DATEADD(year, 1, u.client_categorisation_date)"


        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        dsMandates.Clear()
        Adaptor.Fill(dsMandates)

        ' Get latest updated date

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing


        sSQL = "select p.*, (trim(u.Firstname) || ' ' || trim(u.Lastname)) as Username, t.description as accountType, t.account_Type_ID as accounttypeid
                 from prebids p, Users u, Accounts  a, account_types t
                where p.isactive = 0
                   and a.AccountID = p.AccountID
                   and u.UserID = a.UserID
                   and a.accounttype = t.account_type_id
                   and u.activated = 5
                   and a.activated_bank = 5
                   and u.activated_cert = 5
                   and current_timestamp < DATEADD(year, 1, u.client_categorisation_date)"




        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        dsPreBids.Clear()
        Adaptor.Fill(dsPreBids)

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Function UpdateLoans(dt As DateTime)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim ds As dataset
        Dim i, j, j1, a, b As Integer
        Dim bFound As Boolean
        Dim dr, dr1 As DataRow
        Dim iCounter, iLoanCounter As Integer

        sSQL = "select *
                 from loans l
                 where LastUpdated > @p1"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)
        Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.TimeStamp).Value = LatestDate

        ds = New dataset

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(ds)

        iCounter = ds.Tables(0).Rows.Count
        iLoanCounter = dsLoans.Tables(0).Rows.Count

        For i = 0 To iCounter - 1
            dr = ds.Tables(0).Rows(i)
            j = GenDB.fnDBIntField(dr("LoanID"))

            a = 0
            bFound = False
            While (a < iLoanCounter) And (Not bFound)
                dr1 = dsLoans.Tables(0).Rows(a)
                j1 = GenDB.fnDBIntField(dr1("LoanID"))

                If j = j1 Then
                    bFound = True
                End If
                a += 1
            End While

            If bFound Then
                ' dsLoans.Tables(0).Rows(a). = dr1
            End If
        Next i


        ds = Nothing
        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Function

    Public Shared Sub LoadFinBals()
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i, iCounter As Integer
        Dim dr As DataRow
        Dim CurDate As DateTime

        sSQL = "select *
                 from fin_balances"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(dsFinBals)

        ' Get latest updated date

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Sub LoadLenders()
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i, iCounter As Integer
        Dim dr As DataRow
        Dim CurDate As DateTime

        Dim dateNow As DateTime = DateTime.Now

        'dateNow = dateNow.AddDays(6)

        Dim dateBlockIfNotReg As DateTime = New DateTime(2019, 12, 9, 0, 0, 1)

        sSQL = "select trim(Firstname) as Firstname, trim(Lastname) as Lastname, trim(u.Companyname) as CompanyName, AccountID, t.description as AccountType 
                 from users u, accounts a, account_Types t 
                 where a.UserID = u.UserID
                   and u.UserType = 0
                   and u.isactive = 0
                   and u.activated = 5 
                   and a.activated_bank = 5
                   and a.accounttype = t.account_type_id
                 order by Lastname, Firstname"

        If dateNow > dateBlockIfNotReg Then
            sSQL = "select trim(Firstname) as Firstname, trim(Lastname) as Lastname, trim(u.Companyname) as CompanyName, AccountID, t.description as AccountType 
                 from users u, accounts a, account_Types t 
                 where a.UserID = u.UserID
                   and u.UserType = 0
                   and u.isactive = 0
                   and u.activated = 5 
                   and a.activated_bank = 5
                   and u.activated_cert = 5 
                   and a.accounttype = t.account_type_id
                 order by Lastname, Firstname"
        End If

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.Fill(dsLenders)

        ' Get latest updated date

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Function LoadHoldings(iAccountID As Integer, iLoanID As Integer) As Integer
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i, j, iTotal As Integer
        Dim dr As DataRow
        Dim CurDate As DateTime

        sSQL = "select sum(lh.Num_Units) as TheTotal
                from lh_balances lh
                where lh.accountid = @p1 
                  and lh.lh_id in
                    (select Loan_Holdings_id from Loan_Holdings where LoanID in 
                        (select LoanID from Loans where LoanSetID > 0 and LoanSetID = 
                            (Select LoanSetID from Loans where LoanID = @p2 and LoanSetID > 0)))"


        'this bit of sql will bring back 2 rows - one a total of loan holdings and the second the bids on the current loan for the lender
        'the code will need to be amended to cater for both rows returned as it expects only 1 row.

        'sSQL = "select sum(lh.Num_Units) as TheTotal
        '        from lh_balances lh
        '        where lh.accountid = @p1 
        '          and lh.lh_id in
        '            (select Loan_Holdings_id from Loan_Holdings where LoanID in 
        '                (select LoanID from Loans where LoanSetID > 0 and LoanSetID = 
        '                    (Select LoanSetID from Loans where LoanID = @p2 and LoanSetID > 0)))

        '          union all
        '       select sum(b.amount) as TheTotal
        '        from loan_holdings h, bids b
        '       where h.loanID = @p2
        '        and  b.bidid = h.bidid
        '        and  b.accountid = @p1"


        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)
        Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iAccountID
        Adaptor.SelectCommand.Parameters.Add("@p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iLoanID

        dsHoldings.Clear()
        Adaptor.Fill(dsHoldings)

        Try
            dr = dsHoldings.Tables(0).Rows(0)
            j = dsHoldings.Tables(0).Rows.Count
            j = GenDB.fnDBIntField(dr("TheTotal"))
            iTotal = j
        Catch ex As Exception
            iTotal = 0
        End Try
        ' Get latest updated date

        LoadHoldings = iTotal
        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Function

    Public Shared Function GetFinBal(iAccountID As Integer) As Integer
        Dim i, j, iBal, iCounter As Integer
        Dim dr As DataRow
        Dim bFound As Boolean

        iCounter = dsFinBals.Tables(0).Rows.Count
        iBal = 0
        i = 0
        bFound = False
        While i <= iCounter - 1 And Not bFound
            dr = dsFinBals.Tables(0).Rows(i)
            j = GenDB.fnDBIntField(dr("AccountID"))
            If j = iAccountID Then
                bFound = True
                iBal = GenDB.fnDBIntField(dr("AMOUNT"))
            Else
                i += 1
            End If
        End While

        GetFinBal = iBal
    End Function

    Public Shared Function SetFinBal(iAccountID As Integer, Amount As Integer) As Integer
        Dim i, j, iBal, iCounter As Integer
        Dim dr As DataRow
        Dim bFound As Boolean

        iCounter = dsFinBals.Tables(0).Rows.Count
        iBal = 0
        i = 0
        bFound = False
        While i <= iCounter - 1 And Not bFound
            dr = dsFinBals.Tables(0).Rows(i)
            j = GenDB.fnDBIntField(dr("AccountID"))
            If j = iAccountID Then
                bFound = True
                dsFinBals.Tables(0).Rows(i)("AMOUNT") = Amount
            Else
                i += 1
            End If
        End While

        SetFinBal = Amount
    End Function

    Public Shared Function GotoLoanIDIdx(iLoanID As Integer) As Integer
        Dim i, j, iCounter As Integer
        Dim dr As DataRow
        Dim bFound As Boolean

        iCounter = dsLoans.Tables(0).Rows.Count

        i = 0
        bFound = False
        While i <= iCounter - 1 And Not bFound
            dr = dsLoans.Tables(0).Rows(i)
            j = GenDB.fnDBIntField(dr("LoanID"))
            If j = iLoanID Then
                bFound = True
            Else
                i += 1
            End If
        End While

        If bFound Then
            GotoLoanIDIdx = i
        Else
            GotoLoanIDIdx = -1
        End If
    End Function

    Public Shared Function GotoLoanSetIDIdx(iLoanID As Integer) As Integer
        Dim i, j, iCounter As Integer
        Dim dr As DataRow
        Dim bFound As Boolean

        iCounter = dsLoanSets.Tables(0).Rows.Count

        i = 0
        bFound = False
        While i <= iCounter - 1 And Not bFound
            dr = dsLoans.Tables(0).Rows(i)
            j = GenDB.fnDBIntField(dr("LoanSetID"))
            If j = iLoanID Then
                bFound = True
            Else
                i += 1
            End If
        End While

        If bFound Then
            GotoLoanSetIDIdx = i
        Else
            GotoLoanSetIDIdx = -1
        End If
    End Function

    ' Get the last tranche number for the parent.
    Public Shared Function GetMaxTrancheValue(iLoanID_Idx) As Integer
        Dim iMax As Integer = 0
        Dim dr As DataRow
        Dim iParentLoanID, iRes As Integer
        Dim i, j, k, iCount As Integer

        ' Is the LoanIDX a parent or chilc
        dr = dsLoans.Tables(0).Rows(iLoanID_Idx)

        If GenDB.fnDBIntField(dr("Loan_Parent_ID")) = 0 Then ' it is a parent
            iParentLoanID = GenDB.fnDBIntField(dr("LoanID"))
        Else
            iParentLoanID = GenDB.fnDBIntField(dr("Loan_Parent_ID"))
        End If

        j = dsLoans.Tables(0).Rows.Count
        iCount = 0
        For i = 0 To j - 1
            If dsLoans.Tables(0).Rows(i)("Loan_Parent_ID") = iParentLoanID Then
                k = CInt(dsLoans.Tables(0).Rows(i)("Tranche_Number"))
                If k > iMax Then
                    iMax = k
                End If
                iCount += 1
            End If
        Next i

        If iCount > iMax Then
            iRes = iCount
        Else
            iRes = iMax
        End If

        GetMaxTrancheValue = iRes
    End Function

    Public Shared Sub LoadLoanSets(iLoanID As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim m As Integer
        Dim j As Integer
        Dim dr As DataRow


        sSQL = "select loanid, loanstatus, business_name as LoanName
                from Loans where LoanSetID > 0 and LoanSetID = 
                            (Select LoanSetID from Loans where LoanID = @p2 and LoanSetID > 0)"


        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iLoanID

        dsListLoanSets.Clear()
        Adaptor.Fill(dsListLoanSets)

        Try
            dr = dsListLoanSets.Tables(0).Rows(0)
            j = dsListLoanSets.Tables(0).Rows.Count


        Catch ex As Exception

        End Try

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Sub

    Public Shared Function GetLoanSet(iLoanID As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim m As Integer
        Dim j As Integer
        Dim k As Integer
        Dim dr As DataRow


        sSQL = "select loansetid
                from Loans where LoanID = @p1"


        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Adaptor = New FirebirdSql.Data.FirebirdClient.FbDataAdapter(sSQL, MyConn)

        Adaptor.SelectCommand.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iLoanID

        dsLoanSetiD.Clear()
        Adaptor.Fill(dsLoanSetiD)

        Try
            dr = dsLoanSetiD.Tables(0).Rows(0)
            j = dsLoanSetiD.Tables(0).Rows.Count
            k = GenDB.fnDBIntField(dr("LoanSetID"))

        Catch ex As Exception
            k = 0
        End Try

        GetLoanSet = k

        MyConn.Close()
        Adaptor = Nothing
        MyConn = Nothing
    End Function

    Public Shared Sub UpdateLoanSets(iLoanID As Integer, iloansetID As Integer)
        Dim strConn As String
        Dim MyConn As FirebirdSql.Data.FirebirdClient.FbConnection
        Dim Adaptor As FirebirdSql.Data.FirebirdClient.FbDataAdapter
        Dim sSQL As String
        Dim i As Integer

        Dim Cmd As FirebirdSql.Data.FirebirdClient.FbCommand

        sSQL = "update loans set loansetid = @p1 where loanid = @p2"

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString
        MyConn = New FirebirdSql.Data.FirebirdClient.FbConnection(strConn)

        MyConn.Open()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand
        Cmd.Connection = MyConn
        Cmd.CommandText = sSQL
        Cmd.Parameters.Clear()
        Cmd.Parameters.Add("@p1", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iloansetID
        Cmd.Parameters.Add("@p2", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = iLoanID

        Cmd.ExecuteNonQuery()

        MyConn.Close()
        Cmd = Nothing
        MyConn = Nothing


    End Sub



End Class
