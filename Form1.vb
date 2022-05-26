Imports System.Text.RegularExpressions
Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks


Public Class Form1
    Public SearchID As Integer
    Private Const STR_TabText_IndividualLender As String = "Individual Lender"
    Private Const STR_TabText_InstitutionalLender As String = "Institutional Lender"
    Private Const STR_TabText_ProfileDetails As String = "Profile Details"
    Private Const STR_TabText_TradingHistory As String = "Trading History"
    Private Const STR_TabText_FinancialTransactions As String = "Financial Transactions"
    Private Const STR_TabText_CurrentLoans As String = "Current Loans"
    Private Const STR_TabText_HistoricalLoans As String = "Historical Loans"
    Private Const STR_TabText_OutstandingSell As String = "Outstanding Sell"
    Private Const STR_TabText_MandateBalances As String = "Mandate Balances"
    Public Const STR_LenderFindByText_UserID As String = "User ID"
    Public Const STR_LenderFindByText_LastName As String = "Last Name"
    Public Const STR_LenderFindByText_FirstName As String = "First Name"
    Public Const STR_LenderFindByText_CompanyName As String = "Company Name"
    Public Const STR_LenderFindByText_AccountId As String = "Account ID"
    Private Const STR_TabText_Capacity As String = "Capacity"
    Private Const STR_BritishSterlingFormatForDataGridCloumns As String = "c"
    Private Const STR_BritishDateFormatForDataGridCloumns As String = "{dd/MM/yyyy}"
    Public Const STR_DdMMMyyyyHHmmss As String = "dd-MMM-yyyy HH:mm:ss"
    Public Shared _SystemUser As String
    Private Const STR_TwoDecimalsFormatForDataGridCloumns As String = "n2"

    Public _IndividualLenderStoredBankSortCode As String

    Public _IndividualLenderStoredBankAccountNum As String

    Public _InstitutionalLenderStoredBankSortCode As String

    Public _InstitutionalLenderStoredBankAccountNumber As String

    Public _CurrentSelectedTabIs As String



    'Private Sub PopulateIndividualLenderFilterBy()


    '    Dim mylist As New List(Of IndividualLenderProfileDetails)
    '    mylist = IndividualLenderProfileDetails.CreateIndividualLenderFilterListUID()
    '    Select Case listBoxIndividualLenderFilter.SelectedItem
    '        Case "User ID"
    '            Dim i As Integer = 0
    '            While i < mylist.Count
    '                comboBoxIndividualLenderFindBy.Items.Add(mylist(i).ListOfIndividualLendersByUidLastNameFirstName)
    '                i += 1
    '            End While

    '        Case "Last Name"

    '        Case "First Name"

    '        Case "Account ID"
    '            Dim i As Integer = 0
    '            While i < mylist.Count
    '                comboBoxIndividualLenderFindBy.Items.Add(mylist(i).ListOfIndividualLendersByAccountId)
    '                i += 1
    '            End While

    '        Case Else
    '            Dim i As Integer = 0
    '            While i < mylist.Count
    '                comboBoxIndividualLenderFindBy.Items.Add(mylist(i).ListOfIndividualLendersByAccountId)
    '                i += 1
    '            End While
    '    End Select

    '    comboBoxIndividualLenderFindBy.SelectedIndex = 0



    'End Sub


    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ' DiableInstitutionalLenderTextBoxes()

        DisableIndividualLenderTextBoxes()

        PopulateIndividualLenderFilterBy()

        WindowState = FormWindowState.Maximized
        txtServerName.Text = GetServerName()
        Cursor = Cursors.Default
    End Sub
    Private Sub PopulateIndividualLenderFilterBy()
        IndividualLenderProfileDetails.CreateIndividualLenderFilterList()
        listBoxIndividualLenderFilter.DataSource = IndividualLenderProfileDetails._IndividualLenderFilterNames.Keys.ToList()


    End Sub

    Private Sub PopulateIndividualLenderFindBy(ByVal selectedFilterByUser As String)

        IndividualLenderProfileDetails.PopulateIndividualLenderSearchList(selectedFilterByUser)
        comboBoxIndividualLenderFindBy.Items.Clear()
        Select Case selectedFilterByUser
            Case STR_LenderFindByText_UserID
                comboBoxIndividualLenderFindBy.Items.AddRange(IndividualLenderProfileDetails._ListOfIndividualLendersByUidLastNameFirstName.ToArray())
            Case STR_LenderFindByText_LastName
                'comboBoxIndividualLenderFindBy.ValueMember = "_Accountid"
                'comboBoxIndividualLenderFindBy.DisplayMember = "_Thename"
                'comboBoxIndividualLenderFindBy.DataSource = IndividualLenderProfileDetails.GetLastNameFirstnameAccID()
                comboBoxIndividualLenderFindBy.Items.AddRange(IndividualLenderProfileDetails._ListOfIndividualLendersByLastNameFirstNameUid.ToArray())
            Case STR_LenderFindByText_FirstName
                comboBoxIndividualLenderFindBy.Items.AddRange(IndividualLenderProfileDetails._ListOfIndividualLendersByFirstNameLastNameUid.ToArray())
            Case STR_LenderFindByText_AccountId
                comboBoxIndividualLenderFindBy.Items.AddRange(IndividualLenderProfileDetails._ListOfIndividualLendersByAccountId.ToArray())
            Case STR_LenderFindByText_CompanyName
                comboBoxIndividualLenderFindBy.Items.AddRange(IndividualLenderProfileDetails._ListOfIndividualLendersByUIDCompanyName.ToArray())
        End Select

        Dim cmbAutoStringList As AutoCompleteStringCollection = New AutoCompleteStringCollection()
        comboBoxIndividualLenderFindBy.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        comboBoxIndividualLenderFindBy.AutoCompleteSource = AutoCompleteSource.ListItems

    End Sub
    Private Sub listBoxIndividualLenderFilter__SelectedValueChanged(sender As Object, e As EventArgs) Handles listBoxIndividualLenderFilter.SelectedIndexChanged
        comboBoxIndividualLenderFindBy.Items.Clear()
        comboBoxIndividualLenderFindBy.SelectedIndex = -1
        comboBoxIndividualLenderFindBy.ResetText()
        PopulateIndividualLenderFindBy(listBoxIndividualLenderFilter.Text)
    End Sub

    Private Sub buttonIndividualLenderFind_Click(sender As Object, e As EventArgs) Handles buttonIndividualLenderFind.Click
        IndividualLenderProfileDetails._IndividualLenderIsSelected = True

        Dim strUserSelected As String = String.Empty
        strUserSelected = comboBoxIndividualLenderFindBy.Text.Replace(", ", "-")
        strUserSelected = comboBoxIndividualLenderFindBy.SelectedItem.ToString()
        IndividualLenderProfileDetails._LenderSearchValue = comboBoxIndividualLenderFindBy.SelectedItem.Replace(",", "-")
        IndividualLenderProfileDetails._LenderSearchFilter = listBoxIndividualLenderFilter.Text

        Cursor.Current = Cursors.WaitCursor

        _CurrentSelectedTabIs = STR_TabText_ProfileDetails
        strUserSelected = String.Empty
        strUserSelected = comboBoxIndividualLenderFindBy.Text.Replace(", ", "-")

        IndividualLenderProfileDetails._LenderSearchValue = comboBoxIndividualLenderFindBy.Text.Replace(",", "-")
        IndividualLenderProfileDetails._LenderSearchFilter = listBoxIndividualLenderFilter.Text

        'Profile Data

        GetIndividualLenderProfileData()

        'Enable all the textboxes on Individual lender

        EnableIndividualLenderTextBoxes()



        'Trading History

        Dim thCounts As Integer = PopulateIndividualLenderTradingHistory(IndividualLenderProfileDetails._LenderAccountId)
        If thCounts >= 1 Then
            dataGridViewIndividualLenderTradingHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Else
            dataGridViewIndividualLenderTradingHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If

        ' Financial Transactions

        Dim ftCounts As Integer = PopulateIndividualLenderFinancialTransactions(IndividualLenderProfileDetails._LenderAccountId)

        If ftCounts >= 1 Then
            dataGridViewIndividualLenderFinancialTransactions.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Else
            dataGridViewIndividualLenderFinancialTransactions.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If


        ' Current Loans
        Dim clCounts = PopulateIndividualLenderCurrentLoans()

        If clCounts >= 1 Then
            dataGridViewLenderIndividualCurrentLoans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Else
            dataGridViewLenderIndividualCurrentLoans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If


        ' Historical Loans
        Dim hlCounts As Integer = PopulateIndividualLenderHistoricalLoans(IndividualLenderProfileDetails._LenderAccountId)

        If hlCounts >= 1 Then
            dataGridViewIndividualLenderHistoricalLoans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Else
            dataGridViewIndividualLenderHistoricalLoans.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If

        ' Outstanding Sell

        Dim osCounts As Integer = PopulateIndividualLenderOutstandingSell(IndividualLenderProfileDetails._LenderAccountId)

        If osCounts >= 1 Then
            dataGridViewIndividualLenderOutstandingSell.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        Else
            dataGridViewIndividualLenderOutstandingSell.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If


        ' Mandate Blances

        Dim mbCounts As Integer = PopulateIndividualLenderMandateBalances(IndividualLenderProfileDetails._LenderAccountId)

        If mbCounts >= 1 Then
            dataGridViewIndividualLenderMandateBalances.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        Else
            dataGridViewIndividualLenderMandateBalances.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If

        ' Notes

        If PopulateLenderNotesOfSelectedLender(IndividualLenderProfileDetails._LenderAccountId) >= 1 Then
            dataGridViewLenderNotes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dataGridViewLenderNotes.Visible = True
        End If

        ' fetch Investor Summary

        Dim IsCounts As Integer = PopulateIndividualLenderInvestorSummary(IndividualLenderProfileDetails._LenderAccountId, IndividualLenderProfileDetails._LenderUserId)

        If IsCounts >= 1 Then
            dataGridViewIndividualLenderOutstandingSell.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        Else
            dataGridViewIndividualLenderOutstandingSell.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader
        End If

        ' Fetch Accrued Interest
        Dim accramount As Long
        textBoxIndividualLenderCapacityInterestAccruingOnCurrentLoans.Text = ""
        GenDB.EarnedInterest = 0
        accramount = GenDB.GetInterestAccruingOnCurrentLoans(IndividualLenderProfileDetails._LenderAccountId)
        textBoxIndividualLenderCapacityInterestAccruingOnCurrentLoans.Text = GenDB.PenceToCurrencyStringPounds(accramount)

        'tabPageIndividualLenderProfileDetials.Show()
        tabPageIndividualLenderProfileDetials.Focus()

    End Sub




    Public Sub GetIndividualLenderProfileData()
        Dim individualLenderSelected As String = comboBoxIndividualLenderFindBy.SelectedItem.ToString()


        Dim selectby As String
        selectby = "USERS.lastname"


        Select Case listBoxIndividualLenderFilter.SelectedItem
            Case "Last Name"
                selectby = "ACCOUNTS.accountId"
                individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-") + 1, individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-") - 1)
            Case "First Name"
                selectby = "ACCOUNTS.accountId"
                individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-") + 1, individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-") - 1)
            Case "User ID"
                selectby = "Users.userId"
                ' individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-") + 1, individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-") - 1)
                individualLenderSelected = individualLenderSelected.Substring(0, individualLenderSelected.IndexOf("-"))
            Case "Account ID"
                selectby = "ACCOUNTS.accountId"
                individualLenderSelected = individualLenderSelected.Substring(0, individualLenderSelected.IndexOf("-"))
            Case "Company Name"
                selectby = "Users.userId"
                individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-") + 1, individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-") - 1)
        End Select



        If PopulateIndividualLenderProfileDetails(selectby, individualLenderSelected) >= 1 Then
            _IndividualLenderStoredBankSortCode = txtboxIndividualLenderProfileDetailBankSortCode.Text
            _IndividualLenderStoredBankAccountNum = txtboxIndividualLenderProfileDetailBankAccountNumber.Text
            'LenderNotes._LenderAccountIdForNotes = IndividualLenderProfileDetails.LenderAccountId
            ' comboBoxIndividualLenderFindBy.BackColor = Color.Green
            dataGridViewLenderIndividualProfileDetails.AutoResizeColumns()
            '  ShowIndividualLendersActionButtons()
            '  ShowNotes()
            'buttonIndividualLenderFind.Visible = False
        Else
            comboBoxIndividualLenderFindBy.BackColor = Color.Red
        End If

        '  PopulateIndividualLenderCapacityInfo(IndividualLenderProfileDetails._LenderAccountId)
    End Sub
    Public Function PopulateIndividualLenderProfileDetails(ByVal searchBySqlField As String, ByVal searchValueFromSelectedByUser As String) As Integer
        Dim outputValue As Integer = -1
        IndividualLenderProfileDetails._IndividualLenderProfileDataTable = New DataTable()
        IndividualLenderProfileDetails._IndividualLenderProfileDataTable = IndividualLenderProfileDetails.GetIndividualLenderProfileData(searchBySqlField, searchValueFromSelectedByUser)

        dataGridViewLenderIndividualProfileDetails.DataSource = Nothing
        dataGridViewLenderIndividualProfileDetails.DataSource = IndividualLenderProfileDetails._IndividualLenderProfileDataTable
        dataGridViewLenderIndividualProfileDetails.AutoResizeColumns()


        If dataGridViewLenderIndividualProfileDetails.RowCount > -1 Then
            outputValue = dataGridViewLenderIndividualProfileDetails.RowCount
            PopulateIndividualLenderProfileDetailsinTextBoxes(IndividualLenderProfileDetails._IndividualLenderProfileDataTable)
        Else
            outputValue = -1
        End If

        dataGridViewLenderIndividualProfileDetails.[ReadOnly] = True
        Return outputValue
    End Function
    Public Function PopulateIndividualLenderCurrentLoans() As Integer

        Dim outputValue As Integer = -1

        Dim individualLenderSelected As String = comboBoxIndividualLenderFindBy.SelectedItem.ToString()

        Select Case listBoxIndividualLenderFilter.SelectedItem
            Case "Last Name"
                Dim i As Integer = individualLenderSelected.LastIndexOf("-")
                Dim t As Integer = individualLenderSelected.Length
                individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-") + 1, individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-") - 1)

            Case "First Name"
                individualLenderSelected = individualLenderSelected.Substring(individualLenderSelected.LastIndexOf("-"), individualLenderSelected.Length - individualLenderSelected.LastIndexOf("-"))

            Case "User ID"
                individualLenderSelected = individualLenderSelected.Substring(0, individualLenderSelected.IndexOf("-"))
            Case "Account ID"
                individualLenderSelected = individualLenderSelected.Substring(0, individualLenderSelected.IndexOf("-"))
        End Select





        Dim IndividualLenderCurrentLoansDataTable As DataTable = New DataTable()
        dataGridViewLenderIndividualCurrentLoans.DataSource = Nothing






        'Dim mylist As New List(Of IndividualLenderProfileDetails)
        'mylist = IndividualLenderProfileDetails.GetIndividualLenderCurrentLoans(SearchID)

        IndividualLenderCurrentLoansDataTable = IndividualLenderCurrentLoans.GetIndividualLenderCurrentLoansDT(individualLenderSelected)
        Dim IndividualLenderCurrentLoansDataTableRowsCount As Integer = IndividualLenderCurrentLoansDataTable.Rows.Count

        dataGridViewLenderIndividualCurrentLoans.DataSource = IndividualLenderCurrentLoansDataTable
        If dataGridViewLenderIndividualCurrentLoans.RowCount > -1 Then
            outputValue = dataGridViewLenderIndividualCurrentLoans.RowCount
            PopulateIndividualLenderProfileDetailsinTextBoxes(IndividualLenderProfileDetails._IndividualLenderProfileDataTable)
        Else
            outputValue = -1
        End If

        ' labelLoanName.Text = String.Empty
        textBoxIndividualLenderSelectedLastLentLoanId.Text = String.Empty
        textBoxIndividualLenderSelectedLastLentDate.Text = String.Empty
        textBoxIndividualLenderSelectedLastLentAmount.Text = "£0.00"
        textBoxIndividualLenderNextMaturity.Text = String.Empty
        If IndividualLenderCurrentLoansDataTableRowsCount > 0 Then

            labelLoanName.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsLoanName
            labelLoanName.Visible = True
            textBoxIndividualLenderSelectedLastLentLoanId.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsLoanID
            textBoxIndividualLenderSelectedLastLentDate.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsLoanDateTime
            labelLoanName.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsLoanName
            textBoxIndividualLenderSelectedLastLentAmount.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsLoanAmount
            textBoxIndividualLenderNextMaturity.Text = IndividualLenderCurrentLoans._IndividualLenderLastLentDetailsNextMaturityDateTime

            dataGridViewLenderIndividualCurrentLoans.Columns("description").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("earlyredemption").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("earlyredemptiondate").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("secondary_trading").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("lh_bals_id").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("lh_id").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("StartDate").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("background").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("term").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("maxloanamount").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("company_risk_as_stars").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("security_guarantees").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("ipo_end").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("purpose_of_loan").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("people").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("anythingelse").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("loanstatus").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("dd_Date").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("loantype").Visible = False
            dataGridViewLenderIndividualCurrentLoans.Columns("Loan_Parent_id").Visible = False

            dataGridViewLenderIndividualCurrentLoans.Columns("TheLoan").DisplayIndex = 0
            dataGridViewLenderIndividualCurrentLoans.Columns("CompanyName").DisplayIndex = 1
            dataGridViewLenderIndividualCurrentLoans.Columns("therate").DisplayIndex = 2
            dataGridViewLenderIndividualCurrentLoans.Columns("outstanding").DisplayIndex = 3
            dataGridViewLenderIndividualCurrentLoans.Columns("LastDate").DisplayIndex = 4

            dataGridViewLenderIndividualCurrentLoans.Columns("TheLoan").HeaderText = "Loan Ref"
            dataGridViewLenderIndividualCurrentLoans.Columns("CompanyName").HeaderText = "Loan Name"
            dataGridViewLenderIndividualCurrentLoans.Columns("therate").HeaderText = "Interest Rate"
            dataGridViewLenderIndividualCurrentLoans.Columns("outstanding").HeaderText = "Outstanding Balance"

            dataGridViewLenderIndividualCurrentLoans.Columns("LastDate").HeaderText = "Maturity Date"


            FormatDataGridViewToDecimal(dataGridViewLenderIndividualCurrentLoans, "therate")
            FormatDataGridViewToBritishMonetory(dataGridViewLenderIndividualCurrentLoans, "outstanding")

            dataGridViewLenderIndividualCurrentLoans.AutoGenerateColumns = True
            dataGridViewLenderIndividualCurrentLoans.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Else
            textBoxIndividualLenderSelectedLastLentLoanId.Text = "No current Loans"
            labelLoanName.Text = "No current loans"
            textBoxIndividualLenderSelectedLastLentLoanId.Text = String.Empty
            textBoxIndividualLenderSelectedLastLentDate.Text = String.Empty
            textBoxIndividualLenderSelectedLastLentAmount.Text = "£0.00"
        End If
        Return outputValue
    End Function

    Public Function PopulateIndividualLenderTradingHistory(ByVal searchValueFromSelectedByUser As String)
        Dim outputValue As Integer = -1

        dataGridViewIndividualLenderTradingHistory.DataSource = IndividualLenderTradingHistory.GetIndividualLenderTradingHistory(searchValueFromSelectedByUser)


        If dataGridViewIndividualLenderTradingHistory.RowCount > -1 Then
            outputValue = dataGridViewIndividualLenderTradingHistory.RowCount
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderTradingHistory, "ORDER_AMOUNT")
            dataGridViewIndividualLenderTradingHistory.AutoResizeColumns()
            dataGridViewIndividualLenderTradingHistory.[ReadOnly] = True
            dataGridViewIndividualLenderTradingHistory.AutoGenerateColumns = True
            dataGridViewIndividualLenderTradingHistory.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dataGridViewIndividualLenderTradingHistory.Columns("LH_ID").Visible = False
            dataGridViewIndividualLenderTradingHistory.Columns("ORDER_TRANS_DATE").HeaderText = "Transaction Date"
            dataGridViewIndividualLenderTradingHistory.Columns("LOAN_NAME").HeaderText = "Loan Name"
            dataGridViewIndividualLenderTradingHistory.Columns("LOAN_ID").HeaderText = "Loan Ref"
            dataGridViewIndividualLenderTradingHistory.Columns("THE_BUYER").HeaderText = "Buyer"
            dataGridViewIndividualLenderTradingHistory.Columns("THE_SELLER").HeaderText = "Seller"
            dataGridViewIndividualLenderTradingHistory.Columns("ORDER_AMOUNT").HeaderText = "Amount"
            dataGridViewIndividualLenderTradingHistory.Columns("Premium").HeaderText = "Premium"
            dataGridViewIndividualLenderTradingHistory.AutoGenerateColumns = False
        Else
            outputValue = -1
        End If

        dataGridViewIndividualLenderTradingHistory.AutoResizeColumns()
        dataGridViewIndividualLenderTradingHistory.[ReadOnly] = True

        Return outputValue
    End Function

    Public Function PopulateIndividualLenderFinancialTransactions(ByVal searchValueFromSelectedByUser As String)
        Dim outputValue As Integer = -1

        dataGridViewIndividualLenderFinancialTransactions.DataSource = Nothing

        IndividualLenderFinancialTransactions._IndividualLenderTransactionsReportDefault = True
        dataGridViewIndividualLenderFinancialTransactions.DataSource = IndividualLenderFinancialTransactions.GetIndividualLenderAnnualStatementData().Tables(0)


        dataGridViewIndividualLenderFinancialTransactions.AutoResizeColumns()


        If dataGridViewIndividualLenderFinancialTransactions.RowCount > -1 Then
            outputValue = dataGridViewIndividualLenderFinancialTransactions.RowCount
            dataGridViewIndividualLenderFinancialTransactions.AutoGenerateColumns = True
            dataGridViewIndividualLenderFinancialTransactions.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dataGridViewIndividualLenderFinancialTransactions.Columns("LOANDESC2").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNT1").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEBALANCE1").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEACTUALBALANCE1").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("BIDID").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("STATEMENTBALANCEAMOUNT").Visible = False
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEDATE").HeaderText = "Date"
            dataGridViewIndividualLenderFinancialTransactions.Columns("TRANSTYPE").HeaderText = "Transaction Type Code"
            dataGridViewIndividualLenderFinancialTransactions.Columns("TRANSDESC").HeaderText = "Transaction Type Description"
            dataGridViewIndividualLenderFinancialTransactions.Columns("LOANDESC").HeaderText = "Loan"
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTDEPOSIT").HeaderText = "Deposit"
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTDEPOSIT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTWITHDRAW").HeaderText = "Withdraw"
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTWITHDRAW").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEACTUALBALANCE").HeaderText = "Balance"
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEACTUALBALANCE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dataGridViewIndividualLenderFinancialTransactions.Columns("FIN_TRANSID").HeaderText = "Fin Transactions ID"
            dataGridViewIndividualLenderFinancialTransactions.Columns("HEADER_ID").HeaderText = "Header ID"
            dataGridViewIndividualLenderFinancialTransactions.Columns("FTORDERID").HeaderText = "FT Order D"
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEDATE").DisplayIndex = 0
            dataGridViewIndividualLenderFinancialTransactions.Columns("TRANSDESC").DisplayIndex = 1
            dataGridViewIndividualLenderFinancialTransactions.Columns("LOANDESC").DisplayIndex = 2
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTDEPOSIT").DisplayIndex = 3
            dataGridViewIndividualLenderFinancialTransactions.Columns("AMOUNTWITHDRAW").DisplayIndex = 4
            dataGridViewIndividualLenderFinancialTransactions.Columns("THEACTUALBALANCE").DisplayIndex = 5
            dataGridViewIndividualLenderFinancialTransactions.Columns("FIN_TRANSID").DisplayIndex = 6
            dataGridViewIndividualLenderFinancialTransactions.Columns("HEADER_ID").DisplayIndex = 7
            dataGridViewIndividualLenderFinancialTransactions.Columns("FTORDERID").DisplayIndex = 8
            dataGridViewIndividualLenderFinancialTransactions.Columns("LOANID").DisplayIndex = 9
            dataGridViewIndividualLenderFinancialTransactions.Columns("LH_ID").DisplayIndex = 10
            dataGridViewIndividualLenderFinancialTransactions.Columns("TRANSTYPE").DisplayIndex = 11
            dataGridViewIndividualLenderFinancialTransactions.AutoGenerateColumns = False
        Else
            outputValue = -1
        End If

        dataGridViewIndividualLenderFinancialTransactions.[ReadOnly] = True

        Return outputValue
    End Function

    Public Function PopulateIndividualLenderHistoricalLoans(ByVal searchValueFromSelectedByUser As String)
        Dim outputValue As Integer = -1

        dataGridViewIndividualLenderHistoricalLoans.AutoGenerateColumns = True
        dataGridViewIndividualLenderHistoricalLoans.DataSource = IndividualLenderHistoricalLoans.GetIndividualLenderHistoricalLoans(searchValueFromSelectedByUser)

        dataGridViewIndividualLenderHistoricalLoans.AutoResizeColumns()

        If dataGridViewIndividualLenderHistoricalLoans.RowCount > -1 Then
            outputValue = dataGridViewIndividualLenderHistoricalLoans.RowCount
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderHistoricalLoans, "Interest_Received")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderHistoricalLoans, "PREMIUM_AMOUNT")
            dataGridViewIndividualLenderHistoricalLoans.AutoGenerateColumns = True
            dataGridViewIndividualLenderHistoricalLoans.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dataGridViewIndividualLenderHistoricalLoans.Columns("The_Loan").HeaderText = "Loan Ref"
            dataGridViewIndividualLenderHistoricalLoans.Columns("Company_Name").HeaderText = "Company Name"
            dataGridViewIndividualLenderHistoricalLoans.Columns("DD_Lastdate").HeaderText = "Date Completed"
            dataGridViewIndividualLenderHistoricalLoans.Columns("Interest_Received").HeaderText = "Interest Received"
            dataGridViewIndividualLenderHistoricalLoans.Columns("Description").Visible = False
            dataGridViewIndividualLenderHistoricalLoans.AutoGenerateColumns = False
        Else
            outputValue = -1
        End If

        dataGridViewIndividualLenderHistoricalLoans.AutoResizeColumns()
        dataGridViewIndividualLenderHistoricalLoans.[ReadOnly] = True

        Return outputValue
    End Function

    Public Function PopulateIndividualLenderMandateBalances(ByVal searchValueFromSelectedByUser As String)
        Dim outputValue As Integer = -1
        txtCurrentConcentration.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderTotalFundBalance)
        dataGridViewIndividualLenderOutstandingSell.DataSource = IndividualLenderOutstandingSell.GetIndividualLenderOutstandingSell(searchValueFromSelectedByUser)

        dataGridViewIndividualLenderOutstandingSell.AutoResizeColumns()

        If dataGridViewIndividualLenderOutstandingSell.RowCount > -1 Then
            FormatDataGridViewToDecimal(dataGridViewIndividualLenderOutstandingSell, "NEW_RATE")
            dataGridViewIndividualLenderOutstandingSell.AutoGenerateColumns = True
            dataGridViewIndividualLenderOutstandingSell.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            outputValue = dataGridViewIndividualLenderOutstandingSell.RowCount
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "MAX_LOAN_AMOUNT")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "AMOUNT_FULL")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "PREMIUM")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "SALE_AMOUNT")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "OUTSTANDING_AMOUNT")
            dataGridViewIndividualLenderOutstandingSell.Columns("MAX_LOAN_AMOUNT").HeaderText = "Max Loan Name"
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").HeaderText = "Company Name"
            dataGridViewIndividualLenderOutstandingSell.Columns("DESCRIPTION").HeaderText = "Description"
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").HeaderText = "Loan"
            dataGridViewIndividualLenderOutstandingSell.Columns("IPO_END").HeaderText = "Ipo End"
            dataGridViewIndividualLenderOutstandingSell.Columns("SECURITY_GURANTEES").HeaderText = "Security Gurantees"
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").HeaderText = "Term"
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_RISK_STARS").HeaderText = "Company Risk Stars"
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_ORDER").HeaderText = "Order"
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").HeaderText = "Account Id"
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").HeaderText = "Outstanding Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("MONTHS_LEFT").HeaderText = "Months Left"
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").HeaderText = "Sale Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").HeaderText = "Premium Percent"
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").HeaderText = "Premium Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").HeaderText = "Full Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").HeaderText = "Intetest Rate"
            dataGridViewIndividualLenderOutstandingSell.Columns("ANYTHING_ELSE").HeaderText = "Anything Else"
            dataGridViewIndividualLenderOutstandingSell.Columns("PURPOSE_OF_LOAN").HeaderText = "Purpose of Loan"
            dataGridViewIndividualLenderOutstandingSell.Columns("BACKGROUND").HeaderText = "Background"
            dataGridViewIndividualLenderOutstandingSell.Columns("PEOPLE").HeaderText = "People"
            dataGridViewIndividualLenderOutstandingSell.Columns("MAX_LOAN_AMOUNT").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("DESCRIPTION").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("IPO_END").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("SECURITY_GURANTEES").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_RISK_STARS").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_ORDER").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("MONTHS_LEFT").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("ANYTHING_ELSE").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("PURPOSE_OF_LOAN").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("BACKGROUND").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("PEOPLE").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").DisplayIndex = 0
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").DisplayIndex = 1
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").DisplayIndex = 2
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").DisplayIndex = 3
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").DisplayIndex = 4
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").DisplayIndex = 5
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").DisplayIndex = 6
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").DisplayIndex = 7
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").DisplayIndex = 8
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").DisplayIndex = 9
        Else
            outputValue = -1
        End If

        dataGridViewIndividualLenderOutstandingSell.AutoResizeColumns()
        dataGridViewIndividualLenderOutstandingSell.[ReadOnly] = True

        Return outputValue
    End Function
    Public Function PopulateIndividualLenderOutstandingSell(ByVal searchValueFromSelectedByUser As String)
        Dim outputValue As Integer = -1

        dataGridViewIndividualLenderOutstandingSell.DataSource = IndividualLenderOutstandingSell.GetIndividualLenderOutstandingSell(searchValueFromSelectedByUser)

        dataGridViewIndividualLenderOutstandingSell.AutoResizeColumns()

        If dataGridViewIndividualLenderOutstandingSell.RowCount > -1 Then
            FormatDataGridViewToDecimal(dataGridViewIndividualLenderOutstandingSell, "NEW_RATE")
            dataGridViewIndividualLenderOutstandingSell.AutoGenerateColumns = True
            dataGridViewIndividualLenderOutstandingSell.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            outputValue = dataGridViewIndividualLenderOutstandingSell.RowCount
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "MAX_LOAN_AMOUNT")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "AMOUNT_FULL")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "PREMIUM")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "SALE_AMOUNT")
            FormatDataGridViewToBritishMonetory(dataGridViewIndividualLenderOutstandingSell, "OUTSTANDING_AMOUNT")
            dataGridViewIndividualLenderOutstandingSell.Columns("MAX_LOAN_AMOUNT").HeaderText = "Max Loan Name"
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").HeaderText = "Company Name"
            dataGridViewIndividualLenderOutstandingSell.Columns("DESCRIPTION").HeaderText = "Description"
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").HeaderText = "Loan"
            dataGridViewIndividualLenderOutstandingSell.Columns("IPO_END").HeaderText = "Ipo End"
            dataGridViewIndividualLenderOutstandingSell.Columns("SECURITY_GURANTEES").HeaderText = "Security Gurantees"
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").HeaderText = "Term"
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_RISK_STARS").HeaderText = "Company Risk Stars"
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_ORDER").HeaderText = "Order"
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").HeaderText = "Account Id"
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").HeaderText = "Outstanding Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("MONTHS_LEFT").HeaderText = "Months Left"
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").HeaderText = "Sale Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").HeaderText = "Premium Percent"
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").HeaderText = "Premium Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").HeaderText = "Full Amount"
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").HeaderText = "Intetest Rate"
            dataGridViewIndividualLenderOutstandingSell.Columns("ANYTHING_ELSE").HeaderText = "Anything Else"
            dataGridViewIndividualLenderOutstandingSell.Columns("PURPOSE_OF_LOAN").HeaderText = "Purpose of Loan"
            dataGridViewIndividualLenderOutstandingSell.Columns("BACKGROUND").HeaderText = "Background"
            dataGridViewIndividualLenderOutstandingSell.Columns("PEOPLE").HeaderText = "People"
            dataGridViewIndividualLenderOutstandingSell.Columns("MAX_LOAN_AMOUNT").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("DESCRIPTION").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("IPO_END").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("SECURITY_GURANTEES").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_RISK_STARS").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_ORDER").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("MONTHS_LEFT").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").Visible = True
            dataGridViewIndividualLenderOutstandingSell.Columns("ANYTHING_ELSE").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("PURPOSE_OF_LOAN").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("BACKGROUND").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("PEOPLE").Visible = False
            dataGridViewIndividualLenderOutstandingSell.Columns("COMPANY_NAME").DisplayIndex = 0
            dataGridViewIndividualLenderOutstandingSell.Columns("THE_LOAN").DisplayIndex = 1
            dataGridViewIndividualLenderOutstandingSell.Columns("SALE_AMOUNT").DisplayIndex = 2
            dataGridViewIndividualLenderOutstandingSell.Columns("OUTSTANDING_AMOUNT").DisplayIndex = 3
            dataGridViewIndividualLenderOutstandingSell.Columns("NEW_RATE").DisplayIndex = 4
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM_PERCENT").DisplayIndex = 5
            dataGridViewIndividualLenderOutstandingSell.Columns("PREMIUM").DisplayIndex = 6
            dataGridViewIndividualLenderOutstandingSell.Columns("AMOUNT_FULL").DisplayIndex = 7
            dataGridViewIndividualLenderOutstandingSell.Columns("TERM").DisplayIndex = 8
            dataGridViewIndividualLenderOutstandingSell.Columns("ACCOUNT_ID").DisplayIndex = 9
        Else
            outputValue = -1
        End If

        dataGridViewIndividualLenderOutstandingSell.AutoResizeColumns()
        dataGridViewIndividualLenderOutstandingSell.[ReadOnly] = True

        Return outputValue
    End Function

    Public Function PopulateLenderNotesOfSelectedLender(ByVal searchValueFromSelectedByUser As String) As Integer
        Dim outputValue As Integer = -1

        dataGridViewLenderNotes.DataSource = LenderNotes.GetLenderNotes(searchValueFromSelectedByUser)

        dataGridViewLenderNotes.AutoResizeColumns()

        If dataGridViewLenderNotes.RowCount > -1 Then
            outputValue = dataGridViewLenderNotes.RowCount
        Else
            outputValue = -1
        End If

        dataGridViewLenderNotes.AutoResizeColumns()
        dataGridViewLenderNotes.[ReadOnly] = True
        Return outputValue
    End Function

    Public Function PopulateIndividualLenderInvestorSummary(ByVal LenderAccountID As String, ByVal LenderUserID As String)
        Dim outputValue As Integer = -1

        If IndividualLenderInvestorSummary.GetIndividualLenderIS(LenderAccountID, LenderUserID) = True Then


            textBoxIndividualLenderCapacityTotalLent.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderTotalLent)
            textBoxIndividualLenderCapacityTotalFundsBalance.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderTotalFundBalance)
            textBoxIndividualLenderCapacityTotalCurrentLoans.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderTotalCurrentLoans)
            textBoxIndividualLenderCapacityTotalInterestEarned.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderTotalGrossInterestEarned)
            textBoxIndividualLenderCapacityAmountAvailableToLend.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderAmountAvailableToLend)
            textBoxIndividualLenderCapacityBidsOutstanding.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderBidsOutstanding)
            textBoxIndividualLenderCapacityInterestAccruingOnCurrentLoans.Text = GenDB.PenceToCurrencyStringPounds(IndividualLenderInvestorSummary._LenderInterestAccrued)

        Else

        End If


    End Function


    Public Shared Sub FormatDataGridViewToDecimal(ByVal dataGridViewToFormat As DataGridView, ByVal columnNameToFormat As String)
        dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Format = STR_TwoDecimalsFormatForDataGridCloumns
        dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Public Shared Sub FormatDataGridViewToBritishMonetory(ByVal dataGridViewToFormat As DataGridView, ByVal columnNameToFormat As String)
        If (dataGridViewToFormat.ColumnCount > 0) AndAlso dataGridViewToFormat.Columns.Contains(columnNameToFormat) Then
            dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Format = STR_BritishSterlingFormatForDataGridCloumns
            dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    Public Shared Sub FormatDataGridViewToBritishDate(ByVal dataGridViewToFormat As DataGridView, ByVal columnNameToFormat As String)
        If (dataGridViewToFormat.ColumnCount > 0) AndAlso dataGridViewToFormat.Columns.Contains(columnNameToFormat) Then
            dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Format = STR_BritishDateFormatForDataGridCloumns
            dataGridViewToFormat.Columns(columnNameToFormat).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
    End Sub

    Private Sub PopulateIndividualLenderProfileDetailsinTextBoxes(ByVal indivudualLenderProfileDetailsDataTable As DataTable)
        ClearIndividualLenderProfileDetailsTextBoxes()
        txtboxIndividualLenderProfileDetailAccountCode.Text = IndividualLenderProfileDetails._LenderAccountId
        txtboxIndividualLenderProfileDetailAccountType.Text = IndividualLenderProfileDetails._LenderAccountType
        txtboxIndividualLenderProfileDetailLenderType.Text = IndividualLenderProfileDetails._LenderType
        txtboxIndividualLenderProfileDetailLenderCategory.Text = String.Empty
        txtboxIndividualLenderProfileDetailLenderCategory.Text = IndividualLenderProfileDetails._LenderCategory
        txtboxIndividualLenderProfileDetailActivationType.Text = IndividualLenderProfileDetails._LenderActivationType
        'txtboxInstitutionalLenderProfileDetailActivatedCert.Text = InstitutionalLenderProfileDetails._LenderActivatedCert
        txtboxIndividualLenderProfileDetailBankActivationType.Text = IndividualLenderProfileDetails._LenderBankActivationType
        txtboxIndividualLenderProfileDetailActivatedCert.Text = IndividualLenderProfileDetails._LenderActivatedCert
        txtboxIndividualLenderProfileDetailCreatedDate.Text = IndividualLenderProfileDetails._LenderCreatedDate
        txtboxIndividualLenderProfileDetail1stEmailSent.Text = IndividualLenderProfileDetails._LenderFirstEmailSentOn
        txtboxIndividualLenderProfileDetailEmailActivation.Text = IndividualLenderProfileDetails._LenderEmailActivation
        txtboxIndividualLenderProfileDetailForeName.Text = IndividualLenderProfileDetails._LenderForeName
        txtboxIndividualLenderProfileDetailSurname.Text = IndividualLenderProfileDetails._LenderSurname
        txtboxIndividualLenderProfileDetailDOB.Text = IndividualLenderProfileDetails._LenderDOB
        IndividualLenderProfileDetails._LenderAccountId = $"{txtboxIndividualLenderProfileDetailAccountCode.Text}"
        txtboxIndividualLenderProfileDetailEmailAddress.Text = IndividualLenderProfileDetails._LenderEmailAddress
        txtboxIndividualLenderProfileDetailPhoneNumber.Text = IndividualLenderProfileDetails._LenderPhoneNumber
        txtboxIndividualLenderProfileDetail1stEmailSent.Text = IndividualLenderProfileDetails._LenderFirstEmailSentOn
        txtboxIndividualLenderProfileDetailBankAccountName.Text = IndividualLenderProfileDetails._LenderBankAccountName
        txtboxIndividualLenderProfileDetailBankAccountNumber.Text = IndividualLenderProfileDetails._LenderBankAccountNumber
        txtboxIndividualLenderProfileDetailBankSortCode.Text = IndividualLenderProfileDetails._LenderBankSortCode
        IndividualLenderProfileDetails._LenderBankAccountNumber = txtboxIndividualLenderProfileDetailBankAccountNumber.Text
        IndividualLenderProfileDetails._LenderBankSortCode = $"{txtboxIndividualLenderProfileDetailBankSortCode.Text}"
        txtboxIndividualLenderProfileDetailBuildingSocietyRollNum.Text = IndividualLenderProfileDetails._LenderBuildingSocietyRollNum
        txtboxIndividualLenderProfileDetailHouseNameNum.Text = IndividualLenderProfileDetails._LenderHouseNameNum
        txtboxIndividualLenderProfileDetailAddress1.Text = IndividualLenderProfileDetails._LenderAddress1
        txtboxIndividualLenderProfileDetailAddress2.Text = IndividualLenderProfileDetails._LenderAddress2
        txtboxIndividualLenderProfileDetailAddress3.Text = IndividualLenderProfileDetails._LenderAddress3
        txtboxIndividualLenderProfileDetailAddress4.Text = IndividualLenderProfileDetails._LenderAddress4
        txtboxIndividualLenderProfileDetailTownCity.Text = IndividualLenderProfileDetails._LenderTownCity
        txtboxIndividualLenderProfileDetailCounty.Text = IndividualLenderProfileDetails._LenderCounty
        txtboxIndividualLenderProfileDetailCountry.Text = IndividualLenderProfileDetails._LenderCountry
        txtboxIndividualLenderProfileDetailPostCode.Text = IndividualLenderProfileDetails._LenderPostCode
    End Sub


    Private Sub ClearIndividualLenderProfileDetailsTextBoxes()
        txtboxIndividualLenderProfileDetailAccountCode.Text = String.Empty
        txtboxIndividualLenderProfileDetailLenderType.Text = String.Empty
        txtboxIndividualLenderProfileDetailProfClient.Text = String.Empty
        txtboxIndividualLenderProfileDetailActivationType.Text = String.Empty
        txtboxIndividualLenderProfileDetailBankActivationType.Text = String.Empty
        txtboxIndividualLenderProfileDetailCreatedDate.Text = String.Empty
        txtboxIndividualLenderProfileDetail1stEmailSent.Text = String.Empty
        txtboxIndividualLenderProfileDetailEmailActivation.Text = String.Empty
        txtboxIndividualLenderProfileDetailForeName.Text = String.Empty
        txtboxIndividualLenderProfileDetailSurname.Text = String.Empty
        txtboxIndividualLenderProfileDetailDOB.Text = String.Empty
        txtboxIndividualLenderProfileDetailLenderCategory.Text = String.Empty
        txtboxIndividualLenderProfileDetailEmailAddress.Text = String.Empty
        txtboxIndividualLenderProfileDetailPhoneNumber.Text = String.Empty
        txtboxIndividualLenderProfileDetail1stEmailSent.Text = String.Empty
        txtboxIndividualLenderProfileDetailBankAccountName.Text = String.Empty
        txtboxIndividualLenderProfileDetailBankAccountNumber.Text = String.Empty
        txtboxIndividualLenderProfileDetailBankSortCode.Text = String.Empty
        txtboxIndividualLenderProfileDetailBuildingSocietyRollNum.Text = String.Empty
        txtboxIndividualLenderProfileDetailHouseNameNum.Text = String.Empty
        txtboxIndividualLenderProfileDetailAddress1.Text = String.Empty
        txtboxIndividualLenderProfileDetailAddress2.Text = String.Empty
        txtboxIndividualLenderProfileDetailAddress3.Text = String.Empty
        txtboxIndividualLenderProfileDetailAddress4.Text = String.Empty
        txtboxIndividualLenderProfileDetailTownCity.Text = String.Empty
        txtboxIndividualLenderProfileDetailCounty.Text = String.Empty
        txtboxIndividualLenderProfileDetailCountry.Text = String.Empty
        txtboxIndividualLenderProfileDetailPostCode.Text = String.Empty
    End Sub

    Private Sub DisableIndividualLenderTextBoxes()
        txtboxIndividualLenderProfileDetailAccountCode.Enabled = False
        txtboxIndividualLenderProfileDetailLenderType.Enabled = False
        txtboxIndividualLenderProfileDetailLenderCategory.Enabled = False
        txtboxIndividualLenderProfileDetailLenderCategory.Enabled = False
        txtboxIndividualLenderProfileDetailActivationType.Enabled = False
        txtboxIndividualLenderProfileDetailBankActivationType.Enabled = False
        txtboxIndividualLenderProfileDetailCreatedDate.Enabled = False
        txtboxIndividualLenderProfileDetail1stEmailSent.Enabled = False
        txtboxIndividualLenderProfileDetailEmailActivation.Enabled = False
        txtboxIndividualLenderProfileDetailForeName.Enabled = False
        txtboxIndividualLenderProfileDetailSurname.Enabled = False
        txtboxIndividualLenderProfileDetailDOB.Enabled = False
        txtboxIndividualLenderProfileDetailEmailAddress.Enabled = False
        txtboxIndividualLenderProfileDetailPhoneNumber.Enabled = False
        txtboxIndividualLenderProfileDetail1stEmailSent.Enabled = False
        txtboxIndividualLenderProfileDetailBankAccountName.Enabled = False
        txtboxIndividualLenderProfileDetailBankAccountNumber.Enabled = False
        txtboxIndividualLenderProfileDetailBankSortCode.Enabled = False
        txtboxIndividualLenderProfileDetailBuildingSocietyRollNum.Enabled = False
        txtboxIndividualLenderProfileDetailHouseNameNum.Enabled = False
        txtboxIndividualLenderProfileDetailAddress1.Enabled = False
        txtboxIndividualLenderProfileDetailAddress2.Enabled = False
        txtboxIndividualLenderProfileDetailAddress3.Enabled = False
        txtboxIndividualLenderProfileDetailAddress4.Enabled = False
        txtboxIndividualLenderProfileDetailTownCity.Enabled = False
        txtboxIndividualLenderProfileDetailCounty.Enabled = False
        txtboxIndividualLenderProfileDetailCountry.Enabled = False
        txtboxIndividualLenderProfileDetailPostCode.Enabled = False
    End Sub
    Private Sub EnableIndividualLenderTextBoxes()
        txtboxIndividualLenderProfileDetailAccountCode.Enabled = True
        txtboxIndividualLenderProfileDetailLenderType.Enabled = True
        txtboxIndividualLenderProfileDetailLenderCategory.Enabled = True
        txtboxIndividualLenderProfileDetailLenderCategory.Enabled = True
        txtboxIndividualLenderProfileDetailActivationType.Enabled = True
        txtboxIndividualLenderProfileDetailBankActivationType.Enabled = True
        txtboxIndividualLenderProfileDetailCreatedDate.Enabled = True
        txtboxIndividualLenderProfileDetail1stEmailSent.Enabled = True
        txtboxIndividualLenderProfileDetailEmailActivation.Enabled = True
        txtboxIndividualLenderProfileDetailForeName.Enabled = True
        txtboxIndividualLenderProfileDetailSurname.Enabled = True
        txtboxIndividualLenderProfileDetailDOB.Enabled = True
        txtboxIndividualLenderProfileDetailEmailAddress.Enabled = True
        txtboxIndividualLenderProfileDetailPhoneNumber.Enabled = True
        txtboxIndividualLenderProfileDetail1stEmailSent.Enabled = True
        txtboxIndividualLenderProfileDetailBankAccountName.Enabled = True
        txtboxIndividualLenderProfileDetailBankAccountNumber.Enabled = True
        txtboxIndividualLenderProfileDetailBankSortCode.Enabled = True
        txtboxIndividualLenderProfileDetailBuildingSocietyRollNum.Enabled = True
        txtboxIndividualLenderProfileDetailHouseNameNum.Enabled = True
        txtboxIndividualLenderProfileDetailAddress1.Enabled = True
        txtboxIndividualLenderProfileDetailAddress2.Enabled = True
        txtboxIndividualLenderProfileDetailAddress3.Enabled = True
        txtboxIndividualLenderProfileDetailAddress4.Enabled = True
        txtboxIndividualLenderProfileDetailTownCity.Enabled = True
        txtboxIndividualLenderProfileDetailCounty.Enabled = True
        txtboxIndividualLenderProfileDetailCountry.Enabled = True
        txtboxIndividualLenderProfileDetailPostCode.Enabled = True
    End Sub


    Private Sub buttonIndividualLenderProfileDetailBankSortCodeDecrypt_Click(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeDecrypt.Click
        If _IndividualLenderStoredBankSortCode.Equals(txtboxIndividualLenderProfileDetailBankSortCode.Text) Then
            txtboxIndividualLenderProfileDetailBankSortCode.Text = Crypt.Decrypt(txtboxIndividualLenderProfileDetailBankSortCode.Text)
        End If

    End Sub



    Private Sub buttonIndividualLenderProfileDetailBankAccountNumDecrypt_Click(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumDecrypt.Click
        If _IndividualLenderStoredBankAccountNum.Equals(txtboxIndividualLenderProfileDetailBankAccountNumber.Text) Then
            txtboxIndividualLenderProfileDetailBankAccountNumber.Text = Crypt.Decrypt(txtboxIndividualLenderProfileDetailBankAccountNumber.Text)
        End If
    End Sub


    Private Sub buttonIndividualLenderProfileDetailBankAccountNumDecrypt_MouseHover_1(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumDecrypt.MouseHover
        Cursor = Cursors.Hand
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankAccountNumDecrypt_MouseLeave(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumDecrypt.MouseLeave
        Cursor = Cursors.Default
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankSortCodeDecrypt_MouseHover(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeDecrypt.MouseHover
        Cursor = Cursors.Hand
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankSortCodeDecrypt_MouseLeave(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeDecrypt.MouseLeave
        Cursor = Cursors.Default
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankSortCodeEncrypt_Click(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeEncrypt.Click
        If Not _IndividualLenderStoredBankSortCode.Equals(txtboxIndividualLenderProfileDetailBankAccountNumber.Text) Then
            txtboxIndividualLenderProfileDetailBankSortCode.Text = _IndividualLenderStoredBankSortCode
        End If
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankAccountNumEncrypt_Click(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumEncrypt.Click
        If Not _IndividualLenderStoredBankAccountNum.Equals(txtboxIndividualLenderProfileDetailBankAccountNumber.Text) Then
            txtboxIndividualLenderProfileDetailBankAccountNumber.Text = _IndividualLenderStoredBankAccountNum
        End If
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankAccountNumEncrypt_MouseHover(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumEncrypt.MouseHover
        Cursor = Cursors.Hand
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankAccountNumEncrypt_MouseLeave(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankAccountNumEncrypt.MouseLeave
        Cursor = Cursors.Default
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankSortCodeEncrypt_MouseHover(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeEncrypt.MouseHover
        Cursor = Cursors.Hand
    End Sub

    Private Sub buttonIndividualLenderProfileDetailBankSortCodeEncrypt_MouseLeave(sender As Object, e As EventArgs) Handles buttonIndividualLenderProfileDetailBankSortCodeEncrypt.MouseLeave
        Cursor = Cursors.Default
    End Sub
    Public Function GetServerName() As String
        Dim s, strConn As String

        strConn = System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString 'System.Configuration.ConfigurationManager.ConnectionStrings("FBConnectionString").ConnectionString.ToUpper
        strConn = strConn.ToLower()
        s = ""
        If strConn.ToLower.Contains("testing") Then
            s = "Testing Server"
        End If

        If s = "" Then
            If strConn.ToLower.Contains("dev") Then
                s = "Dev Server"
            End If
        End If

        If s = "" Then
            If strConn.ToLower.Contains("finaltest") Then
                s = "User Acceptance Server"
            End If
        End If

        If s = "" Then
            If strConn.ToLower.Contains("shadow") Then
                s = "Shadow Server"
            End If
        End If
        If s = "" Then
            If strConn.ToLower.Contains("uat") Then
                s = "UAT Server"
            End If
        End If
        If s = "" Then
            If strConn.ToLower.Contains("main") Then
                s = "Live Server"
            End If
        End If

        GetServerName = s
    End Function

End Class
