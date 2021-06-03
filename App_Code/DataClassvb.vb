Imports System.Data.SqlClient
Public Class DataClassvb
    Implements IDisposable

    Private MySQL, strConn As String

    Public dsData As New DataSet


    Public CurrUser As UserDetails

    <Serializable>
    Public Structure SessionItems
        Public Property Key As String
        Public Property Value As Object
    End Structure

    <Serializable>
    Public Structure UserDetails
        'Stage 1 Data
        Public Property UserName As String
        Public Property Email As String
        Public Property Password As String
        Public Property FirstName As String
        Public Property LastName As String
        'Stage 3 Data
        Public Property Address1 As String
        Public Property Address2 As String
        Public Property Address3 As String
        Public Property Address4 As String
        Public Property Town As String
        Public Property County As String
        Public Property Country As String
        Public Property PostCode As String
        Public Property DateofBirth As String
        'Stage 5 Data
        Public Property BankSortCode As String
        Public Property BankAccountNumber As String
        Public Property BankAccountName As String
        Public Property UserImageURL As String
        Public Property USERID As Integer
        Public Property UserType As Integer 'UserType == 0 - Lender  1 - Borrower
        Public Property ActivationState As Integer
        Public Property IndividOrg As String
        'ActivationState
        '===============
        '1 - Signed up and verification email sent  
        '2 - Verified Email  
        '3 - AML Details entered  
        '4 - AML Successfully Run  
        '5 - BANK VER Entered  
        '6 - Bank VER Succeeded
    End Structure

    <Serializable>
    Public Structure ResaleDetails
        Public Property CompanyName As String
        Public Property CoRef As String
        Public Property TermRemaining As String
        Public Property Amount As String
        Public Property TotalAmount As String
        Public Property TotalForLoan As String
        Public Property Premium As String
        Public Property StarRating As String
        Public Property TotalFunds As String
        Public Property TermLong As String
        Public Property EncryptLoanID As String
        Public Property EncryptOrderID As String
        Public Property LoanID As String
        Public Property Background As String
        Public Property Security As String
        Public Property SecurityYesNo As String
        Public Property Purpose As String
        Public Property KeyPersonel As String
        Public Property OtherInfo As String
        Public Property QandA As String
        Public Property EndDate As String
        Public Property RevisedRate As String
        Public Property PitchURL As String
        Public Property URL As String
        Public Property ImageURL As String
        Public Property AccruedInterest As String
        Public Property LoanType As Integer
        Public Property StartDate As Date
        Public Property EarlyRedemption As Integer
        Public Property EarlyRedemptionDate As String




    End Structure

    <Serializable>
    Public Structure CompanyDetails
        Public Property CompanyName As String
        Public Property BidRef As String
        Public Property TotalInvestment As String
        Public Property TermShort As String
        Public Property StarRating As String
        Public Property StartDate As String
        Public Property EndDate As String
        Public Property LastDate As String
        Public Property TheOrder As String
        Public Property TheLoan As String
        Public Property myrate As String
        Public Property lowrate As String
        Public Property AverageRate As String
        Public Property RaisedSoFar As String
        Public Property Amount As String
        Public Property TotalFunds As String
        Public Property TotalForLoan As String
        Public Property MaturityDate As String
        Public Property PitchURL As String
        Public Property URL As String
        Public Property ImageURL As String
        Public Property Website As String
        Public Property Background As String
        Public Property FundsRequired As String
        Public Property TermLong As String
        Public Property Security As String
        Public Property SecurityYesNo As String
        Public Property ClosingDate As String
        Public Property Purpose As String
        Public Property KeyPersonel As String
        Public Property OtherInfo As String
        Public Property QandA As String
        Public Property TimeLeftDays As String
        Public Property TimeLeftHours As String
        Public Property EncryptLoanID As String
        Public Property OutstandingBalance As String
        Public Property InterestRate As String
        Public Property IFISAInterestRate As String
        Public Property InterestReceived As String
        Public Property InterestBought As String
        Public Property lhBalsID As String
        Public Property AccruedInterest As String
        Public Property LoanStatus As Integer
        Public Property LoanType As Integer
        Public Property LoanID As Integer
        Public Property CoRef As String
        Public Property TermRemainingDaysInt As Integer
        Public Property TermRemaining As String
        Public Property SaleAmount As String
        Public Property Premium As String
        Public Property PremiumPercent As String
        Public Property TotalAmount As String
        Public Property CancelURL As String
        Public Property RevisedRate As String
        Public Property InterestType As String
        Public Property EarlyRedemption As Integer
        Public Property EarlyRedemptionDate As String
        Public Property SecondaryTrading As Integer
        Public Property ParentLoanID As String
        Public Property FacilityAmount As String
        Public Property Total As String
        Public Property Suspense As String
        Public Property SelfBal As String
        Public Property IFLenderFeeRA As Integer
        Public Property IFLenderFeeIFISA As Integer
        Public Property AER As String
        Public Property Frequency As String
        Public Property AnnualEffective As String

        Public Sub New(dummy As Integer)
            CompanyName = "Uninitialized"
            BidRef = "Uninitialized"
            TotalInvestment = "Uninitialized"
            TermShort = "Uninitialized"
            StarRating = "Uninitialized"
            StartDate = "Uninitialized"
            EndDate = "Uninitialized"
            LastDate = "Uninitialized"
            TheOrder = "Uninitialized"
            TheLoan = "Uninitialized"
            myrate = "Uninitialized"
            lowrate = "Uninitialized"
            AverageRate = "Uninitialized"
            RaisedSoFar = "Uninitialized"
            Amount = "Uninitialized"
            TotalFunds = "Uninitialized"
            TotalForLoan = "Uninitialized"
            MaturityDate = "Uninitialized"
            PitchURL = "Uninitialized"
            URL = "Uninitialized"
            ImageURL = "Uninitialized"
            Website = "Uninitialized"
            Background = "Uninitialized"
            FundsRequired = "Uninitialized"
            TermLong = "Uninitialized"
            Security = "Uninitialized"
            SecurityYesNo = "Uninitialized"
            ClosingDate = "Uninitialized"
            Purpose = "Uninitialized"
            KeyPersonel = "Uninitialized"
            OtherInfo = "Uninitialized"
            QandA = "Uninitialized"
            TimeLeftDays = "Uninitialized"
            TimeLeftHours = "Uninitialized"
            EncryptLoanID = "Uninitialized"
            OutstandingBalance = "Uninitialized"
            InterestRate = "Uninitialized"
            IFISAInterestRate = "Uninitialized"
            InterestReceived = "Uninitialized"
            InterestBought = "Uninitialized"
            lhBalsID = "Uninitialized"
            AccruedInterest = "Uninitialized"
            LoanStatus = Integer.MinValue
            LoanType = Integer.MinValue
            LoanID = Integer.MinValue
            CoRef = "Uninitialized"
            TermRemainingDaysInt = Integer.MinValue
            TermRemaining = "Uninitialized"
            SaleAmount = "Uninitialized"
            Premium = "Uninitialized"
            PremiumPercent = "Uninitialized"
            TotalAmount = "Uninitialized"
            CancelURL = "Uninitialized"
            RevisedRate = "Uninitialized"
            InterestType = "Uninitialized"
            EarlyRedemption = Integer.MinValue
            EarlyRedemptionDate = "Uninitialized"
            SecondaryTrading = Integer.MinValue
            ParentLoanID = "Uninitialized"
            FacilityAmount = "Uninitialized"
            Total = "Uninitialized"
            SelfBal = "Uninitialized"
            IFLenderFeeRA = Integer.MinValue
            IFLenderFeeIFISA = Integer.MinValue
            AER = Integer.MinValue
            Frequency = "Uninitialised"
            AnnualEffective = "Uninitialised"
        End Sub


    End Structure

    Public Enum QueryType
        Selekt = 0
        Insert = 1
        Update = 2
    End Enum

    Sub Dispose() Implements IDisposable.Dispose

    End Sub




    Public Function GetData(SQLStr As String) As DataSet

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                Dim returnData As New DataSet
                Dim SQl1 As String = SQLStr.ToUpper.Replace("CURRENT_DATE", "GETDATE()")
                Dim SQL2 As String = SQl1.ToUpper.Replace("FIRST", "TOP")
                Dim cmd As SqlCommand = New SqlCommand(SQL2, con)
                con.Open()

                adapter.SelectCommand = cmd

                returnData = New DataSet
                adapter.Fill(returnData)
                Return returnData

            Catch ex As Exception
            Finally

            End Try
        End Using


    End Function








End Class

