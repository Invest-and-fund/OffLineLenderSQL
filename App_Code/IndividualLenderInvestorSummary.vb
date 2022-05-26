Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports System.Data.SqlClient
Public Class IndividualLenderInvestorSummary
    Public Shared _LenderTotalLent As String
    Public Shared _LenderTotalFundBalance As String
    Public Shared _LenderTotalCurrentLoans As String
    Public Shared _LenderTotalGrossInterestEarned As String
    Public Shared _LenderAmountAvailableToLend As String
    Public Shared _LenderBidsOutstanding As String
    Public Shared _LenderInterestAccrued As String
    Public Shared _LenderGROSS_YIELD As String
    Public Shared _LenderDATE_JOINED As String
    Public Shared _LenderCURRENT_INVESTMENTS As String
    Public Shared _LenderAMOUNT_AVAIL As String
    Public Shared _LenderLIVE_AUCTION_BIDS As String
    Public Shared _LenderTOTAL_FUNDS_BAL As String
    Public Shared _LenderFACILITY_FEES_TOTAL As String
    Public Shared _LenderGROSS_INTEREST_EARNED As String
    Public Shared _LenderACE As String
    Public Shared _LenderACCRUED_SOLD As String
    Public Shared _LenderACCRUED_BOUGHT As String
    Public Shared _LenderNO_ACTIVE_INVESTMENTS As String
    Public Shared _LenderTRANSACTION_FEES_TOTALL As String


    Public Shared Function GetIndividualLenderIS(ByVal AccountID As String, ByVal UserID As String)

        Dim FlagReturn As Boolean

        FlagReturn = False


        Dim strIndividualLenderProfileData = New StringBuilder(512)

        Try
            If Not String.IsNullOrWhiteSpace(AccountID) AndAlso Not String.IsNullOrWhiteSpace(UserID) Then

                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Using cmd As New SqlCommand("INVESTOR_SUMMARYNew")

                            cmd.Parameters.Add("@O_GROSS_YIELD", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_DATE_JOINED", SqlDbType.DateTime).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_CURRENT_INVESTMENTS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_AMOUNT_AVAIL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_LIVE_AUCTION_BIDS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_TOTAL_FUNDS_BAL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_FACILITY_FEES_TOTAL", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_GROSS_INTEREST_EARNED", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACE", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACCRUED_SOLD", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_ACCRUED_BOUGHT", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_NO_ACTIVE_INVESTMENTS", SqlDbType.Int).Direction = ParameterDirection.Output
                            cmd.Parameters.Add("@O_TRANSACTION_FEES_TOTAL", SqlDbType.Int).Direction = ParameterDirection.Output

                            With cmd.Parameters
                                .Add(New SqlParameter("@I_ACC_ID", AccountID))
                                .Add(New SqlParameter("@I_USER_ID", UserID))

                            End With
                            con.Open()
                            cmd.Connection = con
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandTimeout = 0
                            cmd.ExecuteNonQuery()


                            _LenderGROSS_YIELD = GenDB.fnDBStringField(cmd.Parameters("@O_GROSS_YIELD").Value)
                            _LenderDATE_JOINED = GenDB.fnDBStringField(cmd.Parameters("@O_DATE_JOINED").Value)
                            _LenderCURRENT_INVESTMENTS = GenDB.fnDBStringField(cmd.Parameters("@O_CURRENT_INVESTMENTS").Value)
                            _LenderAMOUNT_AVAIL = GenDB.fnDBStringField(cmd.Parameters("@O_AMOUNT_AVAIL").Value)
                            _LenderLIVE_AUCTION_BIDS = GenDB.fnDBStringField(cmd.Parameters("@O_LIVE_AUCTION_BIDS").Value)
                            _LenderTOTAL_FUNDS_BAL = GenDB.fnDBStringField(cmd.Parameters("@O_TOTAL_FUNDS_BAL").Value)
                            _LenderFACILITY_FEES_TOTAL = GenDB.fnDBStringField(cmd.Parameters("@O_FACILITY_FEES_TOTAL").Value)
                            _LenderGROSS_INTEREST_EARNED = GenDB.fnDBStringField(cmd.Parameters("@O_GROSS_INTEREST_EARNED").Value)
                            _LenderACE = GenDB.fnDBStringField(cmd.Parameters("@O_ACE").Value)
                            _LenderACCRUED_SOLD = GenDB.fnDBStringField(cmd.Parameters("@O_ACCRUED_SOLD").Value)
                            _LenderACCRUED_BOUGHT = GenDB.fnDBStringField(cmd.Parameters("@O_ACCRUED_BOUGHT").Value)
                            _LenderNO_ACTIVE_INVESTMENTS = GenDB.fnDBStringField(cmd.Parameters("@O_NO_ACTIVE_INVESTMENTS").Value)
                            _LenderTRANSACTION_FEES_TOTALL = GenDB.fnDBStringField(cmd.Parameters("@O_TRANSACTION_FEES_TOTAL").Value)

                            _LenderTotalLent = _LenderNO_ACTIVE_INVESTMENTS
                            _LenderTotalFundBalance = _LenderTOTAL_FUNDS_BAL
                            _LenderTotalCurrentLoans = _LenderCURRENT_INVESTMENTS
                            _LenderTotalGrossInterestEarned = _LenderGROSS_INTEREST_EARNED
                            _LenderAmountAvailableToLend = _LenderAMOUNT_AVAIL
                            _LenderBidsOutstanding = _LenderLIVE_AUCTION_BIDS
                            _LenderInterestAccrued = _LenderACCRUED_BOUGHT + _LenderACCRUED_SOLD

                            FlagReturn = True
                        End Using
                    Catch ex As Exception

                        FlagReturn = False
                    Finally

                    End Try
                End Using
            End If


        Catch ex As Exception

            Throw
        Finally


            strIndividualLenderProfileData.Length = 0
        End Try

        Return FlagReturn

    End Function
End Class
