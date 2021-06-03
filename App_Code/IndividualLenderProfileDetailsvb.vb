Imports System.Data.SqlClient
Public Class IndividualLenderProfileDetails2
    Public Property LenderName As String
    Public Property LenderAccountId As String
    Public Property LenderAccountType As String
    Public Property LenderUserId As Integer
    Public Property LenderAccountReference As String
    Public Property LenderEmailAddress As String
    Public Property LenderBankAccountNumber As String
    Public Property LenderBankSortCode As String
    Public Property LenderSearchValue As String
    Public Property LenderSearchFilter As String
    Public Property LenderType As String
    Public Property LenderCategory As String
    Public Property LenderActivationType As String
    Public Property LenderActivatedCert As String
    Public Property LenderBankActivationType As String
    Public Property LenderCreatedDate As String
    Public Property LenderEmailActivation As String
    Public Property LenderForeName As String
    Public Property LenderSurname As String
    Public Property LenderDOB As String
    Public Property LenderPhoneNumber As String
    Public Property LenderFirstEmailSentOn As String
    Public Property LenderBuildingSocietyRollNum As String
    Public Property LenderHouseNameNum As String
    Public Property LenderAddress1 As String
    Public Property LenderAddress2 As String
    Public Property LenderAddress3 As String
    Public Property LenderAddress4 As String
    Public Property LenderTownCity As String
    Public Property LenderCounty As String
    Public Property LenderCountry As String
    Public Property LenderPostCode As String
    Public Property LenderBankAccountName As String
    Public Property LenderSeqAlpha As String
    Public Property IndividualLenderCurrentLoansHasRows As Boolean
    Public Property ListOfIndividualLendersByLastNameFirstNameUid As String
    Public Property ListOfIndividualLendersByFirstNameLastNameUid As String
    Public Property ListOfIndividualLendersByUidLastNameFirstName As String
    Public Property ListOfIndividualLendersByAccountId As String
    Public Property ListOfIndividualLendersAccountId As String
    Public Property IndividualLenderFilterNames As Dictionary(Of String, String)
    Public Property IndividualLenderProfileDataTable As DataTable
    Public Property IndividualLenderTradingHistorysHasRows As Boolean
    Public Property IndividualLenderCategorisationHistoryHasRows As Boolean
    Public Property LenderCategoryAmendedTo As Integer
    Public Property LenderForeNameAmendedTo As String
    Public Property LenderSurnameAmendedTo As String
    Public Property LenderDOBAmendedTo As String
    Public Property LenderEmailAddressAmendedTo As String
    Public Property LenderPhoneNumberAmendedTo As String
    Public Property LenderHouseNameNumAmendedTo As String
    Public Property LenderAddress1AmendedTo As String
    Public Property LenderAddress2AmendedTo As String
    Public Property LenderAddress3AmendedTo As String
    Public Property LenderAddress4AmendedTo As String
    Public Property LenderTownCityAmendedTo As String
    Public Property LenderCountyAmendedTo As String
    Public Property LenderCountryAmendedTo As String
    Public Property LenderPostCodeAmendedTo As String
    Public Property LenderFullAddress As String
    Public Property IndividualLenderIsSelected As Boolean
    Public Property lastLentDetails As Date
    Public Property lastLentLoanRef As Integer
    Public Property IndividualLenderLastLentDetailsLoanName As String
    Public Property IndividualLenderLastLentDetailsLoanDateTime As Date
    Public Property IndividualLenderLastLentDetailsLoanID As Integer
    Public Property IndividualLenderLastLentDetailsLoanAmount As Integer
    Public Property IndividualLenderLastLentDetailsNextMaturityDateTime As Date
    Public Const STR_LenderFindByText_UserID As String = "User ID"
    Public Const STR_LenderFindByText_LastName As String = "Last Name"
    Public Const STR_LenderFindByText_FirstName As String = "First Name"
    Public Const STR_LenderFindByText_CompanyName As String = "Company Name"
    Public Const STR_LenderFindByText_AccountId As String = "Account ID"


    Shared Function SafeString(S As Object) As String
        Dim RV As String = ""

        If S IsNot DBNull.Value Then RV = Trim(S.ToString)

        Return RV
    End Function

    Shared Function GetDate(ByVal strDt As Object) As String
        Dim dt1 As DateTime

        If DateTime.TryParse(strDt.ToString(), dt1) Then
            Return dt1
        Else
            Return ""
        End If
    End Function


    Public Shared Function CreateIndividualLenderFilterListUID() As List(Of IndividualLenderProfileDetails2)
        Dim myList As New List(Of IndividualLenderProfileDetails2)
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Dim cmd As New SqlCommand()
            Dim Rdr As SqlDataReader

            Try
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Connection = con
                con.Open()
                cmd.CommandText = "GET_ALLUSERSNEW"
                Rdr = cmd.ExecuteReader
                While Rdr.Read
                    Dim myItem As New IndividualLenderProfileDetails2 With {
                     .ListOfIndividualLendersByUidLastNameFirstName = SafeString(Rdr("thenameandUserId")),
                    .ListOfIndividualLendersByAccountId = SafeString(Rdr("thenameandUserId"))
                    }
                    myList.Add(myItem)
                End While
            Catch ex As Exception
            Finally

            End Try
        End Using
        Return myList
    End Function

    Public Shared Function GetIndividualLenderCurrentLoans(ByVal AccountID As Integer) As List(Of IndividualLenderProfileDetails2)
        Dim myList As New List(Of IndividualLenderProfileDetails2)
        Dim result As DataTable = New DataTable()
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Dim cmd As New SqlCommand()
            Dim Rdr As SqlDataReader

            Try
                cmd.CommandType = CommandType.StoredProcedure
                With cmd.Parameters
                    .Add(New SqlParameter("@AccountID", AccountID))

                End With
                cmd.Connection = con
                con.Open()
                cmd.CommandText = "GetIndividualLenderCurrentLoans"
                Rdr = cmd.ExecuteReader
                While Rdr.Read
                    Dim myItem As New IndividualLenderProfileDetails2 With {
                    .IndividualLenderLastLentDetailsLoanDateTime = GetDate(Rdr("StartDate")),
                    .IndividualLenderLastLentDetailsLoanID = SafeString(Rdr("TheLoan")),
                    .IndividualLenderLastLentDetailsLoanAmount = GenDB.PenceToCurrencyStringPounds(Rdr("outstanding")),
                    .IndividualLenderLastLentDetailsNextMaturityDateTime = GetDate(Rdr("LastDate"))
                      }
                    myList.Add(myItem)
                End While
            Catch ex As Exception
            Finally

            End Try
        End Using
        Return myList
    End Function
    Public Shared Function GetIndividualLenderCurrentLoansDS(ByVal AccountID As Integer)

        Dim sErrorStr As String = ""
        Dim ds As DataSet = New DataSet("IndividualLenderProfileDetails")

        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Dim cmd As SqlCommand = New SqlCommand("GetIndividualLenderCurrentLoans", con)
                With cmd.Parameters
                    .Add(New SqlParameter("@AccountID", AccountID))

                End With


                cmd.CommandType = CommandType.StoredProcedure
                Dim da As SqlDataAdapter = New SqlDataAdapter()
                da.SelectCommand = cmd
                da.Fill(ds)
            Catch ex As Exception
            Finally

            End Try
        End Using
        Return ds
    End Function
    Public Shared Function GetIndividualLenderCurrentLoansDT(ByVal AccountID As Integer)
        Dim dt As New DataTable()
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Try
                Using cmd As New SqlCommand("GetIndividualLenderCurrentLoans")
                    With cmd.Parameters
                        .Add(New SqlParameter("@AccountID", AccountID))

                    End With
                    cmd.Connection = con
                    cmd.CommandType = CommandType.StoredProcedure
                    Using sda As New SqlDataAdapter(cmd)

                        sda.Fill(dt)
                    End Using
                End Using
            Catch ex As Exception
            Finally

            End Try
        End Using
        Return dt
    End Function
End Class
