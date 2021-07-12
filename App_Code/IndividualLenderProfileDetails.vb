Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports System.Data.SqlClient

Public Class IndividualLenderProfileDetails
    Public Shared _LenderName As String
    Public Shared _LenderAccountId As String
    Public Shared _LenderAccountType As String
    Public Shared _LenderUserId As Integer
    Public Shared _LenderAccountReference As String
    Public Shared _LenderEmailAddress As String
    Public Shared _LenderBankAccountNumber As String
    Public Shared _LenderBankSortCode As String
    Public Shared _LenderSearchValue As String
    Public Shared _LenderSearchFilter As String
    Public Shared _LenderType As String
    Public Shared _LenderCategory As String
    Public Shared _LenderActivationType As String
    Public Shared _LenderActivatedCert As String
    Public Shared _LenderBankActivationType As String
    Public Shared _LenderCreatedDate As String
    Public Shared _LenderEmailActivation As String
    Public Shared _LenderForeName As String
    Public Shared _LenderSurname As String
    Public Shared _LenderDOB As String
    Public Shared _LenderPhoneNumber As String
    Public Shared _LenderFirstEmailSentOn As String
    Public Shared _LenderBuildingSocietyRollNum As String
    Public Shared _LenderHouseNameNum As String
    Public Shared _LenderAddress1 As String
    Public Shared _LenderAddress2 As String
    Public Shared _LenderAddress3 As String
    Public Shared _LenderAddress4 As String
    Public Shared _LenderTownCity As String
    Public Shared _LenderCounty As String
    Public Shared _LenderCountry As String
    Public Shared _LenderPostCode As String
    Public Shared _LenderBankAccountName As String
    Public Shared _LenderSeqAlpha As String
    Public Shared _ListOfIndividualLendersByLastNameFirstNameUid As EnumerableRowCollection(Of String)
    Public Shared _ListOfIndividualLendersByFirstNameLastNameUid As EnumerableRowCollection(Of String)
    Public Shared _ListOfIndividualLendersByUidLastNameFirstName As EnumerableRowCollection(Of String)
    Public Shared _ListOfIndividualLendersByUIDCompanyName As EnumerableRowCollection(Of String)
    Public Shared _ListOfIndividualLendersByAccountId As EnumerableRowCollection(Of String)
    Public Shared _ListOfIndividualLendersAccountId As EnumerableRowCollection(Of String)
    Public Shared _IndividualLenderFilterNames As Dictionary(Of String, String)
    Public Shared _IndividualLenderProfileDataTable As DataTable
    Public Shared _IndividualLenderTradingHistorysHasRows As Boolean
    Public Shared _IndividualLenderCategorisationHistoryHasRows As Boolean
    Public Shared _LenderCategoryAmendedTo As Integer
    Public Shared _LenderForeNameAmendedTo As String
    Public Shared _LenderSurnameAmendedTo As String
    Public Shared _LenderDOBAmendedTo As String
    Public Shared _LenderEmailAddressAmendedTo As String
    Public Shared _LenderPhoneNumberAmendedTo As String
    Public Shared _LenderHouseNameNumAmendedTo As String
    Public Shared _LenderAddress1AmendedTo As String
    Public Shared _LenderAddress2AmendedTo As String
    Public Shared _LenderAddress3AmendedTo As String
    Public Shared _LenderAddress4AmendedTo As String
    Public Shared _LenderTownCityAmendedTo As String
    Public Shared _LenderCountyAmendedTo As String
    Public Shared _LenderCountryAmendedTo As String
    Public Shared _LenderPostCodeAmendedTo As String
    Public Shared _LenderFullAddress As String
    Public Shared _IndividualLenderIsSelected As Boolean
    Public Const STR_LenderFindByText_UserID As String = "User ID"
    Public Const STR_LenderFindByText_LastName As String = "Last Name"
    Public Const STR_LenderFindByText_FirstName As String = "First Name"
    Public Const STR_LenderFindByText_CompanyName As String = "Company Name"
    Public Const STR_LenderFindByText_AccountId As String = "Account ID"
    Private _Userid As Integer
    Private _Accountid As Integer

    Private _Activated As Integer

    Private _Thename As String
    Private _Individorg As Integer

    Public Property Accountid() As Integer
        Get
            Return _Accountid
        End Get
        Set(ByVal value As Integer)
            _Accountid = value
        End Set
    End Property

    Public Property Userid() As Integer
        Get
            Return _Userid
        End Get
        Set(ByVal value As Integer)
            _Userid = value
        End Set
    End Property

    Public Property Thename() As Integer
        Get
            Return _Thename
        End Get
        Set(ByVal value As Integer)
            _Thename = value
        End Set
    End Property

    Shared Function SafeString(S As Object) As String
        Dim RV As String = ""

        If S IsNot DBNull.Value Then RV = Trim(S.ToString)

        Return RV
    End Function

    Public Shared Function GetLenderAddress() As String
        _LenderFullAddress = $"{_LenderHouseNameNum.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderAddress1.Replace("N/A", String.Empty)} {_LenderAddress2.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderAddress3.Replace("N/A", String.Empty)} {_LenderAddress4.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderTownCity.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderCounty.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderCountry.Replace("N/A", String.Empty)} {Environment.NewLine}{_LenderPostCode.Replace("N/A", String.Empty)} "
        _LenderFullAddress.Replace("N/A", String.Empty)
        ' _LenderFullAddress = IAFString.RemoveEmptyLines(_LenderFullAddress)  **********Check Later
        _LenderFullAddress = _LenderFullAddress
        Return _LenderFullAddress
    End Function

    Public Shared Function PopulateIndividualLenderSearchList(ByVal selectedFilterByUser As String)

        Dim cnn As IDbConnection = Nothing
        Dim strSqlIndividualLenderSearchList = New StringBuilder()

        Try
            strSqlIndividualLenderSearchList.Append($"SELECT u.userid,        a.accountid,     
   Concat( a.accountid         , '-'         , 
   CASE WHEN a.ACCOUNTTYPE = 0 THEN 'Standard'              
   WHEN a.ACCOUNTTYPE = 1 THEN 'SIPP'          
   WHEN a.ACCOUNTTYPE = 2 THEN 'IFISA'          
   WHEN a.ACCOUNTTYPE = 3 THEN 'SSAS'           
   WHEN a.ACCOUNTTYPE >= 4 THEN 'BadValue'            END    ) 
   AS AccountIdAccountRef,  
   ( Concat(u.userid           , '-'           , 
   Trim(Replace (Replace(u.lastname, ',', ' '), '-', ' '))          
    , '-'           , Trim(Replace (Replace(u.firstname, ',', ' '), '-', ' '))      
   , '-'           , a.accountid )    )  
   AS UidLastNameFirstName,        ( Concat(a.accountid           , '-'       
   , Trim(Replace (Replace(u.lastname, ',', ' '), '-', ' '))         
   , '-'           , Trim(Replace (Replace(u.firstname, ',', ' '), '-', ' '))    
   , '-'          ,u.userid ))    AS AccountIdLastNameFirstNameUId,        
   ( ConCat(Trim(Replace (Replace(u.lastname, ',', ' '), '-', ' '))       
   , '-'           , Trim(Replace (Replace(u.firstname, ',', ' '), '-', ' '))      
   , '-'           , a.accountid )    )                                    
   AS LastNameFirstNameUid,     
   ( Concat(trim(Replace (Replace(u.firstname, ',', ' '), '-', ' '))         
   , '-'           , Trim(Replace (Replace(u.lastname, ',', ' '), '-', ' '))    
   , '-'           , a.accountid )   )                                     
   AS FirstNameLastNameUid,   u.IsActive,      u.UserType,      
   ( Concat( trim(COALESCE(u.companyname, u.companynumber)), '-' ,u.userid )) as companyname,       u.companynumber,        
   u.individorg,      u.isactive,         u.profclient,        
   u.activated,    Trim(u.lastname)                                     
   AS LastName,         Trim(u.firstname)                                    
   AS FirstName  FROM   users u         INNER JOIN accounts a              
   ON u.userid = a.userid  WHERE  ( ( u.usertype = 0 )          
           AND ( a.activated_bank = 5 )     
   AND ( u.activated = 5 )           AND ( u.isactive = 0 )         
   AND ( u.userid IS NOT NULL )           AND ( a.accountid IS NOT NULL )     
   AND ( u.lastname IS NOT NULL )           AND ( u.firstname IS NOT NULL )     
   AND ( a.userid IS NOT NULL ) ) 
 ")


            Select Case selectedFilterByUser
                Case STR_LenderFindByText_UserID
                    Dim t1 = Task.Run(Function() strSqlIndividualLenderSearchList.Append(String.Format("ORDER  BY u.{0}  {1} ", selectedFilterByUser.Replace(" ", String.Empty), Environment.NewLine)))
                    t1.Wait()
                Case STR_LenderFindByText_LastName
                    Dim t2 = Task.Run(Function() strSqlIndividualLenderSearchList.Append(String.Format("ORDER  BY u.{0}  {1} ", selectedFilterByUser.Replace(" ", String.Empty), Environment.NewLine)))
                    t2.Wait()
                Case STR_LenderFindByText_FirstName
                    Dim t3 = Task.Run(Function() strSqlIndividualLenderSearchList.Append(String.Format("ORDER  BY u.{0}  {1} ", selectedFilterByUser.Replace(" ", String.Empty), Environment.NewLine)))
                    t3.Wait()
                Case STR_LenderFindByText_AccountId
                    Dim t4 = Task.Run(Function() strSqlIndividualLenderSearchList.Append(String.Format("ORDER  BY a.{0}  {1} ", selectedFilterByUser.Replace(" ", String.Empty), Environment.NewLine)))
                    t4.Wait()

                Case STR_LenderFindByText_CompanyName
                    Dim t5 = Task.Run(Function() strSqlIndividualLenderSearchList.Append(String.Format("and u.userid = a.accountid and u.individorg = 0 ORDER  BY u.{0}  {1} ", selectedFilterByUser.Replace(" ", String.Empty), Environment.NewLine)))
                    t5.Wait()
            End Select



            Dim sSQL As String
            Dim sErrorStr As String = ""

            Dim ds = New DataSet
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                    sSQL = strSqlIndividualLenderSearchList.ToString()


                    Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                    con.Open()

                    adapter.SelectCommand = cmd

                    adapter.Fill(ds)


                Catch ex As Exception
                Finally

                End Try
            End Using
            Dim dtcount As Integer
            dtcount = ds.Tables(0).Rows.Count


            Dim t As Task(Of EnumerableRowCollection(Of DataRow)) = Nothing
            Dim selectedFilterByUserColumnName As String = selectedFilterByUser.Replace(" ", String.Empty)

            t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
                                                                                  ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
                                                                                  Return ds.Tables(0).AsEnumerable()
                                                                              End Function)

            'Select Case selectedFilterByUser
            '    Case STR_LenderFindByText_UserID
            '        t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
            '                                                                              ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
            '                                                                              Return ds.Tables(0).AsEnumerable()
            '                                                                          End Function)
            '    Case STR_LenderFindByText_LastName
            '        t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
            '                                                                              ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
            '                                                                              Return ds.Tables(0).AsEnumerable()
            '                                                                          End Function)
            '    Case STR_LenderFindByText_FirstName
            '        t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
            '                                                                              ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
            '                                                                              Return ds.Tables(0).AsEnumerable()
            '                                                                          End Function)
            '    Case STR_LenderFindByText_AccountId
            '        t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
            '                                                                              ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
            '                                                                              Return ds.Tables(0).AsEnumerable()
            '                                                                          End Function)
            '    Case STR_LenderFindByText_CompanyName
            '        t = Task(Of EnumerableRowCollection(Of DataRow)).Factory.StartNew(Function()
            '                                                                              ds.Tables(0).DefaultView.Sort = selectedFilterByUserColumnName
            '                                                                              Return ds.Tables(0).AsEnumerable()
            '                                                                          End Function)
            'End Select

            _ListOfIndividualLendersByUIDCompanyName = From lend In t.Result Select lend.Field(Of String)("CompanyName")
            _ListOfIndividualLendersByLastNameFirstNameUid = From lend In t.Result Select lend.Field(Of String)("LastNameFirstNameUid")
            _ListOfIndividualLendersByFirstNameLastNameUid = From lend In t.Result Select lend.Field(Of String)("FirstNameLastNameUid")
            _ListOfIndividualLendersByUidLastNameFirstName = From lend In t.Result Select lend.Field(Of String)("UidLastNameFirstName")
            ' _ListOfIndividualLendersByAccountId = From lend In t.Result Order By lend.Field(Of Integer)("AccountID") Select (lend.Field(Of String)("AccountIdLastNameFirstNameUId"))
            _ListOfIndividualLendersByAccountId = From lend In t.Result Select lend.Field(Of String)("AccountIdLastNameFirstNameUId")
            _ListOfIndividualLendersAccountId = From lend In t.Result Order By lend.Field(Of Integer)("AccountID") Where lend.Field(Of Integer)("userid") = IndividualLenderProfileDetails._LenderUserId Select (lend.Field(Of String)("AccountIdAccountRef"))

        Catch ex As Exception

            If ex.Message.Equals("Exception in GetDataReader Method") Then

            End If

            If ex.InnerException.Message.StartsWith("Can't Read the Connection String Name:") Then

            Else

                If ex.InnerException.Message.StartsWith("Dynamic SQL Error" & vbCrLf & "SQL error code =") Then

                End If
            End If


        Finally

            If (cnn IsNot Nothing) AndAlso cnn.State.Equals(ConnectionState.Open) Then
                cnn.Close()
            End If

            strSqlIndividualLenderSearchList.Length = 0
        End Try

    End Function
    Public Shared Function CreateIndividualLenderFilterList()

        _IndividualLenderFilterNames = New Dictionary(Of String, String)()
        Dim t = Task.Run(Sub() _IndividualLenderFilterNames.Add("Last Name", "USERS.lastname"))
        t.Wait()
        Dim t1 = Task.Run(Sub() _IndividualLenderFilterNames.Add("First Name", "USERS.firstname"))
        t1.Wait()
        Dim t2 = Task.Run(Sub() _IndividualLenderFilterNames.Add("User ID", "USERS.userid"))
        t2.Wait()
        Dim t3 = Task.Run(Sub() _IndividualLenderFilterNames.Add("Account ID", "ACCOUNTS.userid"))
        t3.Wait()
        Dim t4 = Task.Run(Sub() _IndividualLenderFilterNames.Add("Company Name", "USERS.Companyname"))
        t4.Wait()
    End Function

    Public Shared Function GetIndividualLenderSearchValue(ByVal strSearchValue As String, ByVal individualLenderFilter As String) As String
        Dim userId = String.Empty

        If Not String.IsNullOrWhiteSpace(strSearchValue) AndAlso Not String.IsNullOrWhiteSpace(individualLenderFilter) Then
            Dim searchValues = strSearchValue.Split("-"c)

            Select Case individualLenderFilter
                Case "User ID"
                    Dim notValidADummyIdEntered As Integer
                    userId = searchValues(0)

                    If Not Integer.TryParse(userId.Substring(0, 1), notValidADummyIdEntered) Then
                        userId = String.Format("""Users.Userid,{0}""", searchValues(2))
                        Exit Select
                    End If

                Case "Last Name"
                    userId = searchValues(2)
                Case "First Name"
                    userId = searchValues(2)
                Case "Account ID"
                    userId = searchValues(3)
            End Select
        End If

        IndividualLenderProfileDetails._LenderAccountId = userId
        IndividualLenderProfileDetails._LenderUserId = Convert.ToInt16(userId)
        Return userId
    End Function


    Public Shared Function GetLastNameFirstnameAccID() As List(Of IndividualLenderProfileDetails)
        Dim myList As New List(Of IndividualLenderProfileDetails)
        Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
            Dim cmd As New SqlCommand()
            Dim Rdr As SqlDataReader

            Try
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Connection = con
                con.Open()
                cmd.CommandText = "GET_ALLUSERS"
                Rdr = cmd.ExecuteReader


                While Rdr.Read
                    Dim myItem As New IndividualLenderProfileDetails With {
                    ._Userid = GenDB.fnDBIntField(Rdr("Userid")),
                    ._Accountid = SafeString(Rdr("Accountid")),
                    ._Thename = SafeString(Rdr("Thename"))
                     }
                    myList.Add(myItem)
                End While

            Catch ex As Exception
            Finally

            End Try
        End Using
        Return myList
    End Function

    Public Shared Function GetIndividualLenderProfileData(ByVal searchBySqlField As String, ByVal searchValueFromSelectedByUser As String) As DataTable

        Dim result = New DataTable()

        Dim strIndividualLenderProfileData = New StringBuilder(512)

        Try

            If Not String.IsNullOrWhiteSpace(searchBySqlField) AndAlso Not String.IsNullOrWhiteSpace(searchValueFromSelectedByUser) Then
                strIndividualLenderProfileData.Append($"                  SELECT   USERS.USERID  as USERS_USERID,  
                            USERS.ISACTIVE as USERS_ISACTIVE,  
                              CASE WHEN USERS.CREATEDBY IS NOT NULL THEN CONVERT(VARCHAR(100), USERS.CREATEDBY)  
                              ELSE 'N/A'  
                              END AS USERS_CREATEDBY,  
                              USERS.DATECREATED  as   USERS_DATECREATED,  
                              CASE WHEN USERS.FIRSTNAME IS NOT NULL THEN Trim(USERS.FIRSTNAME)  
                              ELSE 'N/A'  
                              END AS USERS_FIRSTNAME,  
                              CASE WHEN USERS.LASTNAME IS NOT NULL THEN Trim(USERS.LASTNAME)  
                              ELSE 'N/A'  
                              END AS USERS_LASTNAME,  
                              CASE WHEN USERS.ADDRESS1 IS NOT NULL THEN Trim(USERS.ADDRESS1)  
                              ELSE 'N/A'  
                              END AS USERS_ADDRESS1,  
                              CASE WHEN USERS.ADDRESS2 IS NOT NULL THEN Trim(USERS.ADDRESS2)  
                              ELSE 'N/A'  
                              END AS USERS_ADDRESS2,  
                              CASE WHEN USERS.ADDRESS3 IS NOT NULL THEN Trim(USERS.ADDRESS3)  
                              ELSE 'N/A'  
                              END AS USERS_ADDRESS3,  
                              CASE WHEN USERS.ADDRESS4 IS NOT NULL THEN Trim(USERS.ADDRESS4)  
                              ELSE 'N/A'  
                              END AS USERS_ADDRESS4,  
                              CASE WHEN USERS.TOWN IS NOT NULL THEN Trim(USERS.TOWN)  
                              ELSE 'N/A'  
                              END AS USERS_TOWN,
							            CASE WHEN USERS.COUNTY IS NOT NULL THEN Trim(USERS.COUNTY)  
                              ELSE 'N/A'  
                              END AS USERS_COUNTY,  
                              CASE WHEN USERS.COUNTRY IS NOT NULL THEN Trim(USERS.COUNTRY)  
                              ELSE 'N/A'  
                              END AS USERS_COUNTRY,  
                      
                              CASE WHEN USERS.POSTCODE IS NOT NULL THEN Trim(USERS.POSTCODE)  
                              ELSE 'N/A'  
                              END AS USERS_POSTCODE,  
                              CASE WHEN USERS.COMPANYNAME IS NOT NULL THEN Trim(USERS.COMPANYNAME)  
                              ELSE 'N/A'  
                              END AS USERS_COMPANYNAME,  
                              CASE WHEN USERS.EMAIL IS NOT NULL THEN Trim(USERS.EMAIL)  
                              ELSE 'N/A'  
                              END AS USERS_EMAIL,  
							  USERS.ACTIVATED  AS USERS_ACTIVATED,  
                              CASE WHEN accounts.ACTIVATED_BANK IS NOT NULL THEN  CONVERT(VARCHAR(100), accounts.ACTIVATED_BANK )  
                              ELSE 'N/A'  
                              END AS USERS_ACTIVATED_BANK,  
                              USERS.ACTIVATIONDATE   AS USERS_ACTIVATIONDATE,  
                              CASE WHEN USERS.ACTIVATED_CERT IS NOT NULL THEN  CONVERT(VARCHAR(100), USERS.ACTIVATED_CERT )  
                              ELSE 'N/A'  
                              END AS USERS_ACTIVATED_CERT,
							                                CASE WHEN USERS.HOWHEAR IS NOT NULL THEN Trim(USERS.HOWHEAR)  
                              ELSE 'N/A'  
                              END AS USERS_HOWHEAR,  
                              CASE WHEN USERS.HOUSENAME IS NOT NULL THEN Trim(USERS.HOUSENAME)  
                              ELSE 'N/A'  
                              END AS USERS_HOUSENAME,  
                              CASE WHEN USERS.STREETNAME IS NOT NULL THEN Trim(USERS.STREETNAME)  
                              ELSE 'N/A'  
                              END AS USERS_STREETNAME,  
                              CASE WHEN USERS.DATEOFBIRTH IS NOT NULL THEN LEFT((USERS.DATEOFBIRTH), 12)  
                              ELSE 'N/A'  
                              END AS USERS_DATEOFBIRTH,  
                              CASE WHEN USERS.SEQALPHA IS NOT NULL THEN CONVERT(VARCHAR(100),USERS.SEQALPHA)  
                              ELSE 'N/A'  
                              END AS USERS_SEQALPHA,
							                                USERS.COMPANYNUMBER  AS USERS_COMPANYNUMBER,  
                              USERS.INDIVIDORG AS  USERS_INDIVIDORG,  
                              (USERS.ORGTYPE)  AS USERS_ORGTYPE,  
							  CASE WHEN USERS.TITLE IS NOT NULL THEN Trim(USERS.TITLE)  
                              ELSE 'N/A'  
                              END AS USERS_TITLE,  
                              CASE WHEN USERS.TELEPHONE IS NOT NULL THEN Trim(USERS.TELEPHONE)  
                              ELSE 'N/A'  
                              END AS USERS_TELEPHONE,  
                              CASE WHEN ACCOUNTS.BANKACCNAME IS NOT NULL THEN Trim(ACCOUNTS.BANKACCNAME)  
                              ELSE 'N/A'  
                              END AS ACCOUNTS_BANKACCNAME,  
                              CASE WHEN ACCOUNTS.BANKACCNUMBER IS NOT NULL THEN Trim(ACCOUNTS.BANKACCNUMBER)  
                              ELSE 'N/A'  
                              END AS ACCOUNTS_BANKACCNUMBER,  
                              CASE WHEN ACCOUNTS.BANKACCSORTCODE IS NOT NULL THEN Trim(ACCOUNTS.BANKACCSORTCODE)  
                              ELSE 'N/A'  
                              END AS ACCOUNTS_BANKACCSORTCODE,  
                              CASE WHEN ACCOUNTS.BUILDSOCROLLNUMBER IS NOT NULL THEN Trim(ACCOUNTS.BUILDSOCROLLNUMBER)  
                              ELSE 'N/A'  
                              END AS ACCOUNTS_BUILDSOCROLLNUMBER,  
                              CASE WHEN USERS.CLIENT_CATEGORISATION IS NOT NULL THEN  
                                  CASE USERS.CLIENT_CATEGORISATION WHEN 0 THEN 'Uninitialized'  
                                    WHEN 1 THEN 'Elective Professional'  
                                    WHEN 2 THEN 'Self Certified Sophisticated'  
                                    WHEN 3 THEN 'High Networth Individual'  
                                    WHEN 4 THEN 'Per Se Professional'  
                                  END  
                              ELSE 'Uninitialized'  
                              END AS CLIENT_CATEGORISATION,  
                              CASE WHEN ACCOUNTS.ACCOUNTTYPE = 0 THEN 'Standard'  
                                WHEN ACCOUNTS.ACCOUNTTYPE = 1 THEN 'SIPP'  
                                 WHEN ACCOUNTS.ACCOUNTTYPE = 2 THEN 'IFISA'  
                                 WHEN ACCOUNTS.ACCOUNTTYPE = 3 THEN 'SSAS'  
                                 WHEN ACCOUNTS.ACCOUNTTYPE >= 4 THEN 'BadValue'  
                               END AS ACCOUNTTYPEDESCRIPTION,  
                               CASE WHEN ACCOUNTS.ACCOUNTID IS NOT NULL THEN CONVERT(VARCHAR(100),ACCOUNTS.ACCOUNTID)  
                               ELSE 'N/A'  
                               END AS ACCOUNTS_ACCOUNTID,  
                               CASE WHEN USERS.individorg = 1 THEN 'INDIVIDUAL'  
                                 WHEN USERS.individorg = 0 THEN  
                                   CASE WHEN Users.orgtype = 0 THEN 'LIMITED COMPANY'  
                                     WHEN Users.orgtype = 1 THEN 'LLC'  
                                     WHEN Users.orgtype = 2 THEN 'SOLE TRADER'  
                                     WHEN Users.orgtype = 3 THEN 'SASS'  
                                     WHEN Users.orgtype IS NULL THEN 'N/A'  
                                     WHEN (USERS.orgtype) = '' THEN 'N/A'  
                                   END  
                               END AS LENDERTYPE  {Environment.NewLine}")
                strIndividualLenderProfileData.Append($"FROM   accounts {Environment.NewLine}")
                strIndividualLenderProfileData.Append($"       RIGHT OUTER JOIN users {Environment.NewLine}")
                strIndividualLenderProfileData.Append($"                     ON ( accounts.userid = users.userid ) {Environment.NewLine}")
                strIndividualLenderProfileData.Append($"WHERE  ( {searchBySqlField} = {searchValueFromSelectedByUser} ) {Environment.NewLine}")
                'strIndividualLenderProfileData.Append($"         AND ( users.activated = 5 ) {Environment.NewLine}")
                'strIndividualLenderProfileData.Append($"         AND ( users.activated_bank = 5 ) {Environment.NewLine}")
                'strIndividualLenderProfileData.Append($"         AND ( users.isactive = 0 ) {Environment.NewLine}")
                'strIndividualLenderProfileData.Append($"         AND {searchBySqlField} = {searchValueFromSelectedByUser} ) {Environment.NewLine}")
                strIndividualLenderProfileData.Append($"ORDER  BY users.userid")




                Dim sSQL As String
                Dim sErrorStr As String = ""

                Dim ds = New DataSet
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        sSQL = strIndividualLenderProfileData.ToString()


                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using


                'Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                '    Try
                '        Using cmd As New SqlCommand("GetIndividualLenderProfileData")
                '            With cmd.Parameters
                '                .Add(New SqlParameter("@searchBySqlField", searchBySqlField))
                '                .Add(New SqlParameter("@searchValueFromSelectedByUser", searchValueFromSelectedByUser))

                '            End With
                '            cmd.Connection = con
                '            cmd.CommandType = CommandType.StoredProcedure
                '            Using sda As New SqlDataAdapter(cmd)

                '                sda.Fill(result)


                '            End Using
                '        End Using
                '    Catch ex As Exception
                '    Finally

                '    End Try
                'End Using


                _LenderType = result.Rows(0)("LENDERTYPE").ToString().Trim()
                    _LenderCategory = result.Rows(0)("CLIENT_CATEGORISATION").ToString().Trim()
                    _LenderActivationType = result.Rows(0)("USERS_ACTIVATED").ToString().Trim()
                    _LenderActivatedCert = result.Rows(0)("USERS_ACTIVATED_CERT").ToString().Trim()
                    _LenderBankActivationType = result.Rows(0)("USERS_ACTIVATED_BANK").ToString().Trim()
                    _LenderCreatedDate = result.Rows(0)("USERS_DATECREATED").ToString().Trim()
                    _LenderFirstEmailSentOn = result.Rows(0)("LENDERTYPE").ToString().Trim()
                    _LenderEmailActivation = result.Rows(0)("USERS_ACTIVATIONDATE").ToString().Trim()
                    _LenderForeName = result.Rows(0)("USERS_FIRSTNAME").ToString().Trim()
                    _LenderSurname = result.Rows(0)("USERS_LASTNAME").ToString().Trim()
                    _LenderDOB = result.Rows(0)("USERS_DATEOFBIRTH").ToString().Trim()
                    _LenderUserId = Convert.ToInt16(result.Rows(0)("USERS_USERID"))
                    _LenderAccountId = result.Rows(0)("ACCOUNTS_ACCOUNTID").ToString().Trim()
                    _LenderAccountType = result.Rows(0)("ACCOUNTTYPEDESCRIPTION").ToString().Trim()
                    _LenderName = $"{IndividualLenderProfileDetails._LenderForeName} {_LenderSurname}"
                    _LenderSeqAlpha = result.Rows(0)("USERS_SEQALPHA").ToString().Trim()
                    _LenderAccountReference = GenDB.GetInvestorRef(_LenderSeqAlpha, Convert.ToInt32(_LenderAccountId))
                    _LenderEmailAddress = result.Rows(0)("USERS_EMAIL").ToString()
                    _LenderPhoneNumber = result.Rows(0)("USERS_TELEPHONE").ToString()
                    ' GetFirstEmailSent(_LenderEmailAddress)   '**************** Check Later
                    _LenderBankAccountName = result.Rows(0)("ACCOUNTS_BANKACCNAME").ToString()
                    _LenderBankAccountNumber = result.Rows(0)("ACCOUNTS_BANKACCNUMBER").ToString()
                    _LenderBankSortCode = result.Rows(0)("ACCOUNTS_BANKACCSORTCODE").ToString()
                    _LenderBuildingSocietyRollNum = result.Rows(0)("ACCOUNTS_BUILDSOCROLLNUMBER").ToString()
                    _LenderHouseNameNum = result.Rows(0)("USERS_HOUSENAME").ToString()
                    _LenderAddress1 = result.Rows(0)("USERS_ADDRESS1").ToString()
                    _LenderAddress2 = result.Rows(0)("USERS_ADDRESS2").ToString()
                    _LenderAddress3 = result.Rows(0)("USERS_ADDRESS3").ToString()
                    _LenderAddress4 = result.Rows(0)("USERS_ADDRESS4").ToString()
                    _LenderTownCity = result.Rows(0)("USERS_TOWN").ToString()
                    _LenderCounty = result.Rows(0)("USERS_COUNTY").ToString()
                    _LenderCountry = result.Rows(0)("USERS_COUNTRY").ToString()
                    _LenderPostCode = result.Rows(0)("USERS_POSTCODE").ToString()
                Else
                    result = Nothing
                End If



        Catch ex As Exception

            Throw
        Finally


            strIndividualLenderProfileData.Length = 0
        End Try

        Return result

    End Function

    'Public Sub GetFirstEmailSent(ByVal emailAddress As String)

    '    Dim cnn As IDbConnection = Nothing
    '    Dim cmd As IDbCommand = Nothing
    '    Dim param As IDataParameter


    '    Try
    '        Dim strIndividualLenderProfileDataFirstEmailSentOn = String.Empty
    '        strIndividualLenderProfileDataFirstEmailSentOn = $"SELECT   First 1  CASE WHEN  datetimecreated IS NOT NULL THEN  Trim(datetimecreated) ELSE 'N/A'  END AS FirstEmailSentOn from email_sent  where emailto = '{IndividualLenderProfileDetails._LenderEmailAddress}'  order by DATETIMECREATED"
    '        cnn = IAFDataManager.Provider.CreateConnection(IAFDataManager.Provider.ConnectString, True)

    '        If cnn.State.Equals(ConnectionState.Closed) Then
    '            cnn.Open()
    '        End If

    '        cmd = idp.CreateCommand(IAFDataManager.Provider.ConnectString)
    '        param = IAFDataManager.Provider.CreateParameter("@I_EMAIL_ID", DbType.Int32)
    '        param.Value = emailAddress
    '        cmd.Parameters.Add(param)
    '        cmd.CommandType = CommandType.Text
    '        cmd.CommandText = strIndividualLenderProfileDataFirstEmailSentOn
    '        cmd.Connection = cnn
    '        _LenderFirstEmailSentOn = Convert.ToDateTime(IAFDataManager.Provider.ExecuteScalar(strIndividualLenderProfileDataFirstEmailSentOn, cnn)).ToString("dd/MM/yyyy")
    '        _LenderFirstEmailSentOn = If((_LenderFirstEmailSentOn = "01/01/0001"), "N/A", _LenderFirstEmailSentOn)


    '    Catch ex As Exception

    '        Throw
    '    Finally

    '        If (cnn IsNot Nothing) AndAlso cnn.State.Equals(ConnectionState.Open) Then
    '            cnn.Close()
    '        End If

    '        If cmd IsNot Nothing Then
    '            cmd.Dispose()
    '        End If
    '    End Try

    'End Sub

    'Sub AddParametersForProfileEdit(ByVal cmd1 As IDbCommand, ByVal idp As IAFDataProvider, ByVal param As IDataParameter)
    '    param = idp.CreateParameter("@S_FIRSTNAME", DbType.String)
    '    param.Value = _LenderForeNameAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_LASTNAME", DbType.String)
    '    param.Value = _LenderSurnameAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_DOB", DbType.String)
    '    param.Value = _LenderDOBAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_HOUSENAME", DbType.String)
    '    param.Value = _LenderHouseNameNumAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_ADDRESS1", DbType.String)
    '    param.Value = _LenderAddress1AmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_ADDRESS2", DbType.String)
    '    param.Value = _LenderAddress2AmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_ADDRESS3", DbType.String)
    '    param.Value = _LenderAddress3AmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_ADDRESS4", DbType.String)
    '    param.Value = _LenderAddress4AmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_TOWN", DbType.String)
    '    param.Value = _LenderTownCityAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_COUNTY", DbType.String)
    '    param.Value = _LenderCountyAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_COUNTRY", DbType.String)
    '    param.Value = _LenderCountryAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_POSTCODE", DbType.String)
    '    param.Value = _LenderPostCodeAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_EMAIL", DbType.String)
    '    param.Value = _LenderEmailAddressAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_TELEPHONE", DbType.String)
    '    param.Value = _LenderPhoneNumberAmendedTo
    '    cmd1.Parameters.Add(param)
    '    Dim dateTimeNow = DateTime.Now.ToString().Replace("/", ".")
    '    param = idp.CreateParameter("@S_ClientCategorisation", DbType.Int16)
    '    param.Value = _LenderCategoryAmendedTo
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@S_ClientCategorisationDate", DbType.String)
    '    param.Value = dateTimeNow
    '    cmd1.Parameters.Add(param)
    '    param = idp.CreateParameter("@DT_UPDATED", DbType.String)
    '    param.Value = dateTimeNow
    '    cmd1.Parameters.Add(param)
    'End Sub

    'Sub SaveIndividualLenderProfileDetailAmendedByUser()

    '    Dim cnn As IDbConnection = Nothing
    '        Dim trans As IDbTransaction = Nothing
    '        Dim cmd1 As IDbCommand = Nothing
    '        Dim param As IDataParameter
    '        Dim sbSqlUpdateIndividualLenderProfileDetail = New StringBuilder(512)
    '        Dim idp As IAFDataProvider = IAFDataManager.Provider
    '        cnn = idp.CreateConnection(IAFDataManager.Provider.ConnectString, True)

    '        Try
    '            trans = cnn.BeginTransaction()
    '            cmd1 = idp.CreateCommand(IAFDataManager.Provider.ConnectString)
    '            param = idp.CreateParameter()
    '            AddParametersForProfileEdit(cmd1, IAFDataManager.Provider, param)
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("UPDATE Users {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("SET    FIRSTNAME = @S_FIRSTNAME, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       LASTNAME = @S_LASTNAME, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       DateOfBirth = @S_DOB, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       HOUSENAME = @S_HOUSENAME, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       ADDRESS1 = @S_ADDRESS1, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       ADDRESS2 = @S_ADDRESS2, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       ADDRESS3 = @S_ADDRESS3, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       ADDRESS4 = @S_ADDRESS4, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       TOWN = @S_TOWN, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       COUNTY = @S_COUNTY, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       COUNTRY = @S_COUNTRY, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       POSTCODE = @S_POSTCODE, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       EMAIL = @S_EMAIL, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       TELEPHONE = @S_TELEPHONE, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       CLIENT_CATEGORISATION = @S_ClientCategorisation, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       CLIENT_CATEGORISATION_DATE = @S_ClientCategorisationDate, {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("       UPDATED = @DT_UPDATED {0}", Environment.NewLine))
    '            sbSqlUpdateIndividualLenderProfileDetail.Append(String.Format("WHERE  userid = '{0}'", IndividualLenderProfileDetails._LenderUserId))
    '            cmd1.CommandType = CommandType.Text
    '            cmd1.CommandText = sbSqlUpdateIndividualLenderProfileDetail.ToString()
    '            cmd1.Transaction = trans
    '            cmd1.Connection = cnn
    '            IAFDataManager.Provider.ExecuteSQL(cmd1, False)
    '            trans.Commit()

    '    Catch ex As Exception

    '        Throw
    '        Finally

    '            If cnn IsNot Nothing Then
    '                cnn.Close()
    '            End If

    '            If Not cmd1.Equals(Nothing) Then
    '                cmd1 = Nothing
    '            End If
    '        End Try

    'End Sub

    Public Shared Function GetLenderCategorisation(ByVal categorisationText As String) As Integer
        Dim result As Integer = 0

        Select Case categorisationText
            Case "Elective Professional"
                result = 1
            Case "Self Certified Sophisticated"
                result = 2
            Case "High Networth Individual"
                result = 3
            Case "Per Se Professional"
                result = 4
        End Select

        Return result
    End Function

    Public Shared Function ExportToCsv() As StringBuilder
        If Not IndividualLenderProfileDetails._IndividualLenderProfileDataTable.Equals(Nothing) Then
            ' Return IAFDataManager.Provider.ExportDataTableToCSV(IndividualLenderProfileDetails._IndividualLenderProfileDataTable, "USERS_", "LENDER_", True)
        End If

        Return Nothing
    End Function

    Public Shared Function GetIndividualLenderCategorisationHistory(ByVal accountIdSelectedByUser As String) As DataTable

        Dim result As DataTable = New DataTable()


        Dim strSqlIndividualLenderCategorisationHistory As StringBuilder = New StringBuilder()

        Try

            If Not String.IsNullOrWhiteSpace(accountIdSelectedByUser) Then
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("SELECT * From (SELECT  {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("CASE WHEN USERS.CLIENT_CATEGORISATION IS NOT NULL THEN {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("           CASE USERS.CLIENT_CATEGORISATION WHEN 0 THEN 'Uninitialized' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 1 THEN 'Elective Professional' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 2 THEN 'Self Certified Sophisticated' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 3 THEN 'High Networth Individual' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 4 THEN 'Per Se Professional' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 5 THEN 'Eligible counterparty' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 9 THEN 'Awaiting Signoff' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("           END {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("       ELSE 'Uninitialized' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("       END AS CLIENT_CATEGORISATION, {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("        CASE WHEN CLIENT_CATEGORISATION_DATE IS NOT NULL THEN CLIENT_CATEGORISATION_DATE ELSE 'N/A' END as CLIENT_CATEGORISATION_DATE {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("FROM   USERS {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("WHERE  ( USERS.USERID = @I_ACCOUNT_ID ) {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format(" {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("UNION ALL {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format(" {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("SELECT  {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("CASE WHEN NEW_USERS_HISTORY.CLIENT_CATEGORISATION IS NOT NULL THEN {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("           CASE NEW_USERS_HISTORY.CLIENT_CATEGORISATION WHEN 0 THEN 'Uninitialized' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 1 THEN 'Elective Professional' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 2 THEN 'Self Certified Sophisticated' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 3 THEN 'High Networth Individual' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 4 THEN 'Per Se Professional' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 5 THEN 'Eligible counterparty' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("             WHEN 9 THEN 'Awaiting Signoff' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("           END {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("       ELSE 'Uninitialized' {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("       END AS CLIENT_CATEGORISATION, {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("        CASE WHEN CLIENT_CATEGORISATION_DATE IS NOT NULL THEN CLIENT_CATEGORISATION_DATE ELSE 'N/A' END as CLIENT_CATEGORISATION_DATE {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append(String.Format("FROM   NEW_USERS_HISTORY {0}", Environment.NewLine))
                strSqlIndividualLenderCategorisationHistory.Append("WHERE  ( NEW_USERS_HISTORY.USERID =@I_ACCOUNT_ID )) ORDER BY  CLIENT_CATEGORISATION_DATE DESC ")

                Dim sSQL As String
                Dim sErrorStr As String = ""

                Dim ds = New DataSet
                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()
                        sSQL = strSqlIndividualLenderCategorisationHistory.ToString().ToString()


                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        With cmd.Parameters
                            .Add(New SqlParameter("@I_ACCOUNT_ID", DbType.Int32))

                        End With
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using

                result.DefaultView.Sort = "CLIENT_CATEGORISATION_DATE"
                _IndividualLenderCategorisationHistoryHasRows = False

                If result.Rows.Count > 0 Then
                    _IndividualLenderCategorisationHistoryHasRows = True
                End If
            Else
                result = Nothing
            End If

        Catch ex As Exception

            Throw
        Finally



            strSqlIndividualLenderCategorisationHistory.Length = 0
        End Try


    End Function

End Class

