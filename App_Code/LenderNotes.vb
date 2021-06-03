Imports System
Imports System.Data
Imports System.Diagnostics
Imports System.Text
Imports System.Data.SqlClient

Public Class LenderNotes
    Public Shared _LenderNoteSystemUserId As String
    Public Shared _LenderNotIsActive As String
    Public Shared _LenderNoteIsPinned As String
    Public Shared _LenderNoteId As String
    Public Shared _LenderNotesDataTable As DataTable
    Public Shared _LenderNoteWorkingOn As String
    Public Const _LenderNotesMaxLength As Integer = 8192
    Public Shared _LenderNotesAsClipboard As String
    Public Shared _LenderNotesNeedsClipboard As Boolean
    Public Shared _LenderNotesTaken As String
    Public Shared _LenderAccountIdForNotes As String

    Public Shared Function GetLenderNotes(ByVal accountIdSelectedByUser As String) As DataTable
        Dim sSQL As String
        Dim result As DataTable = New DataTable()
        Dim strSqlLenderNotesForSelectedLender As StringBuilder = New StringBuilder()
        Dim cnn As IDbConnection = Nothing

        Try

            If Not String.IsNullOrWhiteSpace(accountIdSelectedByUser) Then

                sSQL = "Select 
                      ln.NOTE_ID,  
                      ln.DATECREATED,  
                      ln.LAST_MODIFIED,  
                            concat(Trim(su.firstname) 
                            , ' '  
                            , Trim(su.lastname)) AS IAF_USER,  
                      ln.NOTE_TEXT,  
                      ln.SOURCE_ID,  
                      ln.IS_PINNED 
                      from Lender_Notes ln , System_Users su 
                      Where ln.ISACTIVE = 0  
                            AND ln.SYSTEM_USER_ID = su.SYSTEM_USER_ID  
                    AND (ln.UserId In(@accountIdSelectedByUser))  
                     ORDER BY ln.IS_PINNED Desc,ln.last_modified DESC,  
                              ln.datecreated DESC"

                Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                    Try
                        Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                        Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                        With cmd.Parameters
                            .Add(New SqlParameter("@accountIdSelectedByUser", IndividualLenderProfileDetails._LenderAccountId))

                        End With
                        con.Open()

                        adapter.SelectCommand = cmd

                        adapter.Fill(result)


                    Catch ex As Exception
                    Finally

                    End Try
                End Using


            Else
                result = Nothing
            End If

        Catch ex As Exception

            Throw
        Finally


        End Try

        Return result

    End Function

    Public Shared Function GetLenderNotesSearchValue(ByVal strSearchValue As String, ByVal strLenderFilterForSelectedLender As String) As String
        Dim userId = String.Empty

        If Not String.IsNullOrWhiteSpace(strSearchValue) AndAlso Not String.IsNullOrWhiteSpace(strLenderFilterForSelectedLender) Then
            Dim searchValues As String() = strSearchValue.Split("-"c)
            userId = String.Format("""Users.Userid,{0}""", searchValues(2))

            Select Case strLenderFilterForSelectedLender
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
                Case "Company Name"
                    userId = searchValues(2)
            End Select
        End If

        Return userId
    End Function

    Sub AddParametersForInsertNote(ByVal cmd1 As IDbCommand, ByVal param As IDataParameter)
        'param = idp.CreateParameter("@I_LenderAccountId", DbType.Int32)
        'param.Value = Convert.ToInt32(_LenderAccountIdForNotes)
        'cmd1.Parameters.Add(param)
        'param = idp.CreateParameter("@I_SystemUserId", DbType.Int32)
        'param.Value = Convert.ToInt32(_LenderNoteSystemUserId)
        'cmd1.Parameters.Add(param)
        'param = idp.CreateParameter("@I_NoteText", DbType.String)
        'param.Value = _LenderNotesTaken
        'cmd1.Parameters.Add(param)
        'param = idp.CreateParameter("@I_IsPinned", DbType.Int16)
        'param.Value = Convert.ToInt16(_LenderNoteIsPinned)
        'cmd1.Parameters.Add(param)
    End Sub

    Sub SaveNoteToDB()

        Dim cnn As IDbConnection = Nothing
            Dim trans As IDbTransaction = Nothing
            Dim cmd1 As IDbCommand = Nothing
            Dim param As IDataParameter
            Dim sbStatus = New StringBuilder()
        Dim sSQL As String
        Dim rowsAffectedOnInsertNote As Integer

        Dim sbSqlInsertNote = New StringBuilder(512)

        Try
            sSQL = "INSERT INTO LENDER_NOTES  
                 USERID, SYSTEM_USER_ID, SOURCE_ID, NOTE_TEXT, IS_PINNED ) Values 
             @I_LenderAccountId, @I_SystemUserId, 15, @I_NoteText, @I_IsPinned)"




            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)
                Try
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter()



                    Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                    With cmd.Parameters
                        .Add(New SqlParameter("@I_LenderAccountId", _LenderAccountIdForNotes))
                        .Add(New SqlParameter("@I_SystemUserId", _LenderNoteSystemUserId))
                        .Add(New SqlParameter("@I_NoteText", _LenderNotesTaken))
                        .Add(New SqlParameter("@I_IsPinned", _LenderNoteIsPinned))
                    End With
                    con.Open()

                    adapter.SelectCommand = cmd

                    cmd.ExecuteNonQuery()


                Catch ex As Exception
                Finally

                End Try
            End Using



        Catch ex As Exception

            Throw
        Finally

            If cnn IsNot Nothing Then
                    cnn.Close()
                End If

                If Not cmd1.Equals(Nothing) Then
                    cmd1 = Nothing
                End If
            End Try

    End Sub
End Class
