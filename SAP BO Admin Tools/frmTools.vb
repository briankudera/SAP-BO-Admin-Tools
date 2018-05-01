Imports CrystalDecisions.Enterprise
Imports System.Data.SqlClient
Imports CrystalDecisions.Enterprise.Desktop

Public Class frmTools

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private boEnterpriseSession As EnterpriseSession
    Private boInfoStore As InfoStore
    Private boSessionMgr As SessionMgr
    Private boAlias As UserAliases
    Private conSQLConn As SqlConnection
    Private cmdSQLCmd As SqlCommand

    Private cmdAction As String = ""
    Private cmdCMSServer As String = ""
    Private cmdCMSUser As String = ""
    Private cmdCMSUserPassword As String = ""
    Private cmdCMSAuthentication As String = ""
    Private cmdTargetServer As String = ""
    Private cmdTargetDB As String = ""
    Private cmdCommandLine As String = ""

    Private Sub Tools_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim arguments As String() = Environment.GetCommandLineArgs()

        If arguments.Count > 1 Then
            ConsoleMain(arguments)
        End If

        If cboCMSServer.Items.Count > 0 Then
            cboCMSServer.SelectedIndex = 0    ' The first item has index 0 '
        End If
        If cboCMSAuthentication.Items.Count > 0 Then
            cboCMSAuthentication.SelectedIndex = 0    ' The first item has index 0 '
        End If

    End Sub

    Private Sub ConsoleMain(ByVal arguments() As String)

        cmdCommandLine = arguments(0)
        cmdAction = arguments(1)
        cmdCMSServer = arguments(2)
        cmdCMSUser = arguments(3)
        cmdCMSUserPassword = arguments(4)
        cmdCMSAuthentication = arguments(5)
        cmdTargetServer = arguments(6)
        cmdTargetDB = arguments(7)

        If cmdAction = "LoadUsersToDB" Then
            GetBOUserList(False, True, cmdTargetDB, cmdTargetServer)
        ElseIf cmdAction = "LoadObjectsToDB" Then
            GetBOObjectList(False, True, cmdTargetDB, cmdTargetServer)
        End If

        Application.Exit()

    End Sub

    Private Sub NewBOSession()

        Dim mgr As New SessionMgr
        Dim strUserName As String = ""
        Dim strCMSName As String = ""
        Dim strPassword As String = ""
        Dim strAuthentication As String = ""

        If cmdCMSUser.Length > 0 Then
            strUserName = cmdCMSUser
            strCMSName = cmdCMSServer
            strPassword = cmdCMSUserPassword
            strAuthentication = cmdCMSAuthentication
        Else
            strUserName = Me.txtCMSUserName.Text.ToString()
            strCMSName = Me.cboCMSServer.SelectedItem.ToString()
            strPassword = Me.txtCMSUserPassword.Text.ToString()
            strAuthentication = Me.cboCMSAuthentication.SelectedItem.ToString()
        End If

        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText("************************************************************************" & ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText("Opening session to SAP BO." & ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.boEnterpriseSession = mgr.Logon(strUserName, strPassword, strCMSName, strAuthentication)
        Me.boInfoStore = New InfoStore(Me.boEnterpriseSession.GetService("InfoStore"))

    End Sub

    Private Sub LogoffBOSession()

        Me.boEnterpriseSession.Logoff()
        Me.boEnterpriseSession.Dispose()
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText("Closing session to SAP BO." & ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText("************************************************************************" & ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))

    End Sub


    Private Function GetIdForUser(ByVal strUserName As String) As Integer
        Dim strUserId As String
        Dim intUserId As Integer

        Dim strQuery As String = ("select SI_ID from CI_SYSTEMOBJECTS WHERE SI_KIND = 'User' and SI_NAME = '" & strUserName & "'")
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)
        If (objects.ResultCount = 1) Then
            strUserId = objects.Item(1).Properties.Item("SI_ID").Value.ToString
        Else
            strUserId = "0"
        End If
        If Not Integer.TryParse(strUserId, intUserId) Then
            intUserId = -1
        End If
        Return intUserId

    End Function

    Private Sub btnRepalceOwnerOnAllObjects_Click(sender As Object, e As EventArgs) Handles btnRepalceOwnerOnAllObjects.Click

        Me.NewBOSession()

        Dim intOwnerIdOld As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Text)
        Dim intOwnerIdNew As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Text)
        Dim strQuery As String = ("select top 1000000 * from ci_infoobjects where (si_ownerid = " & intOwnerIdOld.ToString & " or si_author = '" & Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Text & "') and si_kind not in ('FavoritesFolder','PersonalCategory','Inbox')")
        Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If (infoObjects.Count > 0) Then
            Dim enumerator As IEnumerator
            Try
                enumerator = infoObjects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim objCurrentObject As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim strObjectName As String = objCurrentObject.Properties.Item("SI_NAME").Value.ToString
                    Dim strObjectKind As String = objCurrentObject.Properties.Item("SI_KIND").Value.ToString
                    Dim strObjectId As String = objCurrentObject.Properties.Item("SI_ID").Value.ToString
                    Dim blnIsInstance As Boolean = objCurrentObject.Instance

                    'If instance, also set the submitterid field, if the field exists
                    If blnIsInstance Then

                        'This property may not exist
                        Dim strSubmitterId As String = ""
                        Try
                            strSubmitterId = objCurrentObject.SchedulingInfo.Properties.Item("SI_SUBMITTERID").Value.ToString()
                        Catch ex As Exception
                            Exit Try
                        End Try

                        If strSubmitterId <> "" Then
                            objCurrentObject.SchedulingInfo.Properties.Item("SI_SUBMITTERID").Value = intOwnerIdNew
                        End If

                    End If

                    objCurrentObject.Properties.Item("SI_OWNERID").Value = intOwnerIdNew
                    objCurrentObject.Properties.Item("SI_AUTHOR").Value = Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Text

                    Dim txtOutput As String() = New String() {"Object updated: (", strObjectName, ") ", strObjectName, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))

                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Me.boInfoStore.Commit(infoObjects)
        End If

        Me.LogoffBOSession()

    End Sub


    Private Sub btnReplaceOwnerOnSingeDoc_Click(sender As Object, e As EventArgs) Handles btnReplaceOwnerOnSingeDoc.Click

        Dim intDocId As Integer
        Me.NewBOSession()

        If Not Integer.TryParse(Me.txtReplaceOwnerOnSingleDocDocId.Text, intDocId) Then
            intDocId = -1
        End If
        If (intDocId > 0) Then
            Dim intOwnerIdNew As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Text)
            If (intOwnerIdNew > 0) Then

                Dim strQuery As String = ("select top 100000 * from ci_infoobjects where si_id = " & intDocId.ToString)
                Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)
                If (infoObjects.Count > 0) Then
                    Dim enumerator As IEnumerator
                    Try
                        enumerator = infoObjects.GetEnumerator
                        Do While enumerator.MoveNext
                            Dim objCurrentObject As InfoObject = DirectCast(enumerator.Current, InfoObject)
                            Dim strUserId As String = objCurrentObject.Properties.Item("SI_NAME").Value.ToString
                            Dim strObjectName As String = objCurrentObject.Properties.Item("SI_NAME").Value.ToString
                            Dim strObjectKind As String = objCurrentObject.Properties.Item("SI_KIND").Value.ToString
                            Dim blnIsInstance As Boolean = objCurrentObject.Instance
                            Dim strObjectId As String = objCurrentObject.Properties.Item("SI_ID").Value.ToString

                            'If instance, also set the submitterid field, if the field exists
                            If blnIsInstance Then

                                'This property may not exist
                                Dim strSubmitterId As String = ""
                                Try
                                    strSubmitterId = objCurrentObject.SchedulingInfo.Properties.Item("SI_SUBMITTERID").Value.ToString()
                                Catch ex As Exception
                                    Exit Try
                                End Try

                                If strSubmitterId <> "" Then
                                    objCurrentObject.SchedulingInfo.Properties.Item("SI_SUBMITTERID").Value = intOwnerIdNew
                                End If
                            End If

                            objCurrentObject.Properties.Item("SI_OWNERID").Value = intOwnerIdNew

                            Dim txtOutput As String() = New String() {"Object updated: (", strObjectKind, ") ", strObjectName, ChrW(13) & ChrW(10)}
                            Me.rtbOutput.AppendText(String.Concat(txtOutput))
                        Loop
                    Finally
                        If TypeOf enumerator Is IDisposable Then
                            TryCast(enumerator, IDisposable).Dispose()
                        End If
                    End Try
                End If
                Me.boInfoStore.Commit(infoObjects)
            Else
                Me.rtbOutput.AppendText("Unable to find ID of new Owner in system" & ChrW(13) & ChrW(10))
            End If
        Else
            Me.rtbOutput.AppendText("Unable to parse object id" & ChrW(13) & ChrW(10))
        End If
        Me.LogoffBOSession()

    End Sub

    Private Sub btnListObjectsByOwner_Click(sender As Object, e As EventArgs) Handles btnListObjectsByOwner.Click

        Dim strUserOwnerName As String = Me.txtListObjectsByOwnerOwnerName.Text.ToString
        Me.rtbOutput.AppendText("Begin User ID lookup" & ChrW(13) & ChrW(10))

        Me.NewBOSession()

        Dim intOwnerId As Integer = Me.GetIdForUser(strUserOwnerName)
        If (intOwnerId > 0) Then
            Me.rtbOutput.AppendText("User ID lookup successful:" & ChrW(13) & ChrW(10))
        Else
            Me.rtbOutput.AppendText("User ID lookup failed:" & ChrW(13) & ChrW(10))
        End If

        Dim strQuery As String = ("select si_id, si_name, si_ownerid, si_kind from ci_infoobjects where si_ownerid = " & intOwnerId.ToString & " and si_kind not in ('FavoritesFolder','PersonalCategory','Inbox')")
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If (objects.Count > 0) Then
            Dim enumerator As IEnumerator
            Try
                enumerator = objects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim strObjectName As String = current.Properties.Item("SI_NAME").Value.ToString
                    Dim strObjectKind As String = current.Properties.Item("SI_KIND").Value.ToString
                    Dim strObjectId As String = current.Properties.Item("SI_ID").Value.ToString
                    Dim textArray1 As String() = New String() {"Object found: (", strObjectKind, ") (Document ID: ", strObjectId, ") ", strObjectName, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(textArray1))
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If

        Me.LogoffBOSession()

    End Sub

    Private Sub btnGetListOfUsers_Click(sender As Object, e As EventArgs) Handles btnGetListOfUsers.Click

        Me.GetBOUserList(True, False)

    End Sub

    Private Sub btnLoadListOfUsersToDB_Click(sender As Object, e As EventArgs) Handles btnLoadListOfUsersToDB.Click

        GetBOUserList(False, True, Me.txtLoadListOfUsersToDBDatabaseName.Text.ToString(), Me.txtLoadListOfUsersToDBSQLServerName.Text.ToString())

    End Sub

    Protected Sub GetBOUserList(blnDisplay As Boolean, blnLoadDatabase As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "")

        Me.NewBOSession()

        Dim strQuery As String = ("Select SI_ID, SI_CUID, SI_NAME, SI_USERFULLNAME, SI_DESCRIPTION, SI_EMAIL_ADDRESS, SI_LASTLOGONTIME FROM CI_SYSTEMOBJECTS Where SI_KIND='User' ORDER BY SI_ID")
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If (objects.Count > 0) Then
            Dim enumerator As IEnumerator
            Try
                enumerator = objects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim strUserId As String = current.Properties.Item("SI_ID").Value.ToString()
                    Dim strCUID As String = current.Properties.Item("SI_CUID").Value.ToString()
                    Dim strName As String = current.Properties.Item("SI_NAME").Value.ToString()
                    Dim strFullName As String = current.Properties.Item("SI_USERFULLNAME").Value.ToString()
                    Dim strDescription As String = current.Properties.Item("SI_DESCRIPTION").Value.ToString()
                    Dim strEmailAddress As String = current.Properties.Item("SI_EMAIL_ADDRESS").Value.ToString()

                    'This property may not exist if the user has never logged on
                    Dim strLastLogonTime As String
                    Try
                        strLastLogonTime = current.Properties.Item("SI_LASTLOGONTIME").Value.ToString()
                    Catch ex As Exception
                        strLastLogonTime = "01/01/1901"
                        Exit Try
                    End Try

                    Dim strDisabled = CheckIfUserIsDisabled(strUserId)

                    If blnDisplay Then
                        Dim arrResults As String() = New String() {"User found: (", strName, ") User ID: ", strUserId, "  Disabled: ", strDisabled, " Last Logon: ", strLastLogonTime, ChrW(13) & ChrW(10)}
                        Me.rtbOutput.AppendText(String.Concat(arrResults))
                    End If

                    If blnLoadDatabase Then
                        Me.SetSQLConnection(strDatabaseName, strSQLServerName)
                        CreateUserTable()
                        LoadUserToDatabase(strUserId, strCUID, strName, strFullName, strDescription, strEmailAddress, strLastLogonTime, strDisabled)
                    End If

                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If

        Me.LogoffBOSession()

    End Sub

    Private Function CheckIfUserIsDisabled(strUserId As String) As String

        Dim intBOAliasCount As Int16 = 0
        Dim intTotalDisabled As Int16 = 0

        Dim strQuery As String = ("Select SI_ALIASES FROM CI_SYSTEMOBJECTS WHERE SI_ID = " + strUserId)
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        'Loop through objects
        If (objects.Count > 0) Then

            Dim strDisabled As String = ""
            Dim strAliasName As String = ""
            Dim strAliasType As String = ""

            Dim enumerator = objects.GetEnumerator
            Try

                Do While enumerator.MoveNext

                    Dim objInfoObject As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim objUser As User = DirectCast(objInfoObject, User)
                    Dim objUserEnumerator = objUser.Aliases().GetEnumerator

                    intBOAliasCount = objUser.Aliases.Count()

                    Try

                        Do While objUserEnumerator.MoveNext

                            Dim boAlias As UserAlias = DirectCast(objUserEnumerator.Current, UserAlias)

                            strDisabled = boAlias.Disabled.ToString()

                            If strDisabled = "True" Then
                                intTotalDisabled = intTotalDisabled + 1
                            End If

                            strAliasName = boAlias.Name()
                            strAliasType = boAlias.Type()

                        Loop

                    Catch ex As Exception
                    Finally
                        If TypeOf objUserEnumerator Is IDisposable Then
                            TryCast(objUserEnumerator, IDisposable).Dispose()
                        End If

                    End Try

                Loop

            Catch ex As Exception
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try

        End If

        If intBOAliasCount = intTotalDisabled Then
            CheckIfUserIsDisabled = "1"
        Else
            CheckIfUserIsDisabled = "0"
        End If

    End Function

    Private Sub CreateUserTable()

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOUser' and xtype='U') " _
                        & "create table DimSAPBOUser (" _
                        & "    SI_ID int not null primary key" _
                        & "   , SI_CUID varchar(64) Not null" _
                        & "   ,SI_NAME varchar(100) not null" _
                        & "   ,SI_USERFULLNAME varchar(255) null" _
                        & "   ,SI_DESCRIPTION varchar(max) null" _
                        & "   ,SI_EMAIL_ADDRESS varchar(255) null" _
                        & "   ,SI_LASTLOGONTIME datetime not null" _
                        & "   ,Disabled bit not null Default 0" _
                        & "   ,Deleted bit not null default 0" _
                        & "   ,RecordInsertTimestamp datetime2(7) not null default sysdatetime()" _
                        & "   ,RecordUpdateTimestamp datetime2(7) not null default sysdatetime()" _
                        & " )")

    End Sub


    Private Sub CreateUserTableForRepo()

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOObject_Stg' and xtype='U') " _
                        & "create table DimSAPBOObject_Stg (" _
                        & "    SI_ID int not null primary key" _
                        & "   ,SI_CUID varchar(64) not null" _
                        & "   ,SI_NAME varchar(255) null" _
                        & "   ,SI_OWNER varchar(255) null" _
                        & "   ,SI_PARENT_FOLDER int null" _
                        & "   ,SI_INSTANCE bit null default 0" _
                        & "   ,SI_SIZE int null" _
                        & "   ,SI_PARENTID int null" _
                        & "   ,SI_UPDATE_TS datetime null" _
                        & "   ,SI_CREATION_TIME datetime null" _
                        & "   ,SI_KIND varchar(255) null" _
                        & "   ,SI_HAS_CHILDREN bit null default 0" _
                        & "   ,SI_RECURRING bit null default 0" _
                        & "   ,RecordInsertTimestamp datetime2(7) not null default sysdatetime()" _
                        & "   ,RecordUpdateTimestamp datetime2(7) not null default sysdatetime()" _
                        & " )")

        ExecuteQuery("truncate table DimSAPBOObject_Stg")

    End Sub

    Private Sub LoadUserToDatabase(strUserID As String, strCUID As String, strName As String, strFullName As String, strDescription As String, strEmailAddress As String, strLastLogonTime As String, strDisabled As String)

        Dim strQuery As String

        strQuery = "merge " _
                   & "      DimSAPBOUser As tgt" _
                   & " Using " _
                   & "     (" _
                   & "         Select " _
                   & "             " + strUserID + " As SI_ID" _
                   & "             ,'" + strCUID + "' As SI_CUID " _
                   & "             , '" + strName + "' As SI_NAME " _
                   & "             , '" + strFullName + "' As SI_USERFULLNAME" _
                   & "             , '" + strDescription + "' As SI_DESCRIPTION" _
                   & "             , '" + strEmailAddress + "' As SI_EMAIL_ADDRESS" _
                   & "             , '" + strLastLogonTime + "' As SI_LASTLOGONTIME" _
                   & "             , " + strDisabled + " As Disabled" _
                   & "      ) As src" _
                   & " On tgt.SI_ID = src.SI_ID" _
                   & " When matched And" _
                   & " ( " _
                   & "     IsNull(tgt.SI_NAME, '') <> IsNull(src.SI_NAME, '')" _
                   & "  Or IsNull(tgt.SI_USERFULLNAME, '') <> IsNull(src.SI_USERFULLNAME, '')" _
                   & "  Or IsNull(tgt.SI_DESCRIPTION, '') <> IsNull(src.SI_DESCRIPTION, '')" _
                   & "  Or IsNull(tgt.SI_EMAIL_ADDRESS, '') <> IsNull(src.SI_EMAIL_ADDRESS, '')" _
                   & "  Or IsNull(tgt.SI_LASTLOGONTIME, '') <> IsNull(src.SI_LASTLOGONTIME, '')" _
                   & "  Or tgt.Disabled <> src.Disabled" _
                   & " )" _
                   & " then update " _
                   & " set " _
                   & "      tgt.SI_NAME               = src.SI_NAME " _
                   & "     ,tgt.SI_USERFULLNAME       = src.SI_USERFULLNAME " _
                   & "     ,tgt.SI_DESCRIPTION        = src.SI_DESCRIPTION " _
                   & "     ,tgt.SI_EMAIL_ADDRESS      = src.SI_EMAIL_ADDRESS " _
                   & "     ,tgt.SI_LASTLOGONTIME      = src.SI_LASTLOGONTIME " _
                   & "     ,tgt.Disabled              = src.Disabled " _
                   & "     ,tgt.RecordUpdateTimestamp = sysdatetime() " _
                   & " when not matched by target then" _
                   & " insert (" _
                   & "             SI_ID" _
                   & "            ,SI_CUID" _
                   & "            ,SI_NAME" _
                   & "            ,SI_USERFULLNAME" _
                   & "            ,SI_DESCRIPTION" _
                   & "            ,SI_EMAIL_ADDRESS" _
                   & "            ,SI_LASTLOGONTIME" _
                   & "            ,Disabled" _
                   & "        )" _
                   & " values" _
                   & "        (" _
                   & "             " + strUserID + "" _
                   & "             , '" + strCUID + "'" _
                   & "             , '" + strName + "'" _
                   & "             , '" + strFullName + "'" _
                   & "             , '" + strDescription + "'" _
                   & "             , '" + strEmailAddress + "'" _
                   & "             , '" + strLastLogonTime + "'" _
                   & "             , " + strDisabled + "" _
                   & "        )" _
                   & " when not matched by source " _
                   & "      and tgt.SI_ID = " + strUserID + " then " _
                   & "        update set " _
                   & "             tgt.Deleted = 1" _
                   & "            ,tgt.Disabled = 1" _
                   & " ;"

        ExecuteQuery(strQuery)

    End Sub

    Private Sub LoadObjectToDatabase(strId As String, strCUID As String, strName As String, strOwner As String, intParentFolder As String, blnInstance As String, intSize As String, intParentId As String, dteUpdateTimestamp As String, dteCreationTimestamp As String, strKind As String, blnHasChildren As String, blnRecurring As String)

        Dim strQuery As String

        strQuery = "INSERT INTO dbo.DimSAPBOObject_Stg" _
                 & "     ( SI_ID" _
                 & "      ,SI_CUID" _
                 & "      ,SI_NAME" _
                 & "      ,SI_OWNER" _
                 & "      ,SI_PARENT_FOLDER" _
                 & "      ,SI_INSTANCE" _
                 & "      ,SI_SIZE" _
                 & "      ,SI_PARENTID" _
                 & "      ,SI_UPDATE_TS" _
                 & "      ,SI_CREATION_TIME" _
                 & "      ,SI_KIND" _
                 & "      ,SI_HAS_CHILDREN" _
                 & "      ,SI_RECURRING" _
                 & "     )" _
                 & "VALUES" _
                 & "     (" _
                 & "          " + strId + "                  " _
                 & "         ,'" + strCUID + "'              " _
                 & "         ,'" + strName + "'              " _
                 & "         ,'" + strOwner + "'             " _
                 & "         ," + intParentFolder + "        " _
                 & "         ," + blnInstance + "            " _
                 & "         ," + intSize + "                " _
                 & "         ," + intParentId + "            " _
                 & "         ,'" + dteUpdateTimestamp + "'   " _
                 & "         ,'" + dteCreationTimestamp + "' " _
                 & "         ,'" + strKind + "'              " _
                 & "         ," + blnHasChildren + "         " _
                 & "         ," + blnRecurring + "           " _
                 & "     )"

        ExecuteQuery(strQuery)

    End Sub


    Private Sub SetSQLConnection(strDB As String, strServer As String)

        'Create a Connection object.
        conSQLConn = New SqlConnection("Initial Catalog=" + strDB + ";Data Source=" + strServer + ";Integrated Security=SSPI;")

    End Sub

    Private Sub ExecuteQuery(strQuery As String)

        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        conSQLConn.Open()

        cmdSQLCmd.ExecuteNonQuery()

        conSQLConn.Close()

    End Sub

    Private Sub btnExtractRepo_Click(sender As Object, e As EventArgs) Handles btnExtractRepo.Click

        GetBOObjectList(False, True, txtExtractRepoDatabaseName.Text.ToString(), txtExtractRepoSQLServer.Text.ToString())

    End Sub

    Protected Sub GetBOObjectList(blnDisplay As Boolean, blnLoadDatabase As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "")

        Me.NewBOSession()

        Dim strQuery As String = ("Select TOP 1000000 SI_ID, SI_CUID, SI_NAME, SI_OWNER, SI_PARENT_FOLDER, SI_INSTANCE, SI_SIZE, SI_PARENTID, SI_UPDATE_TS, SI_CREATION_TIME, SI_KIND, SI_HAS_CHILDREN, SI_RECURRING FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder')")
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If (objects.Count > 0) Then

            SetSQLConnection(strDatabaseName, strSQLServerName)

            'Create stage table
            CreateUserTableForRepo()

            Dim enumerator As IEnumerator
            Try
                enumerator = objects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim strId As String = current.Properties.Item("SI_ID").Value.ToString()
                    Dim strCUID As String = current.Properties.Item("SI_CUID").Value.ToString()
                    Dim strName As String = Replace(current.Properties.Item("SI_NAME").Value.ToString(), "'", "''")
                    Dim strOwner As String = current.Properties.Item("SI_OWNER").Value.ToString()
                    Dim intParentFolder As String = current.Properties.Item("SI_PARENT_FOLDER").Value.ToString()
                    Dim blnInstance As String = If(current.Properties.Item("SI_INSTANCE").Value.ToString() = "False", "0", "1")
                    Dim intParentId As String = current.Properties.Item("SI_PARENTID").Value.ToString()
                    Dim dteUpdateTimestamp As String = current.Properties.Item("SI_UPDATE_TS").Value.ToString()
                    Dim dteCreationTimestamp As String = current.Properties.Item("SI_CREATION_TIME").Value.ToString()
                    Dim strKind As String = current.Properties.Item("SI_KIND").Value.ToString()
                    Dim blnHasChildren As String = If(current.Properties.Item("SI_HAS_CHILDREN").Value.ToString() = "False", "0", "1")

                    'This property may not exist
                    Dim intSize As String
                    Try
                        intSize = current.Properties.Item("SI_SIZE").Value.ToString()
                    Catch ex As Exception
                        intSize = 0
                        Exit Try
                    End Try

                    'This property may not exist
                    Dim blnRecurring As String
                    Try
                        blnRecurring = If(current.Properties.Item("SI_RECURRING").Value.ToString() = "False", "0", "1")
                    Catch ex As Exception
                        blnRecurring = 0
                        Exit Try
                    End Try

                    LoadObjectToDatabase(strId, strCUID, strName, strOwner, intParentFolder, blnInstance, intSize, intParentId, dteUpdateTimestamp, dteCreationTimestamp, strKind, blnHasChildren, blnRecurring)

                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If

        Me.LogoffBOSession()

    End Sub

End Class
