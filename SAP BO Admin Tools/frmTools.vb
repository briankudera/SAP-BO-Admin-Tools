Imports CrystalDecisions.Enterprise
Imports System.Data.SqlClient
Imports CrystalDecisions.Enterprise.Desktop
Imports NLog
Imports NLog.Targets
Imports NLog.Config


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
    Private cmdDeltaProcessing As Boolean = 0

    Dim myBOProperties As DataTable

    Dim logger As Logger = LogManager.GetLogger("Example")

    Private Sub Tools_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ConfigureLogger()

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

    Private Sub ConfigureLogger()

        Dim config = New LoggingConfiguration()

        Dim consoleTarget = New ColoredConsoleTarget("target1") With {
            .Layout = "${date:format=HH\:mm\:ss} ${level} ${message} ${exception}"
        }

        config.AddTarget(consoleTarget)

        Dim fileTarget = New FileTarget("target2") With {
            .FileName = "${basedir}/SAPBOAdminToolsLog.txt",
            .Layout = "${longdate} ${level} ${message}  ${exception}"
        }
        config.AddTarget(fileTarget)

        config.AddRuleForOneLevel(LogLevel.[Debug], fileTarget)
        config.AddRuleForAllLevels(consoleTarget)

        LogManager.Configuration = config

        'Example logger calls
        logger.Trace("trace log message")
        logger.Debug("debug log message")
        logger.Info("info log message")
        logger.Warn("warn log message")
        logger.[Error]("error log message")
        logger.Fatal("fatal log message")

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
        cmdDeltaProcessing = arguments(8)

        If cmdAction = "LoadUsersToDB" Then
            GetBOUserList(False, cmdTargetDB, cmdTargetServer)
            GetBOUserList(False, cmdTargetDB, cmdTargetServer)
        ElseIf cmdAction = "LoadObjectsToDB" Then
            GetBOObjectList(False, cmdTargetDB, cmdTargetServer)
        ElseIf cmdAction = "LoadObjectPropertiesToDB" Then
            GetBOObjectProperties(False, cmdTargetDB, cmdTargetServer, cmdDeltaProcessing)
        End If

        Application.Exit()

    End Sub

    Private Sub NewBOSession()

        logger.Debug("Begin open session to BI Platform")

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

        Try
            Me.boEnterpriseSession = mgr.Logon(strUserName, strPassword, strCMSName, strAuthentication)
        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error while establishing a BO Enterprise Session")
            Application.Exit()
        End Try

        Try
            Me.boInfoStore = New InfoStore(Me.boEnterpriseSession.GetService("InfoStore"))
        Catch ex As Exception

            logger.[Error](ex, "ow noos! Error while establishing a new InfoStore")
            Application.Exit()
        End Try

    End Sub

    Private Sub LogoffBOSession()

        logger.Debug("Begin close session to BI Platform")

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
        Dim strQuery As String = ("select top 1000000 * from ci_infoobjects where (si_ownerid = " & intOwnerIdOld.ToString & " or si_author = '" & Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Text & "' or si_under_favorites = " & intOwnerIdOld.ToString & ") and si_kind not in ('FavoritesFolder','PersonalCategory','Inbox') and si_name != '~WebIntelligence'")
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

                    'See if this property exists
                    Dim strUnderFavorites As String = ""
                    Try
                        strUnderFavorites = objCurrentObject.Properties.Item("SI_UNDER_FAVORITES").Value.ToString()
                    Catch ex As Exception
                        Exit Try
                    End Try

                    If strUnderFavorites <> "" Then
                        objCurrentObject.Properties.Item("SI_UNDER_FAVORITES").Value = intOwnerIdNew
                    End If

                    'See if this property exists
                    Dim strAuthorSchedule As String = ""
                    Try
                        strAuthorSchedule = objCurrentObject.SchedulingInfo.Properties.Item("SI_AUTHOR").Value.ToString()
                    Catch ex As Exception
                        Exit Try
                    End Try

                    If strAuthorSchedule <> "" Then
                        objCurrentObject.SchedulingInfo.Properties.Item("SI_AUTHOR").Value = Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Text
                    End If

                    'See if this property exists
                    Dim strAuthor As String = ""
                    Try
                        strAuthor = objCurrentObject.Properties.Item("SI_AUTHOR").Value.ToString()
                    Catch ex As Exception
                        Exit Try
                    End Try

                    If strAuthor <> "" Then
                        objCurrentObject.Properties.Item("SI_AUTHOR").Value = Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Text
                    End If

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
                    objCurrentObject.Properties.Item("SI_UPDATE_TS").Value = Date.Now

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

        UpdateOwnerOnObject(Me.txtReplaceOwnerOnSingleDocDocId.Text, Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Text)

    End Sub

    Private Sub btnUpdateOwnersOnDocsFromCSV_Click(sender As Object, e As EventArgs) Handles btnUpdateOwnersOnDocsFromCSV.Click

        Dim strDocID As String
        Dim strNewOwner As String
        Dim txtOutput As String()

        Using MyReader As New Microsoft.VisualBasic.
                      FileIO.TextFieldParser(
                        "C:\SAP\ListOfObjectsAndOwner.txt")

            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")

            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    strDocID = currentRow(0)
                    strNewOwner = currentRow(1)

                    txtOutput = {"Found record to process from CSV: SI_ID: ", strDocID, ", Owner: ", strNewOwner, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))

                    UpdateOwnerOnObject(strDocID, strNewOwner)

                Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                    txtOutput = {"Line " & ex.Message & " is not valid and will be skipped. ChrW(13) & ChrW(10)"}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))
                End Try

            End While
        End Using

    End Sub

    Private Sub UpdateOwnerOnObject(strDocID As String, strNewOwner As String)

        Me.NewBOSession()

        Dim intOwnerIdNew As Integer = Me.GetIdForUser(strNewOwner)

        If (intOwnerIdNew > 0) Then

            Dim strQuery As String = ("select * from ci_infoobjects where si_id = " & strDocID)
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
                        objCurrentObject.Properties.Item("SI_UPDATE_TS").Value = Date.Now

                        Dim txtOutput As String() = New String() {"Object updated: (", strObjectKind, ") ", strObjectName, ChrW(13) & ChrW(10)}
                        Me.rtbOutput.AppendText(String.Concat(txtOutput))

                        Me.boInfoStore.Commit(infoObjects)
                    Loop
                Finally
                    If TypeOf enumerator Is IDisposable Then
                        TryCast(enumerator, IDisposable).Dispose()
                    End If
                End Try

            Else
                Me.rtbOutput.AppendText("Unable to find ID of new Owner in system" & ChrW(13) & ChrW(10))
            End If

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

        Me.GetBOUserList(True)

    End Sub

    Private Sub btnLoadListOfUsersToDB_Click(sender As Object, e As EventArgs) Handles btnLoadListOfUsersToDB.Click

        GetBOUserList(False, Me.txtLoadListOfUsersToDBDatabaseName.Text.ToString(), Me.txtLoadListOfUsersToDBSQLServerName.Text.ToString(), Me.txtSIUserId.Text.ToString())

    End Sub

    Protected Sub GetBOUserList(blnDisplay As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "", Optional strSIID As String = "")


        logger.Debug("Begin GetBOUserList")

        Dim strQuery As String

        Me.NewBOSession()

        If strSIID <> "" Then
            strQuery = ("Select TOP 10000 * FROM CI_SYSTEMOBJECTS Where SI_KIND='User' AND SI_ID = " + strSIID + " ORDER BY SI_ID")
            logger.Debug("Finished building query for GetBOUserList for specific SIID")
        Else
            strQuery = ("Select TOP 10000 * FROM CI_SYSTEMOBJECTS Where SI_KIND='User' ORDER BY SI_ID")
            logger.Debug("Finished building query for GetBOUserList for ALL")
        End If

        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)


        If blnDisplay Then
            Dim arrResults As String() = New String() {"Count of users found: (", objects.Count.ToString(), ChrW(13) & ChrW(10)}
            Me.rtbOutput.AppendText(String.Concat(arrResults))
        End If

        logger.Debug("Queried BI Platform for list of user(s): Users found: " + objects.Count.ToString)

        If (objects.Count > 0) Then

            SetSQLConnection(strDatabaseName, strSQLServerName)

            'Create stage table
            CreateUserTableForRepo()

            Dim strUserId As String = ""
            Dim strCUID As String = ""
            Dim strName As String = ""
            Dim strFullName As String = ""
            Dim strDescription As String = ""
            Dim strEmailAddress = ""

            Dim enumerator As IEnumerator
            Try
                enumerator = objects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)

                    Try
                        strUserId = current.Properties.Item("SI_ID").Value.ToString()
                        strCUID = current.Properties.Item("SI_CUID").Value.ToString()
                        strName = current.Properties.Item("SI_NAME").Value.ToString()
                        strFullName = current.Properties.Item("SI_USERFULLNAME").Value.ToString()
                        strDescription = current.Properties.Item("SI_DESCRIPTION").Value.ToString()
                        strEmailAddress = current.Properties.Item("SI_EMAIL_ADDRESS").Value.ToString()
                    Catch ex As Exception
                        logger.[Error](ex, "ow noos! Error in GetBOUserList")
                    End Try

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

                    'If we've received db info, load it, else skip
                    If strDatabaseName <> "" Then
                        LoadUserToDatabaseStg(strUserId, strCUID, strName, strFullName, strDescription, strEmailAddress, strLastLogonTime, strDisabled)
                    End If

                Loop
            Catch ex As Exception

                logger.[Error](ex, "ow noos! Error in GetBOUserList")
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        Else
            logger.Debug("No user(s) found")
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
                        logger.[Error](ex, "ow noos! Error in CheckIfUserIsDisabled")
                    Finally
                        If TypeOf objUserEnumerator Is IDisposable Then
                            TryCast(objUserEnumerator, IDisposable).Dispose()
                        End If

                    End Try

                Loop

            Catch ex As Exception
                logger.[Error](ex, "ow noos! Error in CheckIfUserIsDisabled")
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

    Private Sub CreateUserTableForRepo()

        logger.Debug("CreateUserTableForRepo: Begin to create table if not exists: DimSAPBOUser_Stg")
        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOUser_Stg' and xtype='U') " _
                        & "create table DimSAPBOUser_Stg (" _
                        & "    SI_ID int not null primary key" _
                        & "   ,SI_CUID varchar(64) Not null" _
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

        logger.Debug("CreateUserTableForRepo: Begin to truncate table DimSAPBOUser_Stg")
        ExecuteQuery("truncate table DimSAPBOUser_Stg")

    End Sub


    Private Sub CreateObjectTableForRepo()

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOObject_Stg' and xtype='U') " _
                        & "create table DimSAPBOObject_Stg (" _
                        & "    SI_ID int not null primary key" _
                        & "   ,SI_CUID varchar(64) not null" _
                        & "   ,SI_NAME varchar(255) null" _
                        & "   ,SI_OWNER varchar(255) null" _
                        & "   ,SI_PARENT_FOLDER int null" _
                        & "   ,SI_INSTANCE bit null default 0" _
                        & "   ,SI_SCHEDULE_STATUS tinyint not null default -1" _
                        & "   ,SI_SIZE int null" _
                        & "   ,SI_PARENTID int null" _
                        & "   ,SI_UPDATE_TS datetime null" _
                        & "   ,SI_CREATION_TIME datetime null" _
                        & "   ,SI_KIND varchar(255) null" _
                        & "   ,SI_HAS_CHILDREN bit null default 0" _
                        & "   ,SI_RECURRING bit null default 0" _
                        & "   ,SI_KEYWORD varchar(255) null" _
                        & "   ,SI_IS_PUBLICATION_JOB bit null default 0" _
                        & "   ,RecordInsertTimestamp datetime2(7) not null default sysdatetime()" _
                        & "   ,RecordUpdateTimestamp datetime2(7) not null default sysdatetime()" _
                        & " )")

        ExecuteQuery("truncate table DimSAPBOObject_Stg")

    End Sub



    Private Sub CreateObjectTablePropertyRepoStage()

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOObjectProperty_Stg' and xtype='U') " _
                        & "create table DimSAPBOObjectProperty_Stg (" _
                        & "    SI_ID int not null" _
                        & "   ,ClassName varchar(64) not null" _
                        & "   ,PropertyName varchar(255) null" _
                        & "   ,PropertyValueInstanceNumber int null" _
                        & "   ,PropertyValue varchar(max) null" _
                        & " )")

        ExecuteQuery("truncate table DimSAPBOObjectProperty_Stg")

    End Sub


    Private Sub CreateObjectTablePropertyRepo()

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sysobjects where name='DimSAPBOObjectProperty' and xtype='U') " _
                        & "CREATE TABLE [dbo].[DimSAPBOObjectProperty]" _
                        & "(" _
                        & "[ObjectPropertyKey] Int Not NULL IDENTITY(1, 1) PRIMARY KEY," _
                        & "[SI_ID] [Int] Not NULL," _
                        & "[ClassName] [varchar](64) COLLATE SQL_Latin1_General_CP1_CI_AS Not NULL," _
                        & "[PropertyName] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," _
                        & "[PropertyValueInstanceNumber] [Int] NULL," _
                        & "[PropertyValue] [varchar](max) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," _
                        & "SysStartTime datetime2 GENERATED ALWAYS As ROW START Not NULL," _
                        & "SysEndTime datetime2 GENERATED ALWAYS As ROW End Not NULL," _
                        & "PERIOD FOR SYSTEM_TIME (SysStartTime, SysEndTime) " _
                        & ") " _
                        & "With (SYSTEM_VERSIONING = ON (HISTORY_TABLE = dbo.DimSAPBOObjectProperty_History))")

    End Sub

    Private Sub LoadUserToDatabaseStg(strID As String, strCUID As String, strName As String, strFullName As String, strDescription As String, strEmailAddress As String, strLastLogonTime As String, strDisabled As String)

        Dim strQuery As String

        strQuery = "INSERT INTO dbo.DimSAPBOUser_Stg" _
                 & "     ( SI_ID" _
                 & "      ,SI_CUID" _
                 & "      ,SI_NAME" _
                 & "      ,SI_USERFULLNAME" _
                 & "      ,SI_DESCRIPTION" _
                 & "      ,SI_EMAIL_ADDRESS" _
                 & "      ,SI_LASTLOGONTIME" _
                 & "      ,Disabled" _
                 & "     )" _
                 & "VALUES" _
                 & "     (" _
                 & "          " + strID + "                  " _
                 & "         ,'" + strCUID + "'              " _
                 & "         ,'" + Replace(strName, "'", "''") + "'              " _
                 & "         ,'" + Replace(strFullName, "'", "''") + "'             " _
                 & "         ,'" + Replace(strDescription, "'", "''") + "'        " _
                 & "         ,'" + strEmailAddress + "'            " _
                 & "         ,'" + strLastLogonTime + "'                " _
                 & "         ," + strDisabled + "            " _
                 & "     )"

        ExecuteQuery(strQuery)

    End Sub

    Private Sub LoadObjectToDatabaseStg(strId As String, strCUID As String, strName As String, strOwner As String, intParentFolder As String, blnInstance As String, intScheduleStatus As String, intSize As String, intParentId As String, dteUpdateTimestamp As String, dteCreationTimestamp As String, strKind As String, blnHasChildren As String, blnRecurring As String, strKeyword As String, blnIsPublication As String)

        Dim strQuery As String

        strQuery = "INSERT INTO dbo.DimSAPBOObject_Stg" _
                 & "     ( SI_ID" _
                 & "      ,SI_CUID" _
                 & "      ,SI_NAME" _
                 & "      ,SI_OWNER" _
                 & "      ,SI_PARENT_FOLDER" _
                 & "      ,SI_INSTANCE" _
                 & "      ,SI_SCHEDULE_STATUS" _
                 & "      ,SI_SIZE" _
                 & "      ,SI_PARENTID" _
                 & "      ,SI_UPDATE_TS" _
                 & "      ,SI_CREATION_TIME" _
                 & "      ,SI_KIND" _
                 & "      ,SI_HAS_CHILDREN" _
                 & "      ,SI_RECURRING" _
                 & "      ,SI_KEYWORD" _
                 & "      ,SI_IS_PUBLICATION_JOB" _
                 & "     )" _
                 & "VALUES" _
                 & "     (" _
                 & "          " + strId + "                  " _
                 & "         ,'" + strCUID + "'              " _
                 & "         ,'" + Replace(strName, "'", "''") + "'              " _
                 & "         ,'" + Replace(strOwner, "'", "''") + "'             " _
                 & "         ," + intParentFolder + "        " _
                 & "         ," + blnInstance + "            " _
                 & "         ," + intScheduleStatus + "      " _
                 & "         ," + intSize + "                " _
                 & "         ," + intParentId + "            " _
                 & "         ,'" + dteUpdateTimestamp + "'   " _
                 & "         ,'" + dteCreationTimestamp + "' " _
                 & "         ,'" + strKind + "'              " _
                 & "         ," + blnHasChildren + "         " _
                 & "         ," + blnRecurring + "           " _
                 & "         ,'" + strKeyword + "'              " _
                 & "         ," + blnIsPublication + "           " _
                 & "     )"

        ExecuteQuery(strQuery)

    End Sub

    Private Sub BuildLoadToDatabaseObjectPropertyStgString(strId As String, strClassName As String, strPropertyName As String, intPropertyValueInstanceNumber As String, strPropertyValue As String)

        ' Add some new rows to the collection.
        Dim row As DataRow
        row = myBOProperties.NewRow()
        row("SI_ID") = strId
        row("ClassName") = strClassName
        row("PropertyName") = strPropertyName
        row("PropertyValueInstanceNumber") = intPropertyValueInstanceNumber
        row("PropertyValue") = strPropertyValue
        myBOProperties.Rows.Add(row)

        myBOProperties.AcceptChanges()


    End Sub
    Private Sub LoadToDatabaseObjectPropertyStgString()

        Try

            conSQLConn.Open()

            Using bulkCopy As SqlBulkCopy =
              New SqlBulkCopy(conSQLConn)
                bulkCopy.DestinationTableName = "dbo.DimSAPBOObjectProperty_Stg"

                Try
                    ' Write from the source to the destination.
                    bulkCopy.WriteToServer(myBOProperties)

                Catch ex As Exception
                    logger.[Error](ex, "ow noos! Error in LoadToDatabaseObjectPropertyStgString")
                End Try
            End Using

            conSQLConn.Close()

        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error in LoadToDatabaseObjectPropertyStgString")
            Exit Try
        End Try
        myBOProperties = Nothing

    End Sub
    Private Sub SetSQLConnection(strDB As String, strServer As String)

        'Create a Connection object.
        conSQLConn = New SqlConnection("Initial Catalog=" + strDB + ";Data Source=" + strServer + ";Integrated Security=SSPI;")

    End Sub

    Private Sub ExecuteQuery(strQuery As String)

        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error in ExecuteQuery")
            Exit Try
        End Try

        conSQLConn.Close()

    End Sub

    Private Sub btnExtractRepo_Click(sender As Object, e As EventArgs) Handles btnExtractRepo.Click

        GetBOObjectList(False, txtExtractRepoDatabaseName.Text.ToString(), txtExtractRepoSQLServer.Text.ToString())

    End Sub

    Protected Sub GetBOObjectList(blnDisplay As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "")

        Me.NewBOSession()

        Dim strQuery As String
        Dim strSIID As String = Me.txtSIID.Text.ToString()
        If strSIID <> "" Then
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox') AND SI_ID = " + strSIID)
        Else
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox')")
        End If

        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        logger.[Debug]("GetBOObjectList: Queried CI_INFOOBJECTS and found records: " + objects.Count.ToString())

        If (objects.Count > 0) Then

            SetSQLConnection(strDatabaseName, strSQLServerName)

            'Create stage table
            CreateObjectTableForRepo()

            logger.[Debug]("GetBOObjectList: Finished creating/truncating the staging table")

            Dim enumerator As IEnumerator
            Try
                enumerator = objects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim strId As String = current.Properties.Item("SI_ID").Value.ToString()
                    Dim strCUID As String = current.Properties.Item("SI_CUID").Value.ToString()
                    Dim strName As String = current.Properties.Item("SI_NAME").Value.ToString()
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

                    'This property may not exist
                    Dim intScheduleStatus As String
                    Try
                        intScheduleStatus = current.Properties.Item("SI_SCHEDULE_STATUS").Value.ToString()
                    Catch ex As Exception
                        intScheduleStatus = -1
                        Exit Try
                    End Try

                    'This property may not exist
                    Dim strKeyword As String
                    Try
                        strKeyword = current.Properties.Item("SI_KEYWORD").Value.ToString()
                    Catch ex As Exception
                        strKeyword = ""
                        Exit Try
                    End Try

                    'This property may not exist
                    Dim blnIsPublication As String
                    Try

                        blnIsPublication = If(current.Properties.Item("SI_IS_PUBLICATION_JOB").Value.ToString() = "False", "0", "1")
                    Catch ex As Exception
                        blnIsPublication = 0
                        Exit Try
                    End Try

                    LoadObjectToDatabaseStg(strId, strCUID, strName, strOwner, intParentFolder, blnInstance, intScheduleStatus, intSize, intParentId, dteUpdateTimestamp, dteCreationTimestamp, strKind, blnHasChildren, blnRecurring, strKeyword, blnIsPublication)

                Loop
            Catch ex As Exception
                logger.[Error](ex, "ow noos! Error while querying and parsing Infostore for a list of objects")
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If

        Me.LogoffBOSession()

    End Sub

    Protected Sub GetBOObjectProperties(blnDisplay As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "", Optional blnDeltas As Boolean = 1)

        Me.NewBOSession()

        Dim strQuery As String
        Dim strProperty1 As String
        Dim strProperty2 As String
        Dim strProperty3 As String
        Dim strProperty4 As String
        Dim strProperty5 As String
        Dim strProperty6 As String
        Dim strProperty7 As String
        Dim strPropertyValue1 As String
        Dim strPropertyValue2 As String
        Dim strPropertyValue3 As String
        Dim strPropertyValue4 As String
        Dim strPropertyValue5 As String
        Dim strPropertyValue6 As String
        Dim strPropertyValue7 As String
        Dim strClass As String
        Dim strId As String
        Dim strSIID As String = Me.txtLoadObjectPropertiesSIID.Text.ToString()

        If strSIID <> "" Then
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox') AND SI_ID = " + strSIID)
        Else
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox')")
        End If

        SetSQLConnection(strDatabaseName, strSQLServerName)

        If blnDeltas Then
            Dim strLastUpdateTimestamp = GetLoadObjectPropertiesDeltaTimestamp()
            strQuery = strQuery + " AND SI_UPDATE_TS > '" + strLastUpdateTimestamp + "'"
        End If

        CreateObjectTablePropertyRepo()

        Dim myInfoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If myInfoObjects.Count > 0 Then

            CreateObjectTablePropertyRepoStage()
            MakeDataTableForObjectPropertyList()

            For iLoop = 1 To myInfoObjects.Count

                strId = myInfoObjects(iLoop).Properties.Item("SI_ID").Value.ToString()

                If myInfoObjects(iLoop).Properties.Count > 0 Then
                    strClass = "Properties"
                    For iProp1 = 1 To myInfoObjects(iLoop).Properties.Count

                        strProperty1 = myInfoObjects(iLoop).Properties(iProp1).Name.ToString()
                        strPropertyValue1 = myInfoObjects(iLoop).Properties(iProp1).Value.ToString()
                        If (strProperty1 <> "SI_ID") And (strPropertyValue1 <> "System.__ComObject") Then
                            BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1, iProp1, strPropertyValue1)
                        End If

                        If myInfoObjects(iLoop).Properties(iProp1).Properties.Count > 0 Then
                            For iProp2 = 1 To myInfoObjects(iLoop).Properties(iProp1).Properties.Count

                                strProperty2 = myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Name.ToString()
                                strPropertyValue2 = myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Value.ToString()
                                If (strPropertyValue2 <> "System.__ComObject") Then
                                    BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)
                                End If

                                If myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Properties.Count > 0 Then
                                    For iProp3 = 1 To myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Properties.Count

                                        strProperty3 = myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Properties(iProp3).Name.ToString()
                                        strPropertyValue3 = myInfoObjects(iLoop).Properties(iProp1).Properties(iProp2).Properties(iProp3).Value.ToString()
                                        If (strPropertyValue3 <> "System.__ComObject") Then
                                            BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)
                                        End If

                                    Next
                                End If

                            Next
                        End If
                    Next
                End If


                If (myInfoObjects(iLoop).Kind <> "Folder") And (myInfoObjects(iLoop).Kind <> "FavoritesFolder") And (myInfoObjects(iLoop).Kind <> "Inbox") Then
                    If myInfoObjects(iLoop).ProcessingInfo.Properties.Count > 0 Then
                        strClass = "Processing"

                        For iProp1 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties.Count

                            strProperty1 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Name.ToString()
                            strPropertyValue1 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Value.ToString()
                            If (strPropertyValue1 <> "SI_ID" And strPropertyValue1 <> "System.__ComObject") Then
                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1, iProp1, strPropertyValue1)
                            End If

                            If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties.Count > 0 Then
                                For iProp2 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties.Count

                                    strProperty2 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Name.ToString()
                                    strPropertyValue2 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Value.ToString
                                    If (strPropertyValue2 <> "System.__ComObject") Then
                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)
                                    End If

                                    If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties.Count > 0 Then
                                        For iProp3 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties.Count

                                            strProperty3 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Name.ToString()
                                            strPropertyValue3 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Value.ToString
                                            If (strPropertyValue3 <> "System.__ComObject") Then
                                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)
                                            End If

                                            If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count > 0 Then
                                                For iProp4 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count

                                                    strProperty4 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Name.ToString()
                                                    strPropertyValue4 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Value.ToString
                                                    If (strPropertyValue4 <> "System.__ComObject") Then
                                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4, iProp4, strPropertyValue4)
                                                    End If

                                                    If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count > 0 Then
                                                        For iProp5 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count

                                                            strProperty5 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Name.ToString()
                                                            strPropertyValue5 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Value.ToString
                                                            If (strPropertyValue5 <> "System.__ComObject") Then
                                                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5, iProp5, strPropertyValue5)
                                                            End If

                                                            If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count > 0 Then
                                                                For iProp6 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count

                                                                    strProperty6 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Name.ToString()
                                                                    strPropertyValue6 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Value.ToString()
                                                                    If (strPropertyValue6 <> "System.__ComObject") Then
                                                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6, iProp6, strPropertyValue6)
                                                                    End If

                                                                    If myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties.Count > 0 Then
                                                                        For iProp7 = 1 To myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties.Count

                                                                            strProperty7 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties(iProp7).Name.ToString()
                                                                            strPropertyValue7 = myInfoObjects(iLoop).ProcessingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties(iProp7).Value.ToString
                                                                            If (strPropertyValue7 <> "System.__ComObject") Then
                                                                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6 + strProperty7, iProp7, strPropertyValue7)
                                                                            End If

                                                                        Next
                                                                    End If

                                                                Next
                                                            End If

                                                        Next
                                                    End If

                                                Next
                                            End If

                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If


                    If myInfoObjects(iLoop).SchedulingInfo.Properties.Count > 0 Then
                        For iProp1 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties.Count
                            strClass = "Scheduling"

                            strProperty1 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Name.ToString()
                            strPropertyValue1 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Value.ToString()
                            If (strProperty1 <> "SI_ID") And (strPropertyValue1 <> "System.__ComObject") Then
                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1, iProp1, strPropertyValue1)
                            End If


                            If myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties.Count > 0 Then
                                For iProp2 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties.Count

                                    strProperty2 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Name.ToString()
                                    strPropertyValue2 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Value.ToString
                                    If (strPropertyValue2 <> "System.__ComObject") Then
                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)
                                    End If

                                    If myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties.Count > 0 Then
                                        For iProp3 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties.Count

                                            strProperty3 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Name.ToString()
                                            strPropertyValue3 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Value.ToString
                                            If (strPropertyValue3 <> "System.__ComObject") Then
                                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)
                                            End If

                                            If myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count > 0 Then
                                                For iProp4 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count

                                                    strProperty4 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Name.ToString()
                                                    strPropertyValue4 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Value.ToString
                                                    If (strPropertyValue4 <> "System.__ComObject") Then
                                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4, iProp4, strPropertyValue4)
                                                    End If

                                                    If myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count > 0 Then
                                                        For iProp5 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count

                                                            strProperty5 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Name.ToString()
                                                            strPropertyValue5 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Value.ToString
                                                            If (strPropertyValue5 <> "System.__ComObject") Then
                                                                BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5, iProp5, strPropertyValue5)
                                                            End If

                                                            If myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count > 0 Then
                                                                For iProp6 = 1 To myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count

                                                                    strProperty6 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Name.ToString()
                                                                    strPropertyValue6 = myInfoObjects(iLoop).SchedulingInfo.Properties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Value.ToString
                                                                    If (strPropertyValue6 <> "System.__ComObject") Then
                                                                        BuildLoadToDatabaseObjectPropertyStgString(strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6, iProp6, strPropertyValue6)
                                                                    End If
                                                                Next
                                                            End If

                                                        Next
                                                    End If

                                                Next
                                            End If

                                        Next
                                    End If


                                Next
                            End If


                        Next
                    End If

                End If

            Next

            'Finally, write to SQL Server
            LoadToDatabaseObjectPropertyStgString()

            MergeObjectPropertyList(blnDeltas)

        End If
        Me.LogoffBOSession()

    End Sub


    Private Sub btnAddDBCredentialsForGroup_Click(sender As Object, e As EventArgs) Handles btnAddDBCredentialsForGroup.Click

        Me.NewBOSession()
        Dim txtAddDBCredentialsPW As String = Me.GetIdForUser(Me.txtAddDBCredentialsPW.Text)
        Dim strQuery As String = ("select top 1000000 SI_ID, SI_NAME, SI_DATA, SI_2ND_CREDS from ci_systemobjects where SI_KIND='User' and SI_NAME = 'BrianK'")
        Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)
        Dim txtOutput As String()
        Dim strUserName As String

        If (infoObjects.Count > 0) Then
            Dim enumerator As IEnumerator
            Try
                enumerator = infoObjects.GetEnumerator
                Do While enumerator.MoveNext

                    Dim objCurrentUser As UserInfo = DirectCast(enumerator.Current, UserInfo)

                    strUserName = objCurrentUser.UserName.ToString

                    'See if this property exists
                    Dim strDBUser As String = ""
                    Try
                        strDBUser = objCurrentUser.GetProfileString("DBUSER")
                    Catch ex As Exception
                        Exit Try
                    End Try

                    If strDBUser = "" Then
                        objCurrentUser.SetProfileString("DBUSER", strUserName)
                        objCurrentUser.SetSecondaryCredential("DBPASS", "")

                        txtOutput = {"Object updated: (", strDBUser, ") ", strUserName, ChrW(13) & ChrW(10)}
                        Me.rtbOutput.AppendText(String.Concat(txtOutput))
                    End If

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

    Private Sub btnResetAdminPW_Click(sender As Object, e As EventArgs) Handles btnResetAdminPW.Click
    End Sub

    Private Sub BtnLoadObjectPropertiesToDB_Click(sender As Object, e As EventArgs) Handles btnLoadObjectPropertiesToDB.Click
        GetBOObjectProperties(False, txtLoadObjectPropertiesDB.Text.ToString(), txtLoadObjectPropertiesServer.Text.ToString(), CBool(txtLoadObjectPropertyDelta.Text.ToString))
    End Sub

    Private Function GetLoadObjectPropertiesDeltaTimestamp() As String

        InitializeETLParmForObjectPropertyDelta()

        Dim strQuery As String = "SELECT ParmValue FROM dbo.ETLParm WHERE ParmName = 'DeltaCMSObjectProperty'"
        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)
        Dim dtParmValue As Date

        Try
            conSQLConn.Open()

            dtParmValue = CDate(cmdSQLCmd.ExecuteScalar())

        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error in GetLoadObjectPropertiesDeltaTimestamp")
            Exit Try
        End Try

        conSQLConn.Close()

        SetLoadObjectPropertiesDeltaTimestamp()

        GetLoadObjectPropertiesDeltaTimestamp = ConvertDateToString(dtParmValue)

    End Function

    Private Sub SetLoadObjectPropertiesDeltaTimestamp()

        Dim strQuery As String = "UPDATE dbo.ETLParm SET ParmValue = '" + DateTime.Now.AddMinutes(-1).ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE ParmName = 'DeltaCMSObjectProperty'"
        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error in SetLoadObjectPropertiesDeltaTimestamp")
            Exit Try
        End Try

        conSQLConn.Close()
    End Sub

    Private Function ConvertDateToString(dtDateValue As Date) As String

        ConvertDateToString = dtDateValue.ToString("yyyy.MM.dd.HH.mm.ss")

    End Function

    Private Sub InitializeETLParmForObjectPropertyDelta()
        Dim strQuery As String = "IF NOT EXISTS (SELECT 1 FROM dbo.ETLParm WHERE BI_DBID = 0 AND DBID = 0 AND ParmName = 'DeltaCMSObjectProperty') INSERT INTO dbo.ETLParm (BI_DBID,DBID,ParmName,ParmValue) VALUES (0,0,'DeltaCMSObjectProperty','1901-01-01')"
        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.[Error](ex, "ow noos! Error in SetLoadObjectPropertiesDeltaTimestamp")
            Exit Try
        End Try

        conSQLConn.Close()
    End Sub

    Private Sub BtnCheckDeltaTsp_Click(sender As Object, e As EventArgs) Handles btnCheckDeltaTsp.Click
        SetSQLConnection(txtLoadObjectPropertiesDB.Text.ToString(), txtLoadObjectPropertiesServer.Text.ToString())
        Me.txtDeltaTsp.Text = GetLoadObjectPropertiesDeltaTimestamp()
        SetLoadObjectPropertiesDeltaTimestamp()
    End Sub

    Private Function MakeDataTableForObjectPropertyList() As DataTable

        myBOProperties = New DataTable("MyBOProperties")

        ' Add three column objects to the table.
        Dim intSI_ID As DataColumn = New DataColumn()
        intSI_ID.DataType = System.Type.GetType("System.Int32")
        intSI_ID.ColumnName = "SI_ID"
        myBOProperties.Columns.Add(intSI_ID)

        Dim strClassName As DataColumn = New DataColumn()
        strClassName.DataType = System.Type.GetType("System.String")
        strClassName.ColumnName = "ClassName"
        myBOProperties.Columns.Add(strClassName)

        Dim strPropertyName As DataColumn = New DataColumn()
        strPropertyName.DataType = System.Type.GetType("System.String")
        strPropertyName.ColumnName = "PropertyName"
        myBOProperties.Columns.Add(strPropertyName)

        Dim intPropertyValueInstanceNumber As DataColumn = New DataColumn()
        intPropertyValueInstanceNumber.DataType = System.Type.GetType("System.Int32")
        intPropertyValueInstanceNumber.ColumnName = "PropertyValueInstanceNumber"
        myBOProperties.Columns.Add(intPropertyValueInstanceNumber)

        Dim strPropertyValue As DataColumn = New DataColumn()
        strPropertyValue.DataType = System.Type.GetType("System.String")
        strPropertyValue.ColumnName = "PropertyValue"
        myBOProperties.Columns.Add(strPropertyValue)

        ' Return the new DataTable.
        Return myBOProperties

    End Function

    Private Sub MergeObjectPropertyList(blnDelta As Boolean)

        Dim strQuery As String
        strQuery = ""
        strQuery = strQuery & "MERGE dbo.DimSAPBOObjectProperty tgt "
        strQuery = strQuery & "Using dbo.DimSAPBOObjectProperty_Stg src "
        strQuery = strQuery & " "
        strQuery = strQuery & "On ( tgt.SI_ID = src.SI_ID "
        strQuery = strQuery & "     And tgt.ClassName = src.ClassName "
        strQuery = strQuery & "     And tgt.PropertyName = src.PropertyName "
        strQuery = strQuery & "     And tgt.PropertyValueInstanceNumber = src.PropertyValueInstanceNumber ) "
        strQuery = strQuery & " "
        strQuery = strQuery & "WHEN MATCHED "
        strQuery = strQuery & "     And "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               NULLIF(src.PropertyValue, tgt.PropertyValue) Is Not NULL "
        strQuery = strQuery & "          Or NULLIF(tgt.PropertyValue, src.PropertyValue) Is Not NULL "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "THEN UPDATE SET "
        strQuery = strQuery & "          PropertyValue = src.PropertyValue "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "            SI_ID "
        strQuery = strQuery & "           ,ClassName "
        strQuery = strQuery & "           ,PropertyName "
        strQuery = strQuery & "           ,PropertyValueInstanceNumber "
        strQuery = strQuery & "           ,PropertyValue ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, ClassName, PropertyName "
        strQuery = strQuery & "              ,PropertyValueInstanceNumber, PropertyValue "
        strQuery = strQuery & "          ) "
        If blnDelta Then
            strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
            strQuery = strQuery & "THEN DELETE "
        End If
        strQuery = strQuery & ";"

        ExecuteQuery(strQuery)

    End Sub
End Class
