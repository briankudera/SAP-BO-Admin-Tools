﻿Imports CrystalDecisions.Enterprise
Imports System.Data.SqlClient
Imports CrystalDecisions.Enterprise.Desktop
Imports log4net
Imports File = System.IO.File
Imports System.Globalization
Imports System.IO
Imports System.Security

<Assembly: log4net.Config.XmlConfigurator(
      ConfigFile:="Log4Net.config", Watch:=True)>

Public Class frmTools

    Private Shared ReadOnly logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private boEnterpriseSession As EnterpriseSession
    Private boInfoStore As InfoStore
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
    Private cmdQuery As String = ""
    Private cmdSIID As String = ""



    Private Sub Tools_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Config.XmlConfigurator.Configure()

        Dim myFileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        log4net.GlobalContext.Properties("VersionCode") = myFileVersionInfo.FileVersion.ToString

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


        Try
            If arguments.Count = 9 Then
                cmdCommandLine = arguments(0)
                cmdAction = arguments(1)
                cmdCMSServer = arguments(2)
                cmdCMSUser = arguments(3)
                cmdCMSUserPassword = arguments(4)
                cmdCMSAuthentication = arguments(5)
                cmdTargetServer = arguments(6)
                cmdTargetDB = arguments(7)
                cmdDeltaProcessing = arguments(8)
            ElseIf arguments.Count = 11 Then

                cmdQuery = arguments(9)
                cmdSIID = arguments(10)
            End If
        Catch ex As Exception
            logger.Error("Program called via console but error parsing args! Delta processing arg: " & cmdDeltaProcessing, ex)
        End Try

        logger.Info("Program called via console. Program: " + cmdAction)

        Try
            If cmdAction = "LoadUsersToDB" Then
                GetBOUserList(False, cmdTargetDB, cmdTargetServer)
                GetBOUserList(False, cmdTargetDB, cmdTargetServer)
            ElseIf cmdAction = "LoadObjectsToDB" Then
                GetBOObjectList(False, cmdTargetDB, cmdTargetServer)
            ElseIf cmdAction = "LoadObjectPropertiesToDB" Then
                GetBOObjectProperties(False, cmdTargetDB, cmdTargetServer, cmdDeltaProcessing, cmdSIID)
            End If
        Catch ex As Exception
            logger.Error("Program called via console but failed calling sub routine", ex)
        End Try

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

        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText("************************************************************************" & ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText("Opening session to SAP BO." & ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))

        Try
            Me.boEnterpriseSession = mgr.Logon(strUserName, strPassword, strCMSName, strAuthentication)
        Catch ex As Exception
            logger.Error("ow noos! Error while establishing a BO Enterprise Session", ex)
            Application.Exit()
        End Try

        Try
            Me.boInfoStore = New InfoStore(Me.boEnterpriseSession.GetService("InfoStore"))
        Catch ex As Exception

            logger.Error("ow noos! Error while establishing a new InfoStore", ex)
            Application.Exit()
        End Try

        logger.Debug("Finish open session to BI Platform with session CUID: " + boEnterpriseSession.SessionCUID.ToString())

    End Sub

    Private Sub LogoffBOSession()

        logger.Debug("Begin close session to BI Platform")

        Me.boEnterpriseSession.Logoff()
        Me.boEnterpriseSession.Dispose()
        Me.boInfoStore.Dispose()
        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText("Closing session to SAP BO." & ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText("************************************************************************" & ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        'Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))

    End Sub


    Private Function GetIdForUser(ByVal strUserName As String) As Integer

        Dim strUserId As String = ""
        Dim intUserId As Integer

        Dim strQuery As String = ("select SI_ID from CI_SYSTEMOBJECTS WHERE SI_KIND = 'User' and SI_NAME = '" & strUserName & "'")
        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)
        If (objects.ResultCount = 1) Then
            strUserId = objects.Item(1).Properties.Item("SI_ID").Value.ToString
        End If
        intUserId = GetIntegerFromString(strUserId)
        Return intUserId

    End Function

    Private Function GetIntegerFromString(ByVal strString As String) As Integer
        Dim intStringToInt As Integer

        If strString.Trim = "" Then
            strString = "0"
        End If

        If Not Integer.TryParse(strString.Trim, intStringToInt) Then
            intStringToInt = -1
        End If
        Return intStringToInt
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
        Dim intNewOwner As Integer = 0
        Dim txtOutput As String()
        Dim blnOwnerIsSameOnAllDocs As Boolean
        Dim blnRetrievedNewOwnerID As Boolean = False

        blnOwnerIsSameOnAllDocs = Me.chkOwnerIsSameOnAllDocs.Checked

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

                    If blnOwnerIsSameOnAllDocs Then
                        If Not blnRetrievedNewOwnerID Then
                            intNewOwner = Me.GetIdForUser(strNewOwner)
                            blnRetrievedNewOwnerID = True
                        End If
                    End If

                    txtOutput = {"Found record to process from CSV: SI_ID: ", strDocID, ", Owner: ", strNewOwner, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))

                    UpdateOwnerOnObject(strDocID, strNewOwner, intNewOwner)

                Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                    txtOutput = {"Line " & ex.Message & " is not valid and will be skipped. ChrW(13) & ChrW(10)"}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))
                End Try

            End While
        End Using

    End Sub

    Private Sub UpdateOwnerOnObject(strDocID As String, strNewOwner As String, Optional intNewOwner As Integer = 0)

        Me.NewBOSession()

        Dim intOwnerIdNew As Integer
        If intNewOwner = 0 Then
            intOwnerIdNew = Me.GetIdForUser(strNewOwner)
        Else
            intOwnerIdNew = intNewOwner
        End If

        If (intOwnerIdNew > 0) Then

            Dim blnUpdateMade As Boolean = False

            Dim strQuery As String = ("select * from ci_infoobjects where si_id = " & strDocID)
            Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

            If (infoObjects.Count > 0) Then
                Dim enumerator As IEnumerator
                Try
                    enumerator = infoObjects.GetEnumerator
                    Do While enumerator.MoveNext
                        Dim objCurrentObject As InfoObject = DirectCast(enumerator.Current, InfoObject)
                        Dim strObjectName As String = objCurrentObject.Properties.Item("SI_NAME").Value.ToString
                        Dim strObjectKind As String = objCurrentObject.Properties.Item("SI_KIND").Value.ToString
                        Dim strOwnerId As String = objCurrentObject.Properties.Item("SI_OWNERID").Value.ToString
                        Dim intOwnerId As Integer = Me.GetIntegerFromString(strOwnerId)
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

                            Dim intSubmitterId As Integer
                            intSubmitterId = Me.GetIntegerFromString(strSubmitterId)

                            If intSubmitterId <> intOwnerIdNew Then
                                objCurrentObject.SchedulingInfo.Properties.Item("SI_SUBMITTERID").Value = intOwnerIdNew
                                blnUpdateMade = True
                            End If

                        End If

                        If intOwnerId <> intOwnerIdNew Then
                            objCurrentObject.Properties.Item("SI_OWNERID").Value = intOwnerIdNew
                            blnUpdateMade = True
                        End If

                        If blnUpdateMade Then
                            objCurrentObject.Properties.Item("SI_UPDATE_TS").Value = Date.Now
                            Dim txtOutput As String() = New String() {"Object updated: (", strObjectKind, ") ", strObjectName, ChrW(13) & ChrW(10)}
                            Me.rtbOutput.AppendText(String.Concat(txtOutput))

                            Me.boInfoStore.Commit(infoObjects)
                        Else
                            Dim txtOutput As String() = New String() {"Object not updated, no changes found: (", strObjectKind, ") ", strObjectName, ChrW(13) & ChrW(10)}
                            Me.rtbOutput.AppendText(String.Concat(txtOutput))
                        End If

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
                        logger.Error("ow noos! Error in GetBOUserList", ex)
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

                logger.Error("ow noos! Error in GetBOUserList", ex)
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
                        logger.Error("ow noos! Error in CheckIfUserIsDisabled", ex)
                    Finally
                        If TypeOf objUserEnumerator Is IDisposable Then
                            TryCast(objUserEnumerator, IDisposable).Dispose()
                        End If

                    End Try

                Loop

            Catch ex As Exception
                logger.Error("ow noos! Error in CheckIfUserIsDisabled", ex)
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
                        & "(select 1 from sys.objects where name='DimSAPBOUser_Stg' and type='U') " _
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
                        & "(select 1 from sys.objects where name='DimSAPBOObject_Stg' and type='U') " _
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
                        & "(select 1 from sys.objects where name='DimSAPBOObjectProperty_Stg' and type='U') " _
                        & "create table DimSAPBOObjectProperty_Stg (" _
                        & "    SI_ID BIGINT not null" _
                        & "   ,ClassName varchar(64) not null" _
                        & "   ,PropertyName varchar(255) not null" _
                        & "   ,PropertyValueInstanceNumber int null" _
                        & "   ,PropertyValue varchar(max) null" _
                        & " )")

        ExecuteQuery("truncate table DimSAPBOObjectProperty_Stg")

    End Sub


    Private Sub CreateObjectTablePropertyRepo()

        Dim strQuery As String

        ExecuteQuery("if Not exists " _
                        & "(select 1 from sys.objects where name='DimSAPBOObjectProperty' and type='U') " _
                        & "CREATE TABLE [dbo].[DimSAPBOObjectProperty]" _
                        & "(" _
                        & "[ObjectPropertyKey] BIGINT Not NULL IDENTITY(1, 1) PRIMARY KEY," _
                        & "[SI_ID] [BIGINT] Not NULL," _
                        & "[ClassName] [varchar](64) COLLATE SQL_Latin1_General_CP1_CI_AS Not NULL," _
                        & "[PropertyName] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," _
                        & "[PropertyValueInstanceNumber] [Int] NULL," _
                        & "[PropertyValue] [varchar](max) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," _
                        & "SysStartTime datetime2 GENERATED ALWAYS As ROW START Not NULL," _
                        & "SysEndTime datetime2 GENERATED ALWAYS As ROW End Not NULL," _
                        & "PERIOD FOR SYSTEM_TIME (SysStartTime, SysEndTime) " _
                        & ") " _
                        & "With (SYSTEM_VERSIONING = ON (HISTORY_TABLE = dbo.DimSAPBOObjectProperty_History))")

        ExecuteQuery("IF NOT EXISTS (SELECT 'foo' FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.DimSAPBOObjectProperty') AND name='IX01_DimSAPBOObjectProperty') CREATE UNIQUE INDEX IX01_DimSAPBOObjectProperty On DimSAPBOObjectProperty (SI_ID,ClassName,PropertyName)")

        ExecuteQuery("IF NOT EXISTS (SELECT 'foo' FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.DimSAPBOObjectProperty') AND name='IX02_DimSAPBOObjectProperty') CREATE INDEX IX02_DimSAPBOObjectProperty On DimSAPBOObjectProperty (SI_ID,ClassName)")

        ExecuteQuery("IF NOT EXISTS (SELECT 'foo' FROM sys.schemas WHERE name = 'BI') EXEC('CREATE SCHEMA [BI]')")


        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_PromptName') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_PromptName] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "PromptNum Int Not NULL, "
        strQuery = strQuery & "PromptName VARCHAR(2000) Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE(), "
        strQuery = strQuery & "PRIMARY KEY(SI_ID, PromptNum) "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_PromptValue') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_PromptValue] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "PromptNum Int Not NULL, "
        strQuery = strQuery & "PromptValue VARCHAR(400) Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE(), "
        strQuery = strQuery & "PRIMARY KEY(SI_ID, PromptNum, PromptValue) "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_InstanceLastRunDateTime') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_InstanceLastRunDateTime] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL PRIMARY KEY, "
        strQuery = strQuery & "Parent_SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "InstanceLastRunDateTime DateTime Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE() "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_ListOfEmailBcc') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_ListOfEmailBcc] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "EmailAddressBcc VARCHAR(200) Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE(), "
        strQuery = strQuery & "PRIMARY KEY(SI_ID, EmailAddressBcc) "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_ListOfEmailCc') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_ListOfEmailCc] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "EmailAddressCc VARCHAR(200) Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE(), "
        strQuery = strQuery & "PRIMARY KEY(SI_ID, EmailAddressCc) "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_ListOfEmailTo') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_ListOfEmailTo] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL, "
        strQuery = strQuery & "EmailAddressTo VARCHAR(800) Not NULL, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE(), "
        strQuery = strQuery & "PRIMARY KEY(SI_ID, EmailAddressTo) "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_NonRecurring_ListOfIDs') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_NonRecurring_ListOfIDs] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL PRIMARY KEY, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE() "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "If Not exists(select 1 from sys.objects where object_id = object_id('BI.DimBOObject_ScheduledReport_Recurring_ListOfIDs') and type='U') "
        strQuery = strQuery & "CREATE TABLE [BI].[DimBOObject_ScheduledReport_Recurring_ListOfIDs] "
        strQuery = strQuery & "( "
        strQuery = strQuery & "SI_ID BIGINT Not NULL PRIMARY KEY, "
        strQuery = strQuery & "RecordInsertTimestamp DATETIME2(7) Not NULL DEFAULT GETDATE() "
        strQuery = strQuery & ") "
        ExecuteQuery(strQuery)


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

    Private Sub BuildLoadToDatabaseObjectPropertyStgString(ByRef myDT As DataTable, strId As String, strClassName As String, strPropertyName As String, intPropertyValueInstanceNumber As String, strPropertyValue As String)

        If strPropertyValue <> "" And strPropertyValue <> "SI_ID" And strPropertyValue <> "System.__ComObject" And InStr(1, strPropertyName, "SI_INDEX") = 0 And InStr(1, strPropertyName, "SI_MULTIINTERVAL") = 0 And InStr(1, strPropertyName, "SI_DYNAMIC") = 0 And InStr(1, strPropertyName, "SI_INTERVAL") = 0 And InStr(1, strPropertyName, "SI_COMPLEX") = 0 And InStr(1, strPropertyName, "SI_DATAPROVIDERS") = 0 Then
            ' Add some new rows to the collection.
            Dim row As DataRow
            row = myDT.NewRow()
            row("SI_ID") = strId
            row("ClassName") = strClassName
            row("PropertyName") = strPropertyName
            row("PropertyValueInstanceNumber") = intPropertyValueInstanceNumber
            row("PropertyValue") = strPropertyValue
            myDT.Rows.Add(row)
        End If

    End Sub
    Private Sub LoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects As DataTable)

        Try

            conSQLConn.Open()

            Using bulkCopy As SqlBulkCopy =
              New SqlBulkCopy(conSQLConn)
                bulkCopy.DestinationTableName = "dbo.DimSAPBOObjectProperty_Stg"
                bulkCopy.BulkCopyTimeout = 3600
                Try
                    ' Write from the source to the destination.
                    bulkCopy.WriteToServer(myDataTableOfInfoObjects)

                Catch ex As Exception
                    logger.Error("ow noos! Error in LoadToDatabaseObjectPropertyStgString", ex)
                End Try
            End Using

            conSQLConn.Close()

        Catch ex As Exception
            logger.Error("ow noos! Error in LoadToDatabaseObjectPropertyStgString", ex)
            Exit Try
        End Try

    End Sub
    Private Sub SetSQLConnection(strDB As String, strServer As String)

        'Create a Connection object.
        conSQLConn = New SqlConnection("Initial Catalog=" + strDB + ";Data Source=" + strServer + ";Integrated Security=SSPI;")

    End Sub

    Private Sub ExecuteQuery(strQuery As String)

        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)
        cmdSQLCmd.CommandTimeout = 300

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.Error("ow noos! Error in ExecuteQuery: " & strQuery, ex)
            Exit Try
        End Try

        conSQLConn.Close()

    End Sub

    Private Sub btnExtractRepo_Click(sender As Object, e As EventArgs) Handles btnExtractRepo.Click

        GetBOObjectList(False, txtExtractRepoDatabaseName.Text.ToString(), txtExtractRepoSQLServer.Text.ToString())

    End Sub

    Protected Sub GetBOObjectList(blnDisplay As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "")

        logger.[Debug]("GetBOObjectList: Begin sub")

        Me.NewBOSession()

        Dim strQuery As String
        Dim strSIID As String = Me.txtSIID.Text.ToString()
        If strSIID <> "" Then
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox') AND SI_ID = " + strSIID)
        Else
            strQuery = ("Select TOP 1000000 * FROM CI_INFOOBJECTS Where SI_KIND IN ('CrystalReport','Excel','Pdf','Webi','XL.XcelsiusApplication','Folder','FavoritesFolder','Publication','Inbox')")
        End If

        logger.[Debug]("GetBOObjectList: Begin to query CI_INFOOBJECTS")

        Dim objects As InfoObjects = Me.boInfoStore.Query(strQuery)

        logger.[Debug]("GetBOObjectList:      Finished query to CI_INFOOBJECTS and found records: " + objects.Count.ToString())

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
                logger.Error("ow noos! Error while querying and parsing Infostore for a list of objects", ex)
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If

        Me.LogoffBOSession()

    End Sub

    Protected Sub GetBOObjectProperties(blnDisplay As Boolean, Optional strDatabaseName As String = "", Optional strSQLServerName As String = "", Optional blnDeltas As Boolean = 1, Optional strSIID As String = "")

        logger.Info("GetBOObjectProperties:begin sub, operating in delta mode=" + blnDeltas.ToString)

        Dim strQuery As String
        Dim strSystemObjectQuery As String
        Dim strListOfColumnsToReturn As String
        Dim strListOfSystemObjectColumnsToReturn As String
        Dim myDataTableOfInfoObjects As DataTable

        strListOfColumnsToReturn = ""
        strListOfColumnsToReturn = strListOfColumnsToReturn + " SI_ID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_CUID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_CREATION_TIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_DESCRIPTION"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_ENDTIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_HAS_CHILDREN"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_INSTANCE"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_KIND"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_LAST_RUN_TIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_LAST_SUCCESSFUL_INSTANCE_ID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_NAME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_NEW_JOB_ID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_NEXTRUNTIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULE_STATUS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_OWNER"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_OWNERID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_PARENT_FOLDER"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_PARENTID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_RECURRING"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SIZE"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_STATUSINFO"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_UPDATE_TS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_CALENDAR_TEMPLATE_ID"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_TYPE"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_STARTTIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_ENDTIME"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_INTERVAL_HOURS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_INTERVAL_MINUTES"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_INTERVAL_MONTHS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_INTERVAL_NDAYS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_SCHEDULE_INTERVAL_NTHDAY"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_RUN_ON_TEMPLATE"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_SCHEDULEINFO.SI_DESTINATIONS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_PROCESSINFO.SI_HAS_PROMPTS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_PROCESSINFO.SI_WEBI_PROMPTS"
        strListOfColumnsToReturn = strListOfColumnsToReturn + ",SI_PROCESSINFO.SI_WEBI_PROMPTS"

        strListOfSystemObjectColumnsToReturn = ""
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + " SI_ID"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_CUID"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_ALIASES"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_CREATION_TIME"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_DATA"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_DESCRIPTION"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_EMAIL_ADDRESS"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_KIND"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_LASTLOGONTIME"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_NAME"
        strListOfSystemObjectColumnsToReturn = strListOfSystemObjectColumnsToReturn + ",SI_USERFULLNAME"

        logger.Info("GetBOObjectProperties: next step is to set SQL Connection to: " + strDatabaseName + " on " + strSQLServerName)
        SetSQLConnection(strDatabaseName, strSQLServerName)
        logger.Info("GetBOObjectProperties:      finished setting SQL Connection to: " + strDatabaseName + " on " + strSQLServerName)

        logger.Info("GetBOObjectProperties: next step is to call CreateObjectTablePropertyRepo()")
        CreateObjectTablePropertyRepo()
        CreateObjectTablePropertyRepoStage()
        logger.Info("GetBOObjectProperties:      finished call CreateObjectTablePropertyRepo()")

        strQuery = "Select TOP 100000" + strListOfColumnsToReturn + " FROM CI_INFOOBJECTS WHERE SI_KIND != 'LCMJob'"
        strSystemObjectQuery = "Select TOP 100000" + strListOfSystemObjectColumnsToReturn + " FROM CI_SYSTEMOBJECTS WHERE SI_KIND IN ('Calendar','User')"

        If strSIID <> "" Then
            strQuery = strQuery + " AND SI_ID = " + strSIID
        End If

        myDataTableOfInfoObjects = MakeDataTableForObjectPropertyList()

        If blnDeltas Then
            Dim strLastUpdateTimestamp = GetLoadObjectPropertiesDeltaTimestamp()
            strQuery = strQuery + " AND (SI_RECURRING=1 OR SI_UPDATE_TS >= '" + strLastUpdateTimestamp + "')"
            strSystemObjectQuery = strSystemObjectQuery + " AND SI_UPDATE_TS >= '" + strLastUpdateTimestamp + "'"

            SubGetBOObjectProperties(strDatabaseName, strSQLServerName, strQuery, myDataTableOfInfoObjects)
            SubGetBOObjectProperties(strDatabaseName, strSQLServerName, strSystemObjectQuery, myDataTableOfInfoObjects, "SystemProperties")

            'If we are providing an object id, we don't need to loop through dates
        ElseIf strSIID <> "" Then
            SubGetBOObjectProperties(strDatabaseName, strSQLServerName, strQuery, myDataTableOfInfoObjects)
        Else

            Dim dtToday As Date = Date.Now()
            Dim dtEarliestCreationDateToPull As Date = CDate("2014.01.01")
            Dim dtBegin As Date = dtEarliestCreationDateToPull
            Dim dtEnd As Date
            Dim intRange As Integer = 180
            Dim intDaysSpan As Integer = 180

            Do
                dtEnd = dtBegin.AddDays(intRange)

                intDaysSpan = (Date.Now - dtEnd).Days
                If intDaysSpan >= 360 Then
                    intRange = 240
                ElseIf intDaysSpan >= 240 Then
                    intRange = 120
                ElseIf intDaysSpan >= 120 Then
                    intRange = 30
                ElseIf intDaysSpan >= 30 Then
                    intRange = 10
                ElseIf intDaysSpan >= 10 Then
                    intRange = 5
                ElseIf intDaysSpan >= 5 Then
                    intRange = 1
                Else
                    intRange = 0
                End If

                Dim strQueryWithDateParm As String
                Dim strMyNewQueryToExecute As String
                strQueryWithDateParm = " AND SI_CREATION_TIME BETWEEN '" + ConvertDateToDateString(dtBegin) + "' AND '" + ConvertDateToDateString(dtEnd) + "'"
                strMyNewQueryToExecute = strQuery + strQueryWithDateParm

                SubGetBOObjectProperties(strDatabaseName, strSQLServerName, strMyNewQueryToExecute, myDataTableOfInfoObjects)

                'Load what we got and move on
                LoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects)
                myDataTableOfInfoObjects = MakeDataTableForObjectPropertyList()

                dtBegin = dtEnd.AddDays(1)

            Loop While dtBegin < dtToday

            'Full load of system objects query
            SubGetBOObjectProperties(strDatabaseName, strSQLServerName, strSystemObjectQuery, myDataTableOfInfoObjects, "SystemProperties")

        End If

        'Load what we got and move on
        LoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects)

        logger.Info("GetBOObjectProperties: finished sub")

        MergeObjectPropertyList(blnDeltas)

    End Sub

    Private Sub ParseInfoObjectProperties(myInfoObjects As InfoObjects, myDataTableOfInfoObjects As DataTable, Optional strClassName As String = "")

        Dim strClass As String
        Dim strId As String = ""
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

        If myInfoObjects.Count > 0 Then

            logger.Info("ParseInfoObjectProperties: found " + myInfoObjects.Count.ToString() + " InfoObjects to process.")

            For iLoop = 1 To myInfoObjects.Count

                Dim myInfoInfoObject As InfoObject
                Dim myInfoInfoProperties As Properties

                myInfoInfoObject = myInfoObjects(iLoop)
                myInfoInfoProperties = myInfoInfoObject.Properties()

                If myInfoInfoProperties IsNot Nothing Then

                    strId = myInfoInfoProperties.Item("SI_ID").Value.ToString()

                    If myInfoInfoProperties.Count > 0 Then
                        If strClassName = "" Then
                            strClass = "Properties"
                        Else
                            strClass = strClassName
                        End If

                        For iProp1 = 1 To myInfoInfoProperties.Count

                            strProperty1 = myInfoInfoProperties(iProp1).Name.ToString()
                            strPropertyValue1 = myInfoInfoProperties(iProp1).Value.ToString()
                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1, iProp1, strPropertyValue1)

                            If myInfoInfoProperties(iProp1).Properties.Count > 0 Then
                                For iProp2 = 1 To myInfoInfoProperties(iProp1).Properties.Count

                                    strProperty2 = myInfoInfoProperties(iProp1).Properties(iProp2).Name.ToString()
                                    strPropertyValue2 = myInfoInfoProperties(iProp1).Properties(iProp2).Value.ToString()
                                    BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)

                                    If myInfoInfoProperties(iProp1).Properties(iProp2).Properties.Count > 0 Then
                                        For iProp3 = 1 To myInfoInfoProperties(iProp1).Properties(iProp2).Properties.Count

                                            strProperty3 = myInfoInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Name.ToString()
                                            strPropertyValue3 = myInfoInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Value.ToString()
                                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)

                                            myInfoInfoProperties(iProp1).Properties(iProp2).Properties.Dispose()
                                        Next
                                    End If
                                    myInfoInfoProperties(iProp1).Properties.Dispose()
                                Next
                            End If
                            myInfoInfoProperties(iProp1).Dispose()
                        Next
                    End If
                End If


                Dim myProcesingInfoObject As ProcessingInfo
                Dim myProcesingInfoProperties As Properties

                myProcesingInfoObject = myInfoObjects(iLoop).ProcessingInfo()
                If myProcesingInfoObject IsNot Nothing Then

                    myProcesingInfoProperties = myProcesingInfoObject.Properties()

                    'If (myInfoObjects(iLoop).Kind = "Webi" Or myInfoObjects(iLoop).Kind = "Excel" Or myInfoObjects(iLoop).Kind = "Pdf" Or myInfoObjects(iLoop).Kind = "Txt") Then
                    If myProcesingInfoProperties.Count > 0 Then
                        strClass = "Processing"

                        For iProp1 = 1 To myProcesingInfoProperties.Count

                            strProperty1 = myProcesingInfoProperties(iProp1).Name.ToString()
                            strPropertyValue1 = myProcesingInfoProperties(iProp1).Value.ToString()
                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1, iProp1, strPropertyValue1)

                            If myProcesingInfoProperties(iProp1).Properties.Count > 0 Then
                                For iProp2 = 1 To myProcesingInfoProperties(iProp1).Properties.Count

                                    strProperty2 = myProcesingInfoProperties(iProp1).Properties(iProp2).Name.ToString()
                                    strPropertyValue2 = myProcesingInfoProperties(iProp1).Properties(iProp2).Value.ToString
                                    BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)

                                    If myProcesingInfoProperties(iProp1).Properties(iProp2).Properties.Count > 0 Then
                                        For iProp3 = 1 To myProcesingInfoProperties(iProp1).Properties(iProp2).Properties.Count

                                            strProperty3 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Name.ToString()
                                            strPropertyValue3 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Value.ToString
                                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)

                                            If myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count > 0 Then
                                                For iProp4 = 1 To myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties.Count

                                                    strProperty4 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Name.ToString()
                                                    strPropertyValue4 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Value.ToString
                                                    BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4, iProp4, strPropertyValue4)

                                                    If myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count > 0 Then
                                                        For iProp5 = 1 To myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties.Count

                                                            strProperty5 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Name.ToString()
                                                            strPropertyValue5 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Value.ToString
                                                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5, iProp5, strPropertyValue5)

                                                            If myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count > 0 Then
                                                                For iProp6 = 1 To myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties.Count

                                                                    strProperty6 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Name.ToString()
                                                                    strPropertyValue6 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Value.ToString()
                                                                    BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6, iProp6, strPropertyValue6)

                                                                    If myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties.Count > 0 Then
                                                                        For iProp7 = 1 To myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties.Count

                                                                            strProperty7 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties(iProp7).Name.ToString()
                                                                            strPropertyValue7 = myProcesingInfoProperties(iProp1).Properties(iProp2).Properties(iProp3).Properties(iProp4).Properties(iProp5).Properties(iProp6).Properties(iProp7).Value.ToString
                                                                            BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6 + strProperty7, iProp7, strPropertyValue7)

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


                    Dim mySchedulingInfoObject As SchedulingInfo
                    Dim mySchedulingInfoProperties As Properties

                    mySchedulingInfoObject = myInfoObjects(iLoop).SchedulingInfo()

                    If mySchedulingInfoObject IsNot Nothing Then

                        mySchedulingInfoProperties = mySchedulingInfoObject.Properties()

                        If mySchedulingInfoProperties.Count > 0 Then
                            For iProp1 = 1 To mySchedulingInfoProperties.Count
                                strClass = "Scheduling"

                                strProperty1 = mySchedulingInfoProperties(iProp1).Name.ToString()
                                strPropertyValue1 = mySchedulingInfoProperties(iProp1).Value.ToString()
                                BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1, iProp1, strPropertyValue1)

                                Dim mySchedulingInfoProperties1 = mySchedulingInfoProperties(iProp1).Properties

                                If mySchedulingInfoProperties1.Count > 0 Then
                                    For iProp2 = 1 To mySchedulingInfoProperties1.Count

                                        strProperty2 = mySchedulingInfoProperties1(iProp2).Name.ToString()
                                        strPropertyValue2 = mySchedulingInfoProperties1(iProp2).Value.ToString
                                        BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2, iProp2, strPropertyValue2)

                                        Dim mySchedulingInfoProperties2 = mySchedulingInfoProperties1(iProp2).Properties

                                        If mySchedulingInfoProperties2.Count > 0 Then
                                            For iProp3 = 1 To mySchedulingInfoProperties2.Count

                                                strProperty3 = mySchedulingInfoProperties2(iProp3).Name.ToString()
                                                strPropertyValue3 = mySchedulingInfoProperties2(iProp3).Value.ToString
                                                BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3, iProp3, strPropertyValue3)

                                                Dim mySchedulingInfoProperties3 = mySchedulingInfoProperties2(iProp3).Properties

                                                If mySchedulingInfoProperties3.Count > 0 Then
                                                    For iProp4 = 1 To mySchedulingInfoProperties3.Count

                                                        strProperty4 = mySchedulingInfoProperties3(iProp4).Name.ToString()
                                                        strPropertyValue4 = mySchedulingInfoProperties3(iProp4).Value.ToString
                                                        BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4, iProp4, strPropertyValue4)

                                                        Dim mySchedulingInfoProperties4 = mySchedulingInfoProperties3(iProp4).Properties

                                                        If mySchedulingInfoProperties4.Count > 0 Then
                                                            For iProp5 = 1 To mySchedulingInfoProperties4.Count

                                                                strProperty5 = mySchedulingInfoProperties4(iProp5).Name.ToString()
                                                                strPropertyValue5 = mySchedulingInfoProperties4(iProp5).Value.ToString
                                                                BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5, iProp5, strPropertyValue5)

                                                                Dim mySchedulingInfoProperties5 = mySchedulingInfoProperties4(iProp5).Properties

                                                                If mySchedulingInfoProperties5.Count > 0 Then
                                                                    For iProp6 = 1 To mySchedulingInfoProperties5.Count

                                                                        strProperty6 = mySchedulingInfoProperties5(iProp6).Name.ToString()
                                                                        strPropertyValue6 = mySchedulingInfoProperties5(iProp6).Value.ToString
                                                                        BuildLoadToDatabaseObjectPropertyStgString(myDataTableOfInfoObjects, strId, strClass, strProperty1 + "." + strProperty2 + "." + strProperty3 + "." + strProperty4 + "." + strProperty5 + "." + strProperty6, iProp6, strPropertyValue6)
                                                                    Next
                                                                End If
                                                                mySchedulingInfoProperties5.Dispose()
                                                                mySchedulingInfoProperties5 = Nothing

                                                            Next
                                                        End If
                                                        mySchedulingInfoProperties4.Dispose()
                                                        mySchedulingInfoProperties4 = Nothing

                                                    Next
                                                End If
                                                mySchedulingInfoProperties3.Dispose()
                                                mySchedulingInfoProperties3 = Nothing

                                            Next
                                        End If
                                        mySchedulingInfoProperties2.Dispose()
                                        mySchedulingInfoProperties2 = Nothing

                                    Next
                                End If
                                mySchedulingInfoProperties1.Dispose()
                                mySchedulingInfoProperties1 = Nothing

                            Next
                        End If
                        mySchedulingInfoProperties.Dispose()
                        mySchedulingInfoProperties = Nothing
                    End If

                    mySchedulingInfoObject.Dispose()
                    mySchedulingInfoObject = Nothing
                End If
            Next

        End If
    End Sub

    Private Sub SubGetBOObjectProperties(strDatabaseName As String, strSQLServerName As String, strQuery As String, myDataTableOfInfoObjects As DataTable, Optional strClassName As String = "")

        Dim myInfoObjects As InfoObjects
        Me.NewBOSession()

        logger.Info("SubGetBOObjectProperties: next step, execute CMC query: " + strQuery)
        Try
            myInfoObjects = Me.boInfoStore.Query(strQuery)
        Catch ex As Exception
            logger.Error("ow noos! Error in SubGetBOObjectProperties", ex)
            Exit Sub
        End Try
        logger.Info("SubGetBOObjectProperties:      finished executing CMC query")

        ParseInfoObjectProperties(myInfoObjects, myDataTableOfInfoObjects, strClassName)

        Me.LogoffBOSession()
        myInfoObjects.Dispose()
        myInfoObjects = Nothing

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

        cmdCMSServer = Me.cboCMSServer.SelectedItem.ToString()
        cmdCMSUser = Me.txtCMSUserName.Text.ToString()
        cmdCMSUserPassword = Me.txtCMSUserPassword.Text.ToString()
        cmdCMSAuthentication = Me.cboCMSAuthentication.SelectedItem.ToString()
        cmdTargetServer = txtLoadObjectPropertiesServer.Text.ToString()
        cmdTargetDB = txtLoadObjectPropertiesDB.Text.ToString()
        cmdDeltaProcessing = CBool(txtLoadObjectPropertyDelta.Text.ToString)

        'GetBOObjects(False, cmdTargetDB, cmdTargetServer, cmdDeltaProcessing)
        GetBOObjectProperties(False, cmdTargetDB, cmdTargetServer, cmdDeltaProcessing, Me.txtLoadObjectPropertiesSIID.Text.ToString())
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
            logger.Error("ow noos! Error in GetLoadObjectPropertiesDeltaTimestamp", ex)
            Exit Try
        End Try

        conSQLConn.Close()

        SetLoadObjectPropertiesDeltaTimestamp()

        GetLoadObjectPropertiesDeltaTimestamp = ConvertDateToDateTimeString(dtParmValue)

    End Function

    Private Sub SetLoadObjectPropertiesDeltaTimestamp()

        Dim strQuery As String = "UPDATE dbo.ETLParm SET ParmValue = '" + DateTime.UtcNow.AddMinutes(-1).ToString("yyyy-MM-dd HH:mm:ss") + "' WHERE ParmName = 'DeltaCMSObjectProperty'"
        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.Error("ow noos! Error in SetLoadObjectPropertiesDeltaTimestamp", ex)
            Exit Try
        End Try

        conSQLConn.Close()
    End Sub

    Private Function ConvertDateToDateTimeString(dtDateValue As Date) As String

        ConvertDateToDateTimeString = dtDateValue.ToString("yyyy.MM.dd.HH:mm:ss")

    End Function

    Private Function ConvertDateToDateString(dtDateValue As Date) As String

        ConvertDateToDateString = dtDateValue.ToString("yyyy.MM.dd")

    End Function

    Private Sub InitializeETLParmForObjectPropertyDelta()
        Dim strQuery As String = "IF NOT EXISTS (SELECT 1 FROM dbo.ETLParm WHERE BI_DBID = 0 AND DBID = 0 AND ParmName = 'DeltaCMSObjectProperty') INSERT INTO dbo.ETLParm (BI_DBID,DBID,ParmName,ParmValue) VALUES (0,0,'DeltaCMSObjectProperty','1901-01-01')"
        Dim cmdSQLCmd As New SqlCommand(strQuery, conSQLConn)

        Try
            conSQLConn.Open()

            cmdSQLCmd.ExecuteNonQuery()

        Catch ex As Exception
            logger.Error("ow noos! Error in SetLoadObjectPropertiesDeltaTimestamp", ex)
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

        Dim myBOProperties = New DataTable("MyBOProperties")

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

        MakeDataTableForObjectPropertyList = myBOProperties

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
        strQuery = strQuery & "   ) "
        strQuery = strQuery & " "
        strQuery = strQuery & "WHEN MATCHED "
        strQuery = strQuery & "     And "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "             NULLIF(src.PropertyValue, tgt.PropertyValue) Is Not NULL "
        strQuery = strQuery & "          Or NULLIF(tgt.PropertyValue, src.PropertyValue) Is Not NULL "
        strQuery = strQuery & "          Or NULLIF(src.PropertyValueInstanceNumber, tgt.PropertyValueInstanceNumber) Is Not NULL "
        strQuery = strQuery & "          Or NULLIF(tgt.PropertyValueInstanceNumber, src.PropertyValueInstanceNumber) Is Not NULL "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "THEN UPDATE SET "
        strQuery = strQuery & "          PropertyValue = src.PropertyValue "
        strQuery = strQuery & "         ,PropertyValueInstanceNumber = src.PropertyValueInstanceNumber "
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
        If Not blnDelta Then
            strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
            strQuery = strQuery & "THEN DELETE "
        End If
        strQuery = strQuery & ";"
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_NonRecurring_ListOfIDs tgt "
        strQuery = strQuery & "Using  "
        strQuery = strQuery & "     ( "
        strQuery = strQuery & "          SELECT "
        strQuery = strQuery & "               DISTINCT "
        strQuery = strQuery & "                ppt.SI_ID "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty ppt "
        strQuery = strQuery & " "
        strQuery = strQuery & "               INNER Join dbo.DimSAPBOObjectProperty ins "
        strQuery = strQuery & "               On ppt.SI_ID = ins.SI_ID "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                ppt.ClassName = 'Properties' "
        strQuery = strQuery & "               And ppt.PropertyName = 'SI_INSTANCE' "
        strQuery = strQuery & "               And ppt.PropertyValue = 'true' "
        strQuery = strQuery & " "
        strQuery = strQuery & "               And ins.ClassName = 'Properties' "
        strQuery = strQuery & "               And ins.PropertyName = 'SI_RECURRING' "
        strQuery = strQuery & "               And ins.PropertyValue = 'false' "
        strQuery = strQuery & "     ) src "
        strQuery = strQuery & "On ( tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)


        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_ListOfEmailBcc tgt "
        strQuery = strQuery & "Using "
        strQuery = strQuery & "    ( "
        strQuery = strQuery & "          SELECT DISTINCT "
        strQuery = strQuery & "               pptEmail.SI_ID "
        strQuery = strQuery & "              ,CAST(pptEmail.PropertyValue AS VARCHAR(200)) AS EmailAddressBcc "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty pptEmail "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                pptEmail.ClassName = 'Scheduling' "
        strQuery = strQuery & "               And pptEmail.PropertyName Like 'SI_DESTINATIONS.1.SI_DEST_SCHEDULEOPTIONS.SI_MAIL_BCC.%' "
        strQuery = strQuery & " And pptEmail.PropertyName Not Like '%.SI_TOTAL' "
        strQuery = strQuery & "    ) src "
        strQuery = strQuery & "On ( tgt.EmailAddressBcc = src.EmailAddressBcc And tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID, EmailAddressBcc ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, EmailAddressBcc "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_Recurring_ListOfIDs tgt "
        strQuery = strQuery & "Using "
        strQuery = strQuery & "    ( "
        strQuery = strQuery & "          SELECT "
        strQuery = strQuery & "               DISTINCT "
        strQuery = strQuery & "                ppt.SI_ID "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty ppt "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                ppt.ClassName = 'Properties' "
        strQuery = strQuery & "               And ( "
        strQuery = strQuery & "                    (ppt.PropertyName = 'SI_RECURRING' "
        strQuery = strQuery & "                     And ppt.PropertyValue = 'True') "
        strQuery = strQuery & "                     Or "
        strQuery = strQuery & "                    (ppt.PropertyName = 'SI_SCHEDULE_STATUS' "
        strQuery = strQuery & "                     And ppt.PropertyValue = '9') "
        strQuery = strQuery & "                    ) "
        strQuery = strQuery & "    ) src "
        strQuery = strQuery & "On ( tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_PromptName tgt "
        strQuery = strQuery & "Using  "
        strQuery = strQuery & "    ( "
        strQuery = strQuery & "          SELECT "
        strQuery = strQuery & "               promptName.SI_ID "
        strQuery = strQuery & "              ,SUBSTRING(promptName.PropertyName,  17, CHARINDEX('.',SUBSTRING(promptName.PropertyName,  17, 100))-1) AS PromptNum "
        strQuery = strQuery & "              ,promptName.PropertyValue AS PromptName "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty promptName "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                promptName.ClassName = 'Processing' "
        strQuery = strQuery & " And promptName.PropertyName Like 'SI_WEBI_PROMPTS%' "
        strQuery = strQuery & "               And promptName.PropertyName Like '%SI_NAME' "
        strQuery = strQuery & " And promptName.PropertyName Not Like '%SI_TOTAL' "
        strQuery = strQuery & "    ) src "
        strQuery = strQuery & "On ( tgt.PromptNum = src.PromptNum And tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN MATCHED And ( NULLIF(src.PromptName, tgt.PromptName) Is Not NULL "
        strQuery = strQuery & " Or NULLIF(tgt.PromptName, src.PromptName) Is Not NULL ) "
        strQuery = strQuery & "THEN UPDATE SET "
        strQuery = strQuery & "          PromptName = src.PromptName "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID, PromptNum, PromptName ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, PromptNum, PromptName "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)


        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_PromptValue tgt "
        strQuery = strQuery & "Using "
        strQuery = strQuery & "     ( "
        strQuery = strQuery & "          SELECT "
        strQuery = strQuery & "               DISTINCT "
        strQuery = strQuery & "                promptValues.SI_ID "
        strQuery = strQuery & "              ,CAST(SUBSTRING(promptValues.PropertyName,  17, CHARINDEX('.',SUBSTRING(promptValues.PropertyName,  17, 100))-1) AS INT) AS PromptNum "
        strQuery = strQuery & "              ,CAST(promptValues.PropertyValue AS VARCHAR(400)) AS PromptValue "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty promptValues "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                promptValues.ClassName = 'Processing' "
        strQuery = strQuery & "               And promptValues.ClassName = 'Processing' "
        strQuery = strQuery & " And promptValues.PropertyName Like '%SI_VALUES%' "
        strQuery = strQuery & "               And promptValues.PropertyName Not Like '%SI_TOTAL' "
        strQuery = strQuery & "     ) src "
        strQuery = strQuery & "On ( tgt.PromptNum = src.PromptNum "
        strQuery = strQuery & " And tgt.PromptValue = src.PromptValue "
        strQuery = strQuery & "     And tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "(SI_ID, PromptNum, PromptValue) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "( "
        strQuery = strQuery & "               SI_ID, PromptNum, PromptValue "
        strQuery = strQuery & ") "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_ListOfEmailTo tgt "
        strQuery = strQuery & "Using "
        strQuery = strQuery & "     ( "
        strQuery = strQuery & "          SELECT DISTINCT "
        strQuery = strQuery & "               pptEmail.SI_ID "
        strQuery = strQuery & "              ,CAST(pptEmail.PropertyValue AS VARCHAR(800)) AS EmailAddressTo "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty pptEmail "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                pptEmail.ClassName = 'Scheduling' "
        strQuery = strQuery & "               And pptEmail.PropertyName Like 'SI_DESTINATIONS.1.SI_DEST_SCHEDULEOPTIONS.SI_MAIL_ADDRESSES.%' "
        strQuery = strQuery & " And pptEmail.PropertyName Not Like '%.SI_TOTAL' "
        strQuery = strQuery & "     ) src "
        strQuery = strQuery & "On ( tgt.EmailAddressTo = src.EmailAddressTo And tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID, EmailAddressTo ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, EmailAddressTo "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_ListOfEmailCc tgt "
        strQuery = strQuery & "Using  "
        strQuery = strQuery & "    ( "
        strQuery = strQuery & "          SELECT DISTINCT "
        strQuery = strQuery & "               pptEmail.SI_ID "
        strQuery = strQuery & "              ,CAST(pptEmail.PropertyValue AS VARCHAR(200)) AS EmailAddressCc "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty pptEmail "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                pptEmail.ClassName = 'Scheduling' "
        strQuery = strQuery & "               And pptEmail.PropertyName Like 'SI_DESTINATIONS.1.SI_DEST_SCHEDULEOPTIONS.SI_MAIL_CC.%' "
        strQuery = strQuery & " And pptEmail.PropertyName Not Like '%.SI_TOTAL' "
        strQuery = strQuery & "    ) src "
        strQuery = strQuery & "On ( tgt.EmailAddressCc = src.EmailAddressCc And tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID, EmailAddressCc ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, EmailAddressCc "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)

        strQuery = ""
        strQuery = strQuery & "MERGE BI.DimBOObject_ScheduledReport_InstanceLastRunDateTime tgt "
        strQuery = strQuery & "Using  "
        strQuery = strQuery & "     ( "
        strQuery = strQuery & "          SELECT "
        strQuery = strQuery & "               pptScheduleEndTime.SI_ID "
        strQuery = strQuery & "              ,CAST(pptParentInstance.PropertyValue AS BIGINT) as Parent_SI_ID "
        strQuery = strQuery & "              ,CONVERT(DATETIME, pptScheduleEndTime.PropertyValue,121) AS InstanceLastRunDateTime "
        strQuery = strQuery & "          FROM "
        strQuery = strQuery & "                dbo.DimSAPBOObjectProperty pptScheduleEndTime "
        strQuery = strQuery & " "
        strQuery = strQuery & "               INNER Join dbo.DimSAPBOObjectProperty pptParentInstance "
        strQuery = strQuery & "               On pptScheduleEndTime.SI_ID = pptParentInstance.SI_ID "
        strQuery = strQuery & "               And pptParentInstance.ClassName = 'Properties' "
        strQuery = strQuery & "               And pptParentInstance.PropertyName = 'SI_NEW_JOB_ID' "
        strQuery = strQuery & "               And pptParentInstance.SI_ID <> pptParentInstance.PropertyValue "
        strQuery = strQuery & "          WHERE "
        strQuery = strQuery & "                pptScheduleEndTime.ClassName = 'Properties' "
        strQuery = strQuery & "                And pptScheduleEndTime.PropertyName = 'SI_ENDTIME' "
        strQuery = strQuery & "     ) src "
        strQuery = strQuery & "On ( tgt.SI_ID = src.SI_ID ) "
        strQuery = strQuery & " "
        strQuery = strQuery & "WHEN MATCHED "
        strQuery = strQuery & "                 And "
        strQuery = strQuery & "(NULLIF(src.Parent_SI_ID, tgt.Parent_SI_ID) Is Not NULL "
        strQuery = strQuery & "                        Or NULLIF(tgt.Parent_SI_ID, src.Parent_SI_ID) Is Not NULL "
        strQuery = strQuery & " Or NULLIF(src.InstanceLastRunDateTime, tgt.InstanceLastRunDateTime) Is Not NULL "
        strQuery = strQuery & "                        Or NULLIF(tgt.InstanceLastRunDateTime, src.InstanceLastRunDateTime) Is Not NULL ) "
        strQuery = strQuery & "THEN UPDATE SET "
        strQuery = strQuery & "          Parent_SI_ID = src.Parent_SI_ID "
        strQuery = strQuery & "         ,InstanceLastRunDateTime = src.InstanceLastRunDateTime "
        strQuery = strQuery & "WHEN Not MATCHED BY TARGET  "
        strQuery = strQuery & "THEN INSERT "
        strQuery = strQuery & "          ( SI_ID "
        strQuery = strQuery & "           ,Parent_SI_ID "
        strQuery = strQuery & "           ,InstanceLastRunDateTime ) "
        strQuery = strQuery & "     VALUES "
        strQuery = strQuery & "          ( "
        strQuery = strQuery & "               SI_ID, Parent_SI_ID, InstanceLastRunDateTime "
        strQuery = strQuery & "          ) "
        strQuery = strQuery & "WHEN Not MATCHED BY SOURCE "
        strQuery = strQuery & "THEN DELETE "
        strQuery = strQuery & "; "
        ExecuteQuery(strQuery)


    End Sub

    Private Sub btnDeleteReportsFromFile_Click(sender As Object, e As EventArgs) Handles btnDeleteReportsFromFile.Click

        Dim txtOutput As String()
        Dim siIdsFilePath As String = Me.txtFileNameLocationForDeletes.Text

        ' Read SI_IDs from the file
        Dim siIdsToDelete As String()
        Try
            siIdsToDelete = File.ReadAllLines(siIdsFilePath).
                Where(Function(line) Not String.IsNullOrWhiteSpace(line)).
                Select(Function(line) line.Trim()).
                ToArray()
        Catch ex As Exception
            Me.rtbOutput.AppendText($"Error reading file: {ex.Message}")
            Return
        End Try

        If siIdsToDelete.Length = 0 Then
            Me.rtbOutput.AppendText("No valid SI_IDs found In the file.")
            Return
        End If

        Try
            ' Connect to the CMS
            Me.NewBOSession()

            ' Build the query to fetch Webi reports with the specified SI_IDs
            Dim strQuery As String = $"SELECT SI_ID FROM CI_INFOOBJECTS WHERE SI_KIND='Webi' AND ({String.Join(" OR ", siIdsToDelete.Select(Function(id) $"SI_ID='{id}'"))})"
            Me.rtbOutput.AppendText(String.Concat(strQuery, ChrW(13) & ChrW(10), ChrW(13) & ChrW(10)))
            Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

            ' Delete the reports
            Dim deletedCount As Integer = 0

            If infoObjects.Count > 0 Then

                Dim enumerator As IEnumerator

                enumerator = infoObjects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim objCurrentObject As InfoObject = DirectCast(enumerator.Current, InfoObject)

                    txtOutput = {"Marking report for deletion with SI_ID: ", objCurrentObject.ID, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))
                    deletedCount += 1
                    objCurrentObject.DeleteNow()
                Loop

                txtOutput = {"Deletion process completed. Number of deletes: ", deletedCount, ChrW(13) & ChrW(10)}
                Me.rtbOutput.AppendText(String.Concat(txtOutput))

                Me.boInfoStore.Commit(infoObjects)

            Else
                txtOutput = {"No matching reports found for deletion.", ChrW(13) & ChrW(10)}
                Me.rtbOutput.AppendText(String.Concat(txtOutput))
            End If

        Catch ex As Exception
            txtOutput = {$"An error occurred: {ex.Message}", ChrW(13) & ChrW(10)}
            Me.rtbOutput.AppendText(String.Concat(txtOutput))
        Finally
            If Me.boEnterpriseSession IsNot Nothing Then
                Try
                    Me.LogoffBOSession()
                Catch ex As Exception
                    txtOutput = {$"Error during logoff: {ex.Message}", ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(txtOutput))
                End Try
            End If
        End Try

    End Sub

    Private Sub RescheduleInstancesFromCSV(csvFilePath As String)
        logger.Info($"Starting to process CSV file: {csvFilePath}")

        ' Read the CSV file
        Dim lines() As String
        Try
            lines = File.ReadAllLines(csvFilePath)
        Catch ex As IOException
            Dim errorMsg As String = $"Error reading CSV file: {ex.Message}"
            Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
            logger.Error(errorMsg)
            Return
        End Try

        ' Loop through each line in the CSV file (skipping the header)
        For i As Integer = 1 To lines.Length - 1
            Dim line As String = lines(i)
            Dim columns() As String = line.Split(","c)

            ' Ensure there are enough columns
            If columns.Length < 3 Then
                Dim errorMsg As String = $"Malformed CSV line: {line}"
                Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
                logger.Warn(errorMsg)
                Continue For
            End If

            ' Extract data from each column
            Dim username As String = columns(0).Trim()
            Dim siId As String = columns(1).Trim()
            Dim datetimePart As String = columns(2).Trim()

            ' Parse the combined date and time into a DateTime object
            Dim newStartDateTime As DateTime
            Dim formats() As String = {
                "yyyy-MM-dd h:mm tt", "MM/dd/yyyy HH:mm", "MMMM dd, yyyy h:mm tt",
                "MM/dd/yyyy h:mm tt", "yyyy-MM-dd HH:mm", "yyyy-MM-ddTHH:mm:ss",
                "dd/MM/yyyy HH:mm", "M/d/yyyy h:mm tt", "M/d/yyyy H:mm",
                "MM/dd/yyyy h:mm tt", "yyyy-MM-ddTHH:mm:ss", "h:mm tt dd/MM/yyyy",
                "h:mm tt yyyy-MM-dd", "HH:mm yyyy-MM-dd", "dd/MM/yyyy h:mm tt",
                "yyyy-MM-dd hh:mmtt", "MM/dd/yyyy hh:mmtt"  ' Added formats without space between time and AM/PM
            }

            ' Try to parse using different formats
            If Not DateTime.TryParseExact(datetimePart, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, newStartDateTime) Then
                ' Fallback to regular parsing if exact parsing fails
                If Not DateTime.TryParse(datetimePart, newStartDateTime) Then
                    Dim errorMsg As String = $"Error parsing date and time for SI_ID {siId} on line: {line}"
                    Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
                    logger.Error(errorMsg)
                    Continue For
                End If
            End If

            ' Log the start of the instance rescheduling process
            logger.Info($"Starting rescheduling for SI_ID: {siId}, Username: {username}, New Start Time: {newStartDateTime}")

            ' Call the method to reschedule the instance with the parsed DateTime
            RescheduleInstance(username, siId, newStartDateTime)

            logger.Info($"Finished rescheduling for SI_ID: {siId}")
        Next

        logger.Info($"Finished processing CSV file: {csvFilePath}")
    End Sub

    Private Sub RescheduleInstance(username As String, siId As String, newStartDateTime As DateTime)
        Dim cmsName As String = Me.cboCMSServer.SelectedItem.ToString()
        Dim password As String = Me.txtCMSUserPassword.Text.ToString()
        Dim authType As String = "secEnterprise"

        Try
            ' Logon to the CMS
            Dim sessionMgr As New SessionMgr()
            Dim enterpriseSession As EnterpriseSession = sessionMgr.Logon(username, password, cmsName, authType)

            ' Access the InfoStore service
            Dim boInfoStore As InfoStore = CType(enterpriseSession.GetService("InfoStore"), InfoStore)

            ' Query for the specific instance using SI_ID
            Dim strQuery As String = "SELECT * FROM CI_INFOOBJECTS WHERE SI_ID = " & siId
            Dim infoObjects As InfoObjects = boInfoStore.Query(strQuery)

            ' Check if the instance exists
            If infoObjects.Count > 0 Then
                Dim enumerator As IEnumerator = infoObjects.GetEnumerator()
                Try
                    Do While enumerator.MoveNext()
                        ' Retrieve the current InfoObject
                        Dim objCurrentObject As InfoObject = DirectCast(enumerator.Current, InfoObject)

                        ' Update the scheduling info with the new start time
                        objCurrentObject.SchedulingInfo.Properties.Item("SI_STARTTIME").Value = newStartDateTime
                        objCurrentObject.SchedulingInfo.Properties.Item("SI_RECURRING").Value = 1

                        ' Optionally update other properties, such as the update timestamp
                        objCurrentObject.Properties.Item("SI_UPDATE_TS").Value = Date.Now
                    Loop
                Finally
                    ' Dispose of the enumerator properly
                    If TypeOf enumerator Is IDisposable Then
                        DirectCast(enumerator, IDisposable).Dispose()
                    End If
                End Try

                ' Commit the changes to the InfoStore
                boInfoStore.Commit(infoObjects)

                ' Write success message to the RichTextBox and log4net
                Dim successMsg As String = $"Instance with SI_ID {siId} has been rescheduled to {newStartDateTime}."
                Me.rtbOutput.AppendText(successMsg & ChrW(13) & ChrW(10))
                logger.Info(successMsg)

            Else
                ' Write "not found" message to the RichTextBox and log4net
                Dim notFoundMsg As String = $"No instance found with SI_ID {siId}."
                Me.rtbOutput.AppendText(notFoundMsg & ChrW(13) & ChrW(10))
                logger.Warn(notFoundMsg)
            End If

        Catch ex As UnauthorizedAccessException
            Dim errorMsg As String = $"Unauthorized access error for SI_ID {siId}: {ex.Message}"
            Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
            logger.Error(errorMsg)
        Catch ex As SecurityException
            Dim errorMsg As String = $"Security error for SI_ID {siId}: {ex.Message}"
            Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
            logger.Error(errorMsg)
        Catch ex As Exception
            ' Write generic error message to the RichTextBox and log4net
            Dim errorMsg As String = $"Error processing SI_ID {siId}: {ex.Message}"
            Me.rtbOutput.AppendText(errorMsg & ChrW(13) & ChrW(10))
            logger.Error(errorMsg)
        End Try
    End Sub

    Private Sub btnRescheduleInstance_Click(sender As Object, e As EventArgs) Handles btnRescheduleInstance.Click
        RescheduleInstancesFromCSV(Me.txtRescheduleFileList.Text)
    End Sub



    'Sub CleanRepoOrphans()

    '    Dim Run_In_Test_Mode As Boolean = True
    '    Dim pathToInputFRS As String = "\\msp-fs\bi-filestore\Input"
    '    Dim pathToOutputFRS As String = "\\msp-fs\bi-filestore\Output"
    '    Dim username As String = "Administrator"
    '    Dim password As String = "Password1"
    '    Dim cmsname As String = "localhost"
    '    Dim authType As String = "secEnterprise"
    '    Me.NewBOSession()

    '    CheckAndDelete(Run_In_Test_Mode, pathToInputFRS, "Input")
    '    CheckAndDelete(Run_In_Test_Mode, pathToOutputFRS, "Output")

    '    Me.LogoffBOSession()

    'End Sub

    'Private Sub CheckAndDelete(doDelete As Boolean, currentFolderPath As String, currentFolderName As String)
    '    ' Loop through all the files / subdirectories in the current directory

    '    Dim isOrphan As Boolean = False
    '    Dim checkFiles As Boolean = False
    '    Try
    '        Dim currentDir As New DirectoryInfo(currentFolderPath)
    '        Dim myListOfFiles As FileInfo() = currentDir.GetFiles("*", SearchOption.AllDirectories)
    '        Dim myFile As FileInfo

    '        For Each myFile In myListOfFiles
    '            ' Have we already checked the files for this directory.

    '            ' Since there are files - that means this is an infoobject and we need to check it.
    '            ' The current directory name is the SI_ID of the infoobject, so query for that.
    '            Dim strQuery As String
    '            strQuery = "Select * from CI_INFOOBJECTS, CI_APPOBJECTS, CI_SYSTEMOBJECTS where SI_ID = " + currentFolderName
    '            Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

    '            ' Did we get a result
    '            If (infoObjects.Count = 0) Then
    '                    isOrphan = True
    '                End If

    '                ' If we verified that this is an orphan - delete the file
    '                If (isOrphan = True) Then
    '                writeToLog(("Deleting file " _
    '                                + (File.getCanonicalPath + (" from Orphaned Infoobject with SI_ID " + currentFolderName))), fileToLogTo)
    '                ' Are we in demo mode
    '                If (doDelete = False) Then
    '                    File.Delete()
    '                End If

    '            End If

    '            End If

    '        Next
    '    Catch ex As Exception

    '    End Try

    'End Sub

End Class
