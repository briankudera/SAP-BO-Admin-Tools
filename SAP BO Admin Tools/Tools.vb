Imports CrystalDecisions.Enterprise

Public Class Tools

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

    Private Sub NewSession()

        Dim mgr As New SessionMgr
        Dim userName As String = Me.txtCMSUserName.Text.ToString
        Dim password As String = Me.txtCMSUserPassword.Text.ToString
        Dim cMSName As String = Me.cboCMSServer.SelectedItem.ToString
        Dim authentication As String = Me.cboCMSAuthentication.SelectedItem.ToString
        Me.boEnterpriseSession = mgr.Logon(userName, password, cMSName, authentication)
        Me.boInfoStore = New InfoStore(Me.boEnterpriseSession.GetService("InfoStore"))

    End Sub

    Private Sub LogoffSession()

        Me.boEnterpriseSession.Logoff()
        Me.boEnterpriseSession.Dispose()
        Me.rtbOutput.AppendText(ChrW(13) & ChrW(10))
        Me.rtbOutput.AppendText("Closing session." & ChrW(13) & ChrW(10))
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

        Me.NewSession()

        Dim intOwnerIdOld As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Text)
        Dim intOwnerIdNew As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Text)
        Dim strQuery As String = ("select si_id, si_name, si_ownerid, si_kind from ci_infoobjects where si_ownerid = " & intOwnerIdOld.ToString & " and si_kind not in ('FavoritesFolder','PersonalCategory','Inbox')")
        Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)

        If (infoObjects.Count > 0) Then
            Dim enumerator As IEnumerator
            Try
                enumerator = infoObjects.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                    Dim str2 As String = current.Properties.Item("SI_NAME").Value.ToString
                    Dim str3 As String = current.Properties.Item("SI_KIND").Value.ToString
                    current.Properties.Item("SI_OWNERID").Value = intOwnerIdNew
                    Dim textArray1 As String() = New String() {"Object updated: (", str3, ") ", str2, ChrW(13) & ChrW(10)}
                    Me.rtbOutput.AppendText(String.Concat(textArray1))
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Me.boInfoStore.Commit(infoObjects)
        End If

        Me.LogoffSession()

    End Sub

    Private Sub btnReplaceOwnerOnSingeDoc_Click(sender As Object, e As EventArgs) Handles btnReplaceOwnerOnSingeDoc.Click
        Dim intDocId As Integer
        Me.NewSession()

        If Not Integer.TryParse(Me.txtReplaceOwnerOnSingleDocDocId.Text, intDocId) Then
            intDocId = -1
        End If
        If (intDocId > 0) Then
            Dim intOwnerIdNew As Integer = Me.GetIdForUser(Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Text)
            If (intOwnerIdNew > 0) Then

                Dim strQuery As String = ("select si_id, si_name, si_ownerid, si_kind from ci_infoobjects where si_id = " & intDocId.ToString)
                Dim infoObjects As InfoObjects = Me.boInfoStore.Query(strQuery)
                If (infoObjects.Count > 0) Then
                    Dim enumerator As IEnumerator
                    Try
                        enumerator = infoObjects.GetEnumerator
                        Do While enumerator.MoveNext
                            Dim current As InfoObject = DirectCast(enumerator.Current, InfoObject)
                            Dim str2 As String = current.Properties.Item("SI_NAME").Value.ToString
                            Dim str3 As String = current.Properties.Item("SI_KIND").Value.ToString
                            current.Properties.Item("SI_OWNERID").Value = intOwnerIdNew
                            Dim textArray1 As String() = New String() {"Object updated: (", str3, ") ", str2, ChrW(13) & ChrW(10)}
                            Me.rtbOutput.AppendText(String.Concat(textArray1))
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
        Me.LogoffSession()

    End Sub

    Private Sub btnListObjectsByOwner_Click(sender As Object, e As EventArgs) Handles btnListObjectsByOwner.Click

        Dim strUserOwnerName As String = Me.txtListObjectsByOwnerOwnerName.Text.ToString
        Me.rtbOutput.AppendText("Begin User ID lookup" & ChrW(13) & ChrW(10))
        Me.NewSession()

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
        Me.LogoffSession()

    End Sub

    Private Sub Tools_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If cboCMSServer.Items.Count > 0 Then
            cboCMSServer.SelectedIndex = 0    ' The first item has index 0 '
        End If
        If cboCMSAuthentication.Items.Count > 0 Then
            cboCMSAuthentication.SelectedIndex = 0    ' The first item has index 0 '
        End If
    End Sub
End Class
