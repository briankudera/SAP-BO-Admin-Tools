<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
<Global.System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726")>
Partial Class frmTools
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents PasswordLabel As System.Windows.Forms.Label
    Friend WithEvents txtCMSUserName As System.Windows.Forms.TextBox

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTools))
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.txtCMSUserName = New System.Windows.Forms.TextBox()
        Me.cboCMSServer = New System.Windows.Forms.ComboBox()
        Me.txtCMSUserPassword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboCMSAuthentication = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tabFunctions = New System.Windows.Forms.TabControl()
        Me.tabListObjectsByOwner = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnListObjectsByOwner = New System.Windows.Forms.Button()
        Me.txtListObjectsByOwnerOwnerName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tabReplaceOwnerOnAllObjects = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnRepalceOwnerOnAllObjects = New System.Windows.Forms.Button()
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tabReplaceOwnerOnSingleDoc = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnReplaceOwnerOnSingeDoc = New System.Windows.Forms.Button()
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtReplaceOwnerOnSingleDocDocId = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tabGetListOfUsers = New System.Windows.Forms.TabPage()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnGetListOfUsers = New System.Windows.Forms.Button()
        Me.tabLoadListOfUsersToDB = New System.Windows.Forms.TabPage()
        Me.txtSIUserId = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtLoadListOfUsersToDBDatabaseName = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnLoadListOfUsersToDB = New System.Windows.Forms.Button()
        Me.txtLoadListOfUsersToDBSQLServerName = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tabExtractRepo = New System.Windows.Forms.TabPage()
        Me.txtSIID = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtExtractRepoDatabaseName = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtExtractRepoSQLServer = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.btnExtractRepo = New System.Windows.Forms.Button()
        Me.tabAddDBCredentials = New System.Windows.Forms.TabPage()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.txtAddDBCredentialsPW = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btnAddDBCredentialsForGroup = New System.Windows.Forms.Button()
        Me.txtGroupNameToProcess = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.tabResetAdminPW = New System.Windows.Forms.TabPage()
        Me.btnResetAdminPW = New System.Windows.Forms.Button()
        Me.tabUpdateOwnersOnDocsFromCSV = New System.Windows.Forms.TabPage()
        Me.chkOwnerIsSameOnAllDocs = New System.Windows.Forms.CheckBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtCSVFilePath = New System.Windows.Forms.TextBox()
        Me.btnUpdateOwnersOnDocsFromCSV = New System.Windows.Forms.Button()
        Me.tabLoadObjectPropertiesToDB = New System.Windows.Forms.TabPage()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtLoadObjectPropertyDelta = New System.Windows.Forms.TextBox()
        Me.txtDeltaTsp = New System.Windows.Forms.TextBox()
        Me.btnCheckDeltaTsp = New System.Windows.Forms.Button()
        Me.txtLoadObjectPropertiesSIID = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtLoadObjectPropertiesDB = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtLoadObjectPropertiesServer = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.btnLoadObjectPropertiesToDB = New System.Windows.Forms.Button()
        Me.rtbOutput = New System.Windows.Forms.RichTextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tabDeleteReportsBySIID = New System.Windows.Forms.TabPage()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtFileNameLocationForDeletes = New System.Windows.Forms.TextBox()
        Me.btnDeleteReportsFromFile = New System.Windows.Forms.Button()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabFunctions.SuspendLayout()
        Me.tabListObjectsByOwner.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.tabReplaceOwnerOnAllObjects.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.tabReplaceOwnerOnSingleDoc.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.tabGetListOfUsers.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.tabLoadListOfUsersToDB.SuspendLayout()
        Me.tabExtractRepo.SuspendLayout()
        Me.tabAddDBCredentials.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.tabResetAdminPW.SuspendLayout()
        Me.tabUpdateOwnersOnDocsFromCSV.SuspendLayout()
        Me.tabLoadObjectPropertiesToDB.SuspendLayout()
        Me.tabDeleteReportsBySIID.SuspendLayout()
        Me.SuspendLayout()
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), System.Drawing.Image)
        Me.LogoPictureBox.Location = New System.Drawing.Point(0, 0)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.LogoPictureBox.Size = New System.Drawing.Size(165, 142)
        Me.LogoPictureBox.TabIndex = 0
        Me.LogoPictureBox.TabStop = False
        '
        'UsernameLabel
        '
        Me.UsernameLabel.Location = New System.Drawing.Point(185, 26)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(220, 23)
        Me.UsernameLabel.TabIndex = 0
        Me.UsernameLabel.Text = "&CMS Server:"
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PasswordLabel
        '
        Me.PasswordLabel.Location = New System.Drawing.Point(185, 52)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(77, 23)
        Me.PasswordLabel.TabIndex = 2
        Me.PasswordLabel.Text = "&User name:"
        Me.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCMSUserName
        '
        Me.txtCMSUserName.Location = New System.Drawing.Point(264, 54)
        Me.txtCMSUserName.Name = "txtCMSUserName"
        Me.txtCMSUserName.Size = New System.Drawing.Size(220, 20)
        Me.txtCMSUserName.TabIndex = 2
        Me.txtCMSUserName.Text = "Administrator"
        '
        'cboCMSServer
        '
        Me.cboCMSServer.FormattingEnabled = True
        Me.cboCMSServer.Items.AddRange(New Object() {"MSP-BILUM01:6400", "TST-BICMS:6400"})
        Me.cboCMSServer.Location = New System.Drawing.Point(264, 28)
        Me.cboCMSServer.Name = "cboCMSServer"
        Me.cboCMSServer.Size = New System.Drawing.Size(220, 21)
        Me.cboCMSServer.TabIndex = 1
        '
        'txtCMSUserPassword
        '
        Me.txtCMSUserPassword.Location = New System.Drawing.Point(264, 80)
        Me.txtCMSUserPassword.Name = "txtCMSUserPassword"
        Me.txtCMSUserPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCMSUserPassword.Size = New System.Drawing.Size(220, 20)
        Me.txtCMSUserPassword.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(185, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "&Password:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboCMSAuthentication
        '
        Me.cboCMSAuthentication.FormattingEnabled = True
        Me.cboCMSAuthentication.Items.AddRange(New Object() {"secEnterprise", "secWinAD"})
        Me.cboCMSAuthentication.Location = New System.Drawing.Point(264, 107)
        Me.cboCMSAuthentication.Name = "cboCMSAuthentication"
        Me.cboCMSAuthentication.Size = New System.Drawing.Size(220, 21)
        Me.cboCMSAuthentication.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(185, 105)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(220, 23)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "&Authentication:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(172, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(323, 133)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Authenticate"
        '
        'tabFunctions
        '
        Me.tabFunctions.Controls.Add(Me.tabListObjectsByOwner)
        Me.tabFunctions.Controls.Add(Me.tabReplaceOwnerOnAllObjects)
        Me.tabFunctions.Controls.Add(Me.tabReplaceOwnerOnSingleDoc)
        Me.tabFunctions.Controls.Add(Me.tabGetListOfUsers)
        Me.tabFunctions.Controls.Add(Me.tabLoadListOfUsersToDB)
        Me.tabFunctions.Controls.Add(Me.tabExtractRepo)
        Me.tabFunctions.Controls.Add(Me.tabAddDBCredentials)
        Me.tabFunctions.Controls.Add(Me.tabResetAdminPW)
        Me.tabFunctions.Controls.Add(Me.tabUpdateOwnersOnDocsFromCSV)
        Me.tabFunctions.Controls.Add(Me.tabLoadObjectPropertiesToDB)
        Me.tabFunctions.Controls.Add(Me.tabDeleteReportsBySIID)
        Me.tabFunctions.Location = New System.Drawing.Point(8, 149)
        Me.tabFunctions.Name = "tabFunctions"
        Me.tabFunctions.SelectedIndex = 0
        Me.tabFunctions.Size = New System.Drawing.Size(558, 118)
        Me.tabFunctions.TabIndex = 10
        '
        'tabListObjectsByOwner
        '
        Me.tabListObjectsByOwner.Controls.Add(Me.GroupBox2)
        Me.tabListObjectsByOwner.Location = New System.Drawing.Point(4, 22)
        Me.tabListObjectsByOwner.Name = "tabListObjectsByOwner"
        Me.tabListObjectsByOwner.Padding = New System.Windows.Forms.Padding(3)
        Me.tabListObjectsByOwner.Size = New System.Drawing.Size(550, 92)
        Me.tabListObjectsByOwner.TabIndex = 0
        Me.tabListObjectsByOwner.Text = "List Objects by Owner"
        Me.tabListObjectsByOwner.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnListObjectsByOwner)
        Me.GroupBox2.Controls.Add(Me.txtListObjectsByOwnerOwnerName)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 7)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(537, 79)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Enter Info"
        '
        'btnListObjectsByOwner
        '
        Me.btnListObjectsByOwner.Location = New System.Drawing.Point(305, 17)
        Me.btnListObjectsByOwner.Name = "btnListObjectsByOwner"
        Me.btnListObjectsByOwner.Size = New System.Drawing.Size(97, 23)
        Me.btnListObjectsByOwner.TabIndex = 14
        Me.btnListObjectsByOwner.Text = "List Objects"
        Me.btnListObjectsByOwner.UseVisualStyleBackColor = True
        '
        'txtListObjectsByOwnerOwnerName
        '
        Me.txtListObjectsByOwnerOwnerName.Location = New System.Drawing.Point(79, 19)
        Me.txtListObjectsByOwnerOwnerName.Name = "txtListObjectsByOwnerOwnerName"
        Me.txtListObjectsByOwnerOwnerName.Size = New System.Drawing.Size(220, 20)
        Me.txtListObjectsByOwnerOwnerName.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 17)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(105, 23)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "&Owner name:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tabReplaceOwnerOnAllObjects
        '
        Me.tabReplaceOwnerOnAllObjects.Controls.Add(Me.GroupBox3)
        Me.tabReplaceOwnerOnAllObjects.Location = New System.Drawing.Point(4, 22)
        Me.tabReplaceOwnerOnAllObjects.Name = "tabReplaceOwnerOnAllObjects"
        Me.tabReplaceOwnerOnAllObjects.Padding = New System.Windows.Forms.Padding(3)
        Me.tabReplaceOwnerOnAllObjects.Size = New System.Drawing.Size(550, 92)
        Me.tabReplaceOwnerOnAllObjects.TabIndex = 1
        Me.tabReplaceOwnerOnAllObjects.Text = "Replace Owner on all Objects"
        Me.tabReplaceOwnerOnAllObjects.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnRepalceOwnerOnAllObjects)
        Me.GroupBox3.Controls.Add(Me.txtReplaceOwnerOnAllObjectsOwnerNameNew)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.txtReplaceOwnerOnAllObjectsOwnerNameOld)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Location = New System.Drawing.Point(7, 7)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(537, 79)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enter Info"
        '
        'btnRepalceOwnerOnAllObjects
        '
        Me.btnRepalceOwnerOnAllObjects.Location = New System.Drawing.Point(347, 19)
        Me.btnRepalceOwnerOnAllObjects.Name = "btnRepalceOwnerOnAllObjects"
        Me.btnRepalceOwnerOnAllObjects.Size = New System.Drawing.Size(97, 46)
        Me.btnRepalceOwnerOnAllObjects.TabIndex = 11
        Me.btnRepalceOwnerOnAllObjects.Text = "Replace Owner on All Objects"
        Me.btnRepalceOwnerOnAllObjects.UseVisualStyleBackColor = True
        '
        'txtReplaceOwnerOnAllObjectsOwnerNameNew
        '
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Location = New System.Drawing.Point(108, 45)
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Name = "txtReplaceOwnerOnAllObjectsOwnerNameNew"
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.Size = New System.Drawing.Size(220, 20)
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(105, 23)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "O&wner name (new):"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtReplaceOwnerOnAllObjectsOwnerNameOld
        '
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Location = New System.Drawing.Point(108, 19)
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Name = "txtReplaceOwnerOnAllObjectsOwnerNameOld"
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.Size = New System.Drawing.Size(220, 20)
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 23)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "&Owner name (old):"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tabReplaceOwnerOnSingleDoc
        '
        Me.tabReplaceOwnerOnSingleDoc.Controls.Add(Me.GroupBox4)
        Me.tabReplaceOwnerOnSingleDoc.Location = New System.Drawing.Point(4, 22)
        Me.tabReplaceOwnerOnSingleDoc.Name = "tabReplaceOwnerOnSingleDoc"
        Me.tabReplaceOwnerOnSingleDoc.Size = New System.Drawing.Size(550, 92)
        Me.tabReplaceOwnerOnSingleDoc.TabIndex = 2
        Me.tabReplaceOwnerOnSingleDoc.Text = "Replace Owner On Single Doc"
        Me.tabReplaceOwnerOnSingleDoc.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnReplaceOwnerOnSingeDoc)
        Me.GroupBox4.Controls.Add(Me.txtReplaceOwnerOnSingleDocOwnerNameNew)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.txtReplaceOwnerOnSingleDocDocId)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Location = New System.Drawing.Point(7, 7)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(537, 82)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Enter Info"
        '
        'btnReplaceOwnerOnSingeDoc
        '
        Me.btnReplaceOwnerOnSingeDoc.Location = New System.Drawing.Point(347, 19)
        Me.btnReplaceOwnerOnSingeDoc.Name = "btnReplaceOwnerOnSingeDoc"
        Me.btnReplaceOwnerOnSingeDoc.Size = New System.Drawing.Size(97, 46)
        Me.btnReplaceOwnerOnSingeDoc.TabIndex = 16
        Me.btnReplaceOwnerOnSingeDoc.Text = "Replace Owner on Single Doc"
        Me.btnReplaceOwnerOnSingeDoc.UseVisualStyleBackColor = True
        '
        'txtReplaceOwnerOnSingleDocOwnerNameNew
        '
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Location = New System.Drawing.Point(108, 45)
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Name = "txtReplaceOwnerOnSingleDocOwnerNameNew"
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew.Size = New System.Drawing.Size(220, 20)
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 23)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "O&wner name (new):"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtReplaceOwnerOnSingleDocDocId
        '
        Me.txtReplaceOwnerOnSingleDocDocId.Location = New System.Drawing.Point(108, 19)
        Me.txtReplaceOwnerOnSingleDocDocId.Name = "txtReplaceOwnerOnSingleDocDocId"
        Me.txtReplaceOwnerOnSingleDocDocId.Size = New System.Drawing.Size(220, 20)
        Me.txtReplaceOwnerOnSingleDocDocId.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 17)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(105, 23)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "&Document ID:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tabGetListOfUsers
        '
        Me.tabGetListOfUsers.Controls.Add(Me.GroupBox5)
        Me.tabGetListOfUsers.Location = New System.Drawing.Point(4, 22)
        Me.tabGetListOfUsers.Name = "tabGetListOfUsers"
        Me.tabGetListOfUsers.Size = New System.Drawing.Size(550, 92)
        Me.tabGetListOfUsers.TabIndex = 3
        Me.tabGetListOfUsers.Text = "Get List of Users"
        Me.tabGetListOfUsers.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnGetListOfUsers)
        Me.GroupBox5.Location = New System.Drawing.Point(7, 7)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(537, 79)
        Me.GroupBox5.TabIndex = 1
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Enter Info"
        '
        'btnGetListOfUsers
        '
        Me.btnGetListOfUsers.Location = New System.Drawing.Point(6, 19)
        Me.btnGetListOfUsers.Name = "btnGetListOfUsers"
        Me.btnGetListOfUsers.Size = New System.Drawing.Size(97, 23)
        Me.btnGetListOfUsers.TabIndex = 14
        Me.btnGetListOfUsers.Text = "Get List of Users"
        Me.btnGetListOfUsers.UseVisualStyleBackColor = True
        '
        'tabLoadListOfUsersToDB
        '
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.txtSIUserId)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.Label16)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.txtLoadListOfUsersToDBDatabaseName)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.Label9)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.btnLoadListOfUsersToDB)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.txtLoadListOfUsersToDBSQLServerName)
        Me.tabLoadListOfUsersToDB.Controls.Add(Me.Label10)
        Me.tabLoadListOfUsersToDB.Location = New System.Drawing.Point(4, 22)
        Me.tabLoadListOfUsersToDB.Name = "tabLoadListOfUsersToDB"
        Me.tabLoadListOfUsersToDB.Padding = New System.Windows.Forms.Padding(3)
        Me.tabLoadListOfUsersToDB.Size = New System.Drawing.Size(550, 92)
        Me.tabLoadListOfUsersToDB.TabIndex = 4
        Me.tabLoadListOfUsersToDB.Text = "Load List of Users to DB"
        Me.tabLoadListOfUsersToDB.UseVisualStyleBackColor = True
        '
        'txtSIUserId
        '
        Me.txtSIUserId.Location = New System.Drawing.Point(104, 13)
        Me.txtSIUserId.Name = "txtSIUserId"
        Me.txtSIUserId.Size = New System.Drawing.Size(220, 20)
        Me.txtSIUserId.TabIndex = 18
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(1, 10)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(105, 23)
        Me.Label16.TabIndex = 17
        Me.Label16.Text = "SI_ID:"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtLoadListOfUsersToDBDatabaseName
        '
        Me.txtLoadListOfUsersToDBDatabaseName.Location = New System.Drawing.Point(104, 64)
        Me.txtLoadListOfUsersToDBDatabaseName.Name = "txtLoadListOfUsersToDBDatabaseName"
        Me.txtLoadListOfUsersToDBDatabaseName.Size = New System.Drawing.Size(220, 20)
        Me.txtLoadListOfUsersToDBDatabaseName.TabIndex = 10
        Me.txtLoadListOfUsersToDBDatabaseName.Text = "BI_Configuration"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(4, 62)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(105, 23)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "&Database Name:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnLoadListOfUsersToDB
        '
        Me.btnLoadListOfUsersToDB.Location = New System.Drawing.Point(375, 24)
        Me.btnLoadListOfUsersToDB.Name = "btnLoadListOfUsersToDB"
        Me.btnLoadListOfUsersToDB.Size = New System.Drawing.Size(97, 46)
        Me.btnLoadListOfUsersToDB.TabIndex = 11
        Me.btnLoadListOfUsersToDB.Text = "Load List of Users to DB"
        Me.btnLoadListOfUsersToDB.UseVisualStyleBackColor = True
        '
        'txtLoadListOfUsersToDBSQLServerName
        '
        Me.txtLoadListOfUsersToDBSQLServerName.Location = New System.Drawing.Point(104, 38)
        Me.txtLoadListOfUsersToDBSQLServerName.Name = "txtLoadListOfUsersToDBSQLServerName"
        Me.txtLoadListOfUsersToDBSQLServerName.Size = New System.Drawing.Size(220, 20)
        Me.txtLoadListOfUsersToDBSQLServerName.TabIndex = 8
        Me.txtLoadListOfUsersToDBSQLServerName.Text = "SQLAG02"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(4, 36)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(105, 23)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "&SQL Server:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tabExtractRepo
        '
        Me.tabExtractRepo.Controls.Add(Me.txtSIID)
        Me.tabExtractRepo.Controls.Add(Me.Label15)
        Me.tabExtractRepo.Controls.Add(Me.txtExtractRepoDatabaseName)
        Me.tabExtractRepo.Controls.Add(Me.Label11)
        Me.tabExtractRepo.Controls.Add(Me.txtExtractRepoSQLServer)
        Me.tabExtractRepo.Controls.Add(Me.Label12)
        Me.tabExtractRepo.Controls.Add(Me.btnExtractRepo)
        Me.tabExtractRepo.Location = New System.Drawing.Point(4, 22)
        Me.tabExtractRepo.Name = "tabExtractRepo"
        Me.tabExtractRepo.Padding = New System.Windows.Forms.Padding(3)
        Me.tabExtractRepo.Size = New System.Drawing.Size(550, 92)
        Me.tabExtractRepo.TabIndex = 5
        Me.tabExtractRepo.Text = "Extract Repo"
        Me.tabExtractRepo.UseVisualStyleBackColor = True
        '
        'txtSIID
        '
        Me.txtSIID.Location = New System.Drawing.Point(106, 7)
        Me.txtSIID.Name = "txtSIID"
        Me.txtSIID.Size = New System.Drawing.Size(220, 20)
        Me.txtSIID.TabIndex = 16
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(3, 4)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(105, 23)
        Me.Label15.TabIndex = 15
        Me.Label15.Text = "SI_ID:"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExtractRepoDatabaseName
        '
        Me.txtExtractRepoDatabaseName.Location = New System.Drawing.Point(106, 56)
        Me.txtExtractRepoDatabaseName.Name = "txtExtractRepoDatabaseName"
        Me.txtExtractRepoDatabaseName.Size = New System.Drawing.Size(220, 20)
        Me.txtExtractRepoDatabaseName.TabIndex = 14
        Me.txtExtractRepoDatabaseName.Text = "BI_Configuration"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(3, 53)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(105, 23)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "&Database Name:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExtractRepoSQLServer
        '
        Me.txtExtractRepoSQLServer.Location = New System.Drawing.Point(106, 30)
        Me.txtExtractRepoSQLServer.Name = "txtExtractRepoSQLServer"
        Me.txtExtractRepoSQLServer.Size = New System.Drawing.Size(220, 20)
        Me.txtExtractRepoSQLServer.TabIndex = 12
        Me.txtExtractRepoSQLServer.Text = "SQLAG02"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(3, 27)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(105, 23)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "&SQL Server:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnExtractRepo
        '
        Me.btnExtractRepo.Location = New System.Drawing.Point(358, 30)
        Me.btnExtractRepo.Name = "btnExtractRepo"
        Me.btnExtractRepo.Size = New System.Drawing.Size(75, 39)
        Me.btnExtractRepo.TabIndex = 0
        Me.btnExtractRepo.Text = "Extract Repo"
        Me.btnExtractRepo.UseVisualStyleBackColor = True
        '
        'tabAddDBCredentials
        '
        Me.tabAddDBCredentials.Controls.Add(Me.GroupBox7)
        Me.tabAddDBCredentials.Location = New System.Drawing.Point(4, 22)
        Me.tabAddDBCredentials.Name = "tabAddDBCredentials"
        Me.tabAddDBCredentials.Size = New System.Drawing.Size(550, 92)
        Me.tabAddDBCredentials.TabIndex = 6
        Me.tabAddDBCredentials.Text = "Add DB Credentials"
        Me.tabAddDBCredentials.UseVisualStyleBackColor = True
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.txtAddDBCredentialsPW)
        Me.GroupBox7.Controls.Add(Me.Label13)
        Me.GroupBox7.Controls.Add(Me.btnAddDBCredentialsForGroup)
        Me.GroupBox7.Controls.Add(Me.txtGroupNameToProcess)
        Me.GroupBox7.Controls.Add(Me.Label14)
        Me.GroupBox7.Location = New System.Drawing.Point(7, 7)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(537, 79)
        Me.GroupBox7.TabIndex = 2
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Enter Info"
        '
        'txtAddDBCredentialsPW
        '
        Me.txtAddDBCredentialsPW.Location = New System.Drawing.Point(76, 45)
        Me.txtAddDBCredentialsPW.Name = "txtAddDBCredentialsPW"
        Me.txtAddDBCredentialsPW.Size = New System.Drawing.Size(252, 20)
        Me.txtAddDBCredentialsPW.TabIndex = 13
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 43)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(105, 23)
        Me.Label13.TabIndex = 12
        Me.Label13.Text = "&Password:"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnAddDBCredentialsForGroup
        '
        Me.btnAddDBCredentialsForGroup.Location = New System.Drawing.Point(347, 19)
        Me.btnAddDBCredentialsForGroup.Name = "btnAddDBCredentialsForGroup"
        Me.btnAddDBCredentialsForGroup.Size = New System.Drawing.Size(184, 21)
        Me.btnAddDBCredentialsForGroup.TabIndex = 11
        Me.btnAddDBCredentialsForGroup.Text = "Add DB Credentials for Group"
        Me.btnAddDBCredentialsForGroup.UseVisualStyleBackColor = True
        '
        'txtGroupNameToProcess
        '
        Me.txtGroupNameToProcess.Location = New System.Drawing.Point(76, 19)
        Me.txtGroupNameToProcess.Name = "txtGroupNameToProcess"
        Me.txtGroupNameToProcess.Size = New System.Drawing.Size(252, 20)
        Me.txtGroupNameToProcess.TabIndex = 8
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 17)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(105, 23)
        Me.Label14.TabIndex = 7
        Me.Label14.Text = "&Group Name:"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tabResetAdminPW
        '
        Me.tabResetAdminPW.Controls.Add(Me.btnResetAdminPW)
        Me.tabResetAdminPW.Location = New System.Drawing.Point(4, 22)
        Me.tabResetAdminPW.Name = "tabResetAdminPW"
        Me.tabResetAdminPW.Size = New System.Drawing.Size(550, 92)
        Me.tabResetAdminPW.TabIndex = 7
        Me.tabResetAdminPW.Text = "Reset Admin PW"
        Me.tabResetAdminPW.UseVisualStyleBackColor = True
        '
        'btnResetAdminPW
        '
        Me.btnResetAdminPW.Location = New System.Drawing.Point(234, 27)
        Me.btnResetAdminPW.Name = "btnResetAdminPW"
        Me.btnResetAdminPW.Size = New System.Drawing.Size(75, 40)
        Me.btnResetAdminPW.TabIndex = 0
        Me.btnResetAdminPW.Text = "Reset Admin PW"
        Me.btnResetAdminPW.UseVisualStyleBackColor = True
        '
        'tabUpdateOwnersOnDocsFromCSV
        '
        Me.tabUpdateOwnersOnDocsFromCSV.Controls.Add(Me.chkOwnerIsSameOnAllDocs)
        Me.tabUpdateOwnersOnDocsFromCSV.Controls.Add(Me.Label22)
        Me.tabUpdateOwnersOnDocsFromCSV.Controls.Add(Me.Label21)
        Me.tabUpdateOwnersOnDocsFromCSV.Controls.Add(Me.txtCSVFilePath)
        Me.tabUpdateOwnersOnDocsFromCSV.Controls.Add(Me.btnUpdateOwnersOnDocsFromCSV)
        Me.tabUpdateOwnersOnDocsFromCSV.Location = New System.Drawing.Point(4, 22)
        Me.tabUpdateOwnersOnDocsFromCSV.Name = "tabUpdateOwnersOnDocsFromCSV"
        Me.tabUpdateOwnersOnDocsFromCSV.Padding = New System.Windows.Forms.Padding(3)
        Me.tabUpdateOwnersOnDocsFromCSV.Size = New System.Drawing.Size(550, 92)
        Me.tabUpdateOwnersOnDocsFromCSV.TabIndex = 8
        Me.tabUpdateOwnersOnDocsFromCSV.Text = "Update Owners On Docs from CSV"
        Me.tabUpdateOwnersOnDocsFromCSV.UseVisualStyleBackColor = True
        '
        'chkOwnerIsSameOnAllDocs
        '
        Me.chkOwnerIsSameOnAllDocs.AutoSize = True
        Me.chkOwnerIsSameOnAllDocs.Location = New System.Drawing.Point(9, 45)
        Me.chkOwnerIsSameOnAllDocs.Name = "chkOwnerIsSameOnAllDocs"
        Me.chkOwnerIsSameOnAllDocs.Size = New System.Drawing.Size(149, 17)
        Me.chkOwnerIsSameOnAllDocs.TabIndex = 25
        Me.chkOwnerIsSameOnAllDocs.Text = "Owner is same on all docs"
        Me.chkOwnerIsSameOnAllDocs.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(6, 17)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(81, 23)
        Me.Label22.TabIndex = 24
        Me.Label22.Text = "File Location:"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(3, 60)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(261, 29)
        Me.Label21.TabIndex = 23
        Me.Label21.Text = "File Format: CSV SI_ID,NewOwnerName (No Header)"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCSVFilePath
        '
        Me.txtCSVFilePath.Location = New System.Drawing.Point(99, 19)
        Me.txtCSVFilePath.Name = "txtCSVFilePath"
        Me.txtCSVFilePath.ReadOnly = True
        Me.txtCSVFilePath.Size = New System.Drawing.Size(249, 20)
        Me.txtCSVFilePath.TabIndex = 1
        Me.txtCSVFilePath.Text = "C:\SAP\ListOfObjectsAndOwner.txt"
        '
        'btnUpdateOwnersOnDocsFromCSV
        '
        Me.btnUpdateOwnersOnDocsFromCSV.Location = New System.Drawing.Point(354, 13)
        Me.btnUpdateOwnersOnDocsFromCSV.Name = "btnUpdateOwnersOnDocsFromCSV"
        Me.btnUpdateOwnersOnDocsFromCSV.Size = New System.Drawing.Size(179, 31)
        Me.btnUpdateOwnersOnDocsFromCSV.TabIndex = 0
        Me.btnUpdateOwnersOnDocsFromCSV.Text = "Update Owners on Docs from CSV"
        Me.btnUpdateOwnersOnDocsFromCSV.UseVisualStyleBackColor = True
        '
        'tabLoadObjectPropertiesToDB
        '
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.Label20)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.txtLoadObjectPropertyDelta)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.txtDeltaTsp)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.btnCheckDeltaTsp)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.txtLoadObjectPropertiesSIID)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.Label17)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.txtLoadObjectPropertiesDB)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.Label18)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.txtLoadObjectPropertiesServer)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.Label19)
        Me.tabLoadObjectPropertiesToDB.Controls.Add(Me.btnLoadObjectPropertiesToDB)
        Me.tabLoadObjectPropertiesToDB.Location = New System.Drawing.Point(4, 22)
        Me.tabLoadObjectPropertiesToDB.Name = "tabLoadObjectPropertiesToDB"
        Me.tabLoadObjectPropertiesToDB.Padding = New System.Windows.Forms.Padding(3)
        Me.tabLoadObjectPropertiesToDB.Size = New System.Drawing.Size(550, 92)
        Me.tabLoadObjectPropertiesToDB.TabIndex = 9
        Me.tabLoadObjectPropertiesToDB.Text = "Load Object Properties to DB"
        Me.tabLoadObjectPropertiesToDB.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(344, 49)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(108, 13)
        Me.Label20.TabIndex = 27
        Me.Label20.Text = "Deltas (1=Yes, 0=No)"
        '
        'txtLoadObjectPropertyDelta
        '
        Me.txtLoadObjectPropertyDelta.Location = New System.Drawing.Point(344, 65)
        Me.txtLoadObjectPropertyDelta.Name = "txtLoadObjectPropertyDelta"
        Me.txtLoadObjectPropertyDelta.Size = New System.Drawing.Size(100, 20)
        Me.txtLoadObjectPropertyDelta.TabIndex = 26
        Me.txtLoadObjectPropertyDelta.Text = "0"
        '
        'txtDeltaTsp
        '
        Me.txtDeltaTsp.Location = New System.Drawing.Point(344, 23)
        Me.txtDeltaTsp.Name = "txtDeltaTsp"
        Me.txtDeltaTsp.Size = New System.Drawing.Size(185, 20)
        Me.txtDeltaTsp.TabIndex = 25
        '
        'btnCheckDeltaTsp
        '
        Me.btnCheckDeltaTsp.Location = New System.Drawing.Point(344, 0)
        Me.btnCheckDeltaTsp.Name = "btnCheckDeltaTsp"
        Me.btnCheckDeltaTsp.Size = New System.Drawing.Size(93, 23)
        Me.btnCheckDeltaTsp.TabIndex = 24
        Me.btnCheckDeltaTsp.Text = "Get Delta Tsp"
        Me.btnCheckDeltaTsp.UseVisualStyleBackColor = True
        '
        'txtLoadObjectPropertiesSIID
        '
        Me.txtLoadObjectPropertiesSIID.Location = New System.Drawing.Point(109, 11)
        Me.txtLoadObjectPropertiesSIID.Name = "txtLoadObjectPropertiesSIID"
        Me.txtLoadObjectPropertiesSIID.Size = New System.Drawing.Size(220, 20)
        Me.txtLoadObjectPropertiesSIID.TabIndex = 23
        Me.txtLoadObjectPropertiesSIID.Text = "22292"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(6, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(105, 23)
        Me.Label17.TabIndex = 22
        Me.Label17.Text = "SI_ID:"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtLoadObjectPropertiesDB
        '
        Me.txtLoadObjectPropertiesDB.Location = New System.Drawing.Point(109, 60)
        Me.txtLoadObjectPropertiesDB.Name = "txtLoadObjectPropertiesDB"
        Me.txtLoadObjectPropertiesDB.Size = New System.Drawing.Size(220, 20)
        Me.txtLoadObjectPropertiesDB.TabIndex = 21
        Me.txtLoadObjectPropertiesDB.Text = "BI_Configuration"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(6, 57)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(105, 23)
        Me.Label18.TabIndex = 20
        Me.Label18.Text = "&Database Name:"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtLoadObjectPropertiesServer
        '
        Me.txtLoadObjectPropertiesServer.Location = New System.Drawing.Point(109, 34)
        Me.txtLoadObjectPropertiesServer.Name = "txtLoadObjectPropertiesServer"
        Me.txtLoadObjectPropertiesServer.Size = New System.Drawing.Size(220, 20)
        Me.txtLoadObjectPropertiesServer.TabIndex = 19
        Me.txtLoadObjectPropertiesServer.Text = "SQLDEV"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(6, 31)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(105, 23)
        Me.Label19.TabIndex = 18
        Me.Label19.Text = "&SQL Server:"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnLoadObjectPropertiesToDB
        '
        Me.btnLoadObjectPropertiesToDB.Location = New System.Drawing.Point(469, 49)
        Me.btnLoadObjectPropertiesToDB.Name = "btnLoadObjectPropertiesToDB"
        Me.btnLoadObjectPropertiesToDB.Size = New System.Drawing.Size(75, 39)
        Me.btnLoadObjectPropertiesToDB.TabIndex = 17
        Me.btnLoadObjectPropertiesToDB.Text = "Extract Repo"
        Me.btnLoadObjectPropertiesToDB.UseVisualStyleBackColor = True
        '
        'rtbOutput
        '
        Me.rtbOutput.Location = New System.Drawing.Point(8, 294)
        Me.rtbOutput.Name = "rtbOutput"
        Me.rtbOutput.Size = New System.Drawing.Size(554, 479)
        Me.rtbOutput.TabIndex = 11
        Me.rtbOutput.Text = ""
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 274)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 13)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Output"
        '
        'tabDeleteReportsBySIID
        '
        Me.tabDeleteReportsBySIID.Controls.Add(Me.Label23)
        Me.tabDeleteReportsBySIID.Controls.Add(Me.Label24)
        Me.tabDeleteReportsBySIID.Controls.Add(Me.txtFileNameLocationForDeletes)
        Me.tabDeleteReportsBySIID.Controls.Add(Me.btnDeleteReportsFromFile)
        Me.tabDeleteReportsBySIID.Location = New System.Drawing.Point(4, 22)
        Me.tabDeleteReportsBySIID.Name = "tabDeleteReportsBySIID"
        Me.tabDeleteReportsBySIID.Padding = New System.Windows.Forms.Padding(3)
        Me.tabDeleteReportsBySIID.Size = New System.Drawing.Size(550, 92)
        Me.tabDeleteReportsBySIID.TabIndex = 10
        Me.tabDeleteReportsBySIID.Text = "Delete Reports by SI_ID"
        Me.tabDeleteReportsBySIID.UseVisualStyleBackColor = True
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(13, 12)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(76, 23)
        Me.Label23.TabIndex = 28
        Me.Label23.Text = "File Location:"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(13, 47)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(261, 29)
        Me.Label24.TabIndex = 27
        Me.Label24.Text = "File Format: SI_ID (No Header)"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileNameLocationForDeletes
        '
        Me.txtFileNameLocationForDeletes.Location = New System.Drawing.Point(95, 15)
        Me.txtFileNameLocationForDeletes.Name = "txtFileNameLocationForDeletes"
        Me.txtFileNameLocationForDeletes.ReadOnly = True
        Me.txtFileNameLocationForDeletes.Size = New System.Drawing.Size(248, 20)
        Me.txtFileNameLocationForDeletes.TabIndex = 26
        Me.txtFileNameLocationForDeletes.Text = "C:\SAP\ListOfSIIDsToDelete.txt"
        '
        'btnDeleteReportsFromFile
        '
        Me.btnDeleteReportsFromFile.Location = New System.Drawing.Point(361, 8)
        Me.btnDeleteReportsFromFile.Name = "btnDeleteReportsFromFile"
        Me.btnDeleteReportsFromFile.Size = New System.Drawing.Size(179, 31)
        Me.btnDeleteReportsFromFile.TabIndex = 25
        Me.btnDeleteReportsFromFile.Text = "Delete Reports from File"
        Me.btnDeleteReportsFromFile.UseVisualStyleBackColor = True
        '
        'frmTools
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(574, 785)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.rtbOutput)
        Me.Controls.Add(Me.tabFunctions)
        Me.Controls.Add(Me.cboCMSAuthentication)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtCMSUserPassword)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboCMSServer)
        Me.Controls.Add(Me.txtCMSUserName)
        Me.Controls.Add(Me.PasswordLabel)
        Me.Controls.Add(Me.UsernameLabel)
        Me.Controls.Add(Me.LogoPictureBox)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTools"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Tools"
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabFunctions.ResumeLayout(False)
        Me.tabListObjectsByOwner.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.tabReplaceOwnerOnAllObjects.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.tabReplaceOwnerOnSingleDoc.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.tabGetListOfUsers.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.tabLoadListOfUsersToDB.ResumeLayout(False)
        Me.tabLoadListOfUsersToDB.PerformLayout()
        Me.tabExtractRepo.ResumeLayout(False)
        Me.tabExtractRepo.PerformLayout()
        Me.tabAddDBCredentials.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.tabResetAdminPW.ResumeLayout(False)
        Me.tabUpdateOwnersOnDocsFromCSV.ResumeLayout(False)
        Me.tabUpdateOwnersOnDocsFromCSV.PerformLayout()
        Me.tabLoadObjectPropertiesToDB.ResumeLayout(False)
        Me.tabLoadObjectPropertiesToDB.PerformLayout()
        Me.tabDeleteReportsBySIID.ResumeLayout(False)
        Me.tabDeleteReportsBySIID.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cboCMSServer As ComboBox
    Friend WithEvents txtCMSUserPassword As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cboCMSAuthentication As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents tabFunctions As TabControl
    Friend WithEvents tabListObjectsByOwner As TabPage
    Friend WithEvents tabReplaceOwnerOnAllObjects As TabPage
    Friend WithEvents tabReplaceOwnerOnSingleDoc As TabPage
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents txtReplaceOwnerOnAllObjectsOwnerNameNew As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtReplaceOwnerOnAllObjectsOwnerNameOld As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents btnRepalceOwnerOnAllObjects As Button
    Friend WithEvents btnReplaceOwnerOnSingeDoc As Button
    Friend WithEvents txtReplaceOwnerOnSingleDocOwnerNameNew As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtReplaceOwnerOnSingleDocDocId As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents btnListObjectsByOwner As Button
    Friend WithEvents txtListObjectsByOwnerOwnerName As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents rtbOutput As RichTextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents tabGetListOfUsers As TabPage
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents btnGetListOfUsers As Button
    Friend WithEvents tabLoadListOfUsersToDB As TabPage
    Friend WithEvents tabExtractRepo As TabPage
    Friend WithEvents txtExtractRepoDatabaseName As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents txtExtractRepoSQLServer As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents btnExtractRepo As Button
    Friend WithEvents tabAddDBCredentials As TabPage
    Friend WithEvents GroupBox7 As GroupBox
    Friend WithEvents btnAddDBCredentialsForGroup As Button
    Friend WithEvents txtGroupNameToProcess As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txtAddDBCredentialsPW As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtSIID As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents txtSIUserId As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents txtLoadListOfUsersToDBDatabaseName As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents btnLoadListOfUsersToDB As Button
    Friend WithEvents txtLoadListOfUsersToDBSQLServerName As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents tabResetAdminPW As TabPage
    Friend WithEvents btnResetAdminPW As Button
    Friend WithEvents tabUpdateOwnersOnDocsFromCSV As TabPage
    Friend WithEvents txtCSVFilePath As TextBox
    Friend WithEvents btnUpdateOwnersOnDocsFromCSV As Button
    Friend WithEvents tabLoadObjectPropertiesToDB As TabPage
    Friend WithEvents txtLoadObjectPropertiesSIID As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents txtLoadObjectPropertiesDB As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txtLoadObjectPropertiesServer As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents btnLoadObjectPropertiesToDB As Button
    Friend WithEvents txtDeltaTsp As TextBox
    Friend WithEvents btnCheckDeltaTsp As Button
    Friend WithEvents Label20 As Label
    Friend WithEvents txtLoadObjectPropertyDelta As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents chkOwnerIsSameOnAllDocs As CheckBox
    Friend WithEvents tabDeleteReportsBySIID As TabPage
    Friend WithEvents Label23 As Label
    Friend WithEvents Label24 As Label
    Friend WithEvents txtFileNameLocationForDeletes As TextBox
    Friend WithEvents btnDeleteReportsFromFile As Button
End Class
