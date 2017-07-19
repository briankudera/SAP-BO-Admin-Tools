<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
<Global.System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726")> _
Partial Class Tools
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Tools))
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
        Me.tabReplaceOwnerOnAllObjects = New System.Windows.Forms.TabPage()
        Me.tabReplaceOwnerOnSingleDoc = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtReplaceOwnerOnAllObjectsOwnerNameNew = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtReplaceOwnerOnAllObjectsOwnerNameOld = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnRepalceOwnerOnAllObjects = New System.Windows.Forms.Button()
        Me.btnReplaceOwnerOnSingeDoc = New System.Windows.Forms.Button()
        Me.txtReplaceOwnerOnSingleDocOwnerNameNew = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtReplaceOwnerOnSingleDocDocId = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnListObjectsByOwner = New System.Windows.Forms.Button()
        Me.txtListObjectsByOwnerOwnerName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rtbOutput = New System.Windows.Forms.RichTextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabFunctions.SuspendLayout()
        Me.tabListObjectsByOwner.SuspendLayout()
        Me.tabReplaceOwnerOnAllObjects.SuspendLayout()
        Me.tabReplaceOwnerOnSingleDoc.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
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
        Me.cboCMSServer.Items.AddRange(New Object() {"MSP-BICMS01:6400", "TST-BICMS:6400"})
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
        Me.txtCMSUserPassword.Text = "Administrator"
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
        'btnRepalceOwnerOnAllObjects
        '
        Me.btnRepalceOwnerOnAllObjects.Location = New System.Drawing.Point(347, 19)
        Me.btnRepalceOwnerOnAllObjects.Name = "btnRepalceOwnerOnAllObjects"
        Me.btnRepalceOwnerOnAllObjects.Size = New System.Drawing.Size(97, 46)
        Me.btnRepalceOwnerOnAllObjects.TabIndex = 11
        Me.btnRepalceOwnerOnAllObjects.Text = "Replace Owner on All Objects"
        Me.btnRepalceOwnerOnAllObjects.UseVisualStyleBackColor = True
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
        'Tools
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
        Me.Name = "Tools"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Tools"
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabFunctions.ResumeLayout(False)
        Me.tabListObjectsByOwner.ResumeLayout(False)
        Me.tabReplaceOwnerOnAllObjects.ResumeLayout(False)
        Me.tabReplaceOwnerOnSingleDoc.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
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
End Class
