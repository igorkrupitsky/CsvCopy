Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices


Public Class Form1
	Inherits System.Windows.Forms.Form

	Private bStop As Boolean = False

	Dim oAppSetting As New AppSetting
	Friend WithEvents chkHideNotSelected As System.Windows.Forms.CheckBox
	Dim oSelTables As New Hashtable
	Dim oSqlScriptWriter As System.IO.StreamWriter = Nothing

	Friend WithEvents selDelimiter As ComboBox
	Friend WithEvents chkShrinkTable As CheckBox
	Friend WithEvents btnDeleteFolderCache As Button
	Friend WithEvents btnShrinkTables As Button
	Friend WithEvents chkRec_id As CheckBox
	Friend WithEvents chkNvarChar As CheckBox
	Friend WithEvents chkScriptToFile As CheckBox
	Friend WithEvents ToolTip1 As ToolTip
	Friend WithEvents chkBulkInsert As CheckBox
	Friend WithEvents Label3 As Label

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents btnCopy As System.Windows.Forms.Button
	Friend WithEvents btnConnect1 As System.Windows.Forms.Button
	Friend WithEvents txtFolderPath As System.Windows.Forms.TextBox
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents chkCreateTable As System.Windows.Forms.CheckBox
	Friend WithEvents txtLog As System.Windows.Forms.TextBox
	Friend WithEvents chkDeleteData As System.Windows.Forms.CheckBox
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents txtConnectTo As System.Windows.Forms.TextBox
	Friend WithEvents btnConnect2 As System.Windows.Forms.Button
	Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
	Friend WithEvents dgTables As System.Windows.Forms.DataGridView
	Friend WithEvents chkDropTable As System.Windows.Forms.CheckBox
	Friend WithEvents btnCheckAll As System.Windows.Forms.Button
	Friend WithEvents btnUncheckAll As System.Windows.Forms.Button
	Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
	Friend WithEvents btnCheckNew As System.Windows.Forms.Button
	Friend WithEvents btnStop As System.Windows.Forms.LinkLabel
	Friend WithEvents lbCount As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.btnCopy = New System.Windows.Forms.Button()
		Me.txtFolderPath = New System.Windows.Forms.TextBox()
		Me.btnConnect1 = New System.Windows.Forms.Button()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.chkCreateTable = New System.Windows.Forms.CheckBox()
		Me.txtLog = New System.Windows.Forms.TextBox()
		Me.chkDeleteData = New System.Windows.Forms.CheckBox()
		Me.lbCount = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.txtConnectTo = New System.Windows.Forms.TextBox()
		Me.btnConnect2 = New System.Windows.Forms.Button()
		Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
		Me.dgTables = New System.Windows.Forms.DataGridView()
		Me.chkDropTable = New System.Windows.Forms.CheckBox()
		Me.btnCheckAll = New System.Windows.Forms.Button()
		Me.btnUncheckAll = New System.Windows.Forms.Button()
		Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
		Me.btnCheckNew = New System.Windows.Forms.Button()
		Me.btnStop = New System.Windows.Forms.LinkLabel()
		Me.chkHideNotSelected = New System.Windows.Forms.CheckBox()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.selDelimiter = New System.Windows.Forms.ComboBox()
		Me.chkShrinkTable = New System.Windows.Forms.CheckBox()
		Me.btnDeleteFolderCache = New System.Windows.Forms.Button()
		Me.btnShrinkTables = New System.Windows.Forms.Button()
		Me.chkRec_id = New System.Windows.Forms.CheckBox()
		Me.chkNvarChar = New System.Windows.Forms.CheckBox()
		Me.chkScriptToFile = New System.Windows.Forms.CheckBox()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.chkBulkInsert = New System.Windows.Forms.CheckBox()
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'btnCopy
		'
		Me.btnCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCopy.Location = New System.Drawing.Point(678, 188)
		Me.btnCopy.Name = "btnCopy"
		Me.btnCopy.Size = New System.Drawing.Size(100, 24)
		Me.btnCopy.TabIndex = 0
		Me.btnCopy.Text = "Copy tables"
		'
		'txtFolderPath
		'
		Me.txtFolderPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtFolderPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtFolderPath.Location = New System.Drawing.Point(73, 8)
		Me.txtFolderPath.Name = "txtFolderPath"
		Me.txtFolderPath.Size = New System.Drawing.Size(764, 20)
		Me.txtFolderPath.TabIndex = 1
		'
		'btnConnect1
		'
		Me.btnConnect1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnConnect1.Location = New System.Drawing.Point(925, 9)
		Me.btnConnect1.Name = "btnConnect1"
		Me.btnConnect1.Size = New System.Drawing.Size(75, 23)
		Me.btnConnect1.TabIndex = 3
		Me.btnConnect1.Text = "Connect"
		'
		'btnCancel
		'
		Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCancel.Location = New System.Drawing.Point(896, 186)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(104, 24)
		Me.btnCancel.TabIndex = 6
		Me.btnCancel.Text = "Cancel"
		'
		'ProgressBar1
		'
		Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar1.Location = New System.Drawing.Point(224, 352)
		Me.ProgressBar1.Name = "ProgressBar1"
		Me.ProgressBar1.Size = New System.Drawing.Size(691, 11)
		Me.ProgressBar1.TabIndex = 7
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(5, 13)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(36, 13)
		Me.Label1.TabIndex = 8
		Me.Label1.Text = "Folder"
		'
		'chkCreateTable
		'
		Me.chkCreateTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkCreateTable.AutoSize = True
		Me.chkCreateTable.Checked = True
		Me.chkCreateTable.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkCreateTable.Location = New System.Drawing.Point(10, 160)
		Me.chkCreateTable.Name = "chkCreateTable"
		Me.chkCreateTable.Size = New System.Drawing.Size(113, 17)
		Me.chkCreateTable.TabIndex = 11
		Me.chkCreateTable.Text = "Create target table"
		Me.chkCreateTable.UseVisualStyleBackColor = True
		'
		'txtLog
		'
		Me.txtLog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLog.Location = New System.Drawing.Point(8, 218)
		Me.txtLog.Multiline = True
		Me.txtLog.Name = "txtLog"
		Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtLog.Size = New System.Drawing.Size(992, 132)
		Me.txtLog.TabIndex = 12
		'
		'chkDeleteData
		'
		Me.chkDeleteData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDeleteData.AutoSize = True
		Me.chkDeleteData.Checked = True
		Me.chkDeleteData.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkDeleteData.Location = New System.Drawing.Point(256, 160)
		Me.chkDeleteData.Name = "chkDeleteData"
		Me.chkDeleteData.Size = New System.Drawing.Size(118, 17)
		Me.chkDeleteData.TabIndex = 13
		Me.chkDeleteData.Text = "Delete before insert"
		Me.chkDeleteData.UseVisualStyleBackColor = True
		'
		'lbCount
		'
		Me.lbCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lbCount.AutoSize = True
		Me.lbCount.Location = New System.Drawing.Point(955, 353)
		Me.lbCount.Name = "lbCount"
		Me.lbCount.Size = New System.Drawing.Size(25, 13)
		Me.lbCount.TabIndex = 14
		Me.lbCount.Text = "000"
		Me.lbCount.Visible = False
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(5, 41)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(62, 13)
		Me.Label2.TabIndex = 15
		Me.Label2.Text = "SQL Server"
		'
		'txtConnectTo
		'
		Me.txtConnectTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtConnectTo.Location = New System.Drawing.Point(73, 41)
		Me.txtConnectTo.Name = "txtConnectTo"
		Me.txtConnectTo.Size = New System.Drawing.Size(846, 20)
		Me.txtConnectTo.TabIndex = 16
		'
		'btnConnect2
		'
		Me.btnConnect2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnConnect2.Location = New System.Drawing.Point(925, 39)
		Me.btnConnect2.Name = "btnConnect2"
		Me.btnConnect2.Size = New System.Drawing.Size(75, 23)
		Me.btnConnect2.TabIndex = 17
		Me.btnConnect2.Text = "Connect"
		Me.btnConnect2.UseVisualStyleBackColor = True
		'
		'dgTables
		'
		Me.dgTables.AllowUserToAddRows = False
		Me.dgTables.AllowUserToDeleteRows = False
		Me.dgTables.AllowUserToOrderColumns = True
		Me.dgTables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.dgTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgTables.Location = New System.Drawing.Point(5, 68)
		Me.dgTables.Name = "dgTables"
		Me.dgTables.RowHeadersWidth = 62
		Me.dgTables.Size = New System.Drawing.Size(995, 85)
		Me.dgTables.TabIndex = 20
		'
		'chkDropTable
		'
		Me.chkDropTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDropTable.AutoSize = True
		Me.chkDropTable.Location = New System.Drawing.Point(10, 188)
		Me.chkDropTable.Name = "chkDropTable"
		Me.chkDropTable.Size = New System.Drawing.Size(112, 17)
		Me.chkDropTable.TabIndex = 33
		Me.chkDropTable.Text = "Drop table if exists"
		Me.chkDropTable.UseVisualStyleBackColor = True
		'
		'btnCheckAll
		'
		Me.btnCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCheckAll.Location = New System.Drawing.Point(789, 157)
		Me.btnCheckAll.Name = "btnCheckAll"
		Me.btnCheckAll.Size = New System.Drawing.Size(100, 23)
		Me.btnCheckAll.TabIndex = 35
		Me.btnCheckAll.Text = "Check All"
		Me.btnCheckAll.UseVisualStyleBackColor = True
		'
		'btnUncheckAll
		'
		Me.btnUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnUncheckAll.Location = New System.Drawing.Point(896, 157)
		Me.btnUncheckAll.Name = "btnUncheckAll"
		Me.btnUncheckAll.Size = New System.Drawing.Size(104, 23)
		Me.btnUncheckAll.TabIndex = 36
		Me.btnUncheckAll.Text = "Uncheck All"
		Me.btnUncheckAll.UseVisualStyleBackColor = True
		'
		'ProgressBar2
		'
		Me.ProgressBar2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar2.Location = New System.Drawing.Point(5, 352)
		Me.ProgressBar2.Name = "ProgressBar2"
		Me.ProgressBar2.Size = New System.Drawing.Size(213, 11)
		Me.ProgressBar2.TabIndex = 38
		'
		'btnCheckNew
		'
		Me.btnCheckNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCheckNew.Location = New System.Drawing.Point(678, 157)
		Me.btnCheckNew.Name = "btnCheckNew"
		Me.btnCheckNew.Size = New System.Drawing.Size(100, 23)
		Me.btnCheckNew.TabIndex = 37
		Me.btnCheckNew.Text = "Check New Rec"
		Me.btnCheckNew.UseVisualStyleBackColor = True
		'
		'btnStop
		'
		Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnStop.AutoSize = True
		Me.btnStop.Location = New System.Drawing.Point(920, 352)
		Me.btnStop.Name = "btnStop"
		Me.btnStop.Size = New System.Drawing.Size(29, 13)
		Me.btnStop.TabIndex = 40
		Me.btnStop.TabStop = True
		Me.btnStop.Text = "Stop"
		Me.btnStop.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.btnStop.Visible = False
		'
		'chkHideNotSelected
		'
		Me.chkHideNotSelected.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkHideNotSelected.AutoSize = True
		Me.chkHideNotSelected.Location = New System.Drawing.Point(382, 160)
		Me.chkHideNotSelected.Name = "chkHideNotSelected"
		Me.chkHideNotSelected.Size = New System.Drawing.Size(109, 17)
		Me.chkHideNotSelected.TabIndex = 42
		Me.chkHideNotSelected.Text = "Hide not selected"
		Me.chkHideNotSelected.UseVisualStyleBackColor = True
		'
		'Label3
		'
		Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(509, 194)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(47, 13)
		Me.Label3.TabIndex = 43
		Me.Label3.Text = "Delimiter"
		'
		'selDelimiter
		'
		Me.selDelimiter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.selDelimiter.FormattingEnabled = True
		Me.selDelimiter.Items.AddRange(New Object() {",", "|", "Tab"})
		Me.selDelimiter.Location = New System.Drawing.Point(558, 191)
		Me.selDelimiter.Name = "selDelimiter"
		Me.selDelimiter.Size = New System.Drawing.Size(75, 21)
		Me.selDelimiter.TabIndex = 44
		'
		'chkShrinkTable
		'
		Me.chkShrinkTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkShrinkTable.AutoSize = True
		Me.chkShrinkTable.Checked = True
		Me.chkShrinkTable.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkShrinkTable.Location = New System.Drawing.Point(256, 188)
		Me.chkShrinkTable.Name = "chkShrinkTable"
		Me.chkShrinkTable.Size = New System.Drawing.Size(82, 17)
		Me.chkShrinkTable.TabIndex = 45
		Me.chkShrinkTable.Text = "Shrink table"
		Me.chkShrinkTable.UseVisualStyleBackColor = True
		'
		'btnDeleteFolderCache
		'
		Me.btnDeleteFolderCache.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnDeleteFolderCache.Location = New System.Drawing.Point(843, 9)
		Me.btnDeleteFolderCache.Name = "btnDeleteFolderCache"
		Me.btnDeleteFolderCache.Size = New System.Drawing.Size(76, 23)
		Me.btnDeleteFolderCache.TabIndex = 46
		Me.btnDeleteFolderCache.Text = "Delete Cache"
		Me.btnDeleteFolderCache.Visible = False
		'
		'btnShrinkTables
		'
		Me.btnShrinkTables.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnShrinkTables.Location = New System.Drawing.Point(789, 188)
		Me.btnShrinkTables.Name = "btnShrinkTables"
		Me.btnShrinkTables.Size = New System.Drawing.Size(100, 24)
		Me.btnShrinkTables.TabIndex = 47
		Me.btnShrinkTables.Text = "Shrink tables"
		'
		'chkRec_id
		'
		Me.chkRec_id.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkRec_id.AutoSize = True
		Me.chkRec_id.Checked = True
		Me.chkRec_id.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkRec_id.Location = New System.Drawing.Point(132, 160)
		Me.chkRec_id.Name = "chkRec_id"
		Me.chkRec_id.Size = New System.Drawing.Size(77, 17)
		Me.chkRec_id.TabIndex = 48
		Me.chkRec_id.Text = "Add rec_id"
		Me.chkRec_id.UseVisualStyleBackColor = True
		'
		'chkNvarChar
		'
		Me.chkNvarChar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkNvarChar.AutoSize = True
		Me.chkNvarChar.Checked = True
		Me.chkNvarChar.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkNvarChar.Location = New System.Drawing.Point(132, 188)
		Me.chkNvarChar.Name = "chkNvarChar"
		Me.chkNvarChar.Size = New System.Drawing.Size(90, 17)
		Me.chkNvarChar.TabIndex = 49
		Me.chkNvarChar.Text = "Use nvarchar"
		Me.chkNvarChar.UseVisualStyleBackColor = True
		'
		'chkScriptToFile
		'
		Me.chkScriptToFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkScriptToFile.AutoSize = True
		Me.chkScriptToFile.Location = New System.Drawing.Point(512, 160)
		Me.chkScriptToFile.Name = "chkScriptToFile"
		Me.chkScriptToFile.Size = New System.Drawing.Size(88, 17)
		Me.chkScriptToFile.TabIndex = 50
		Me.chkScriptToFile.Text = "Script To File"
		Me.chkScriptToFile.UseVisualStyleBackColor = True
		'
		'chkBulkInsert
		'
		Me.chkBulkInsert.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkBulkInsert.AutoSize = True
		Me.chkBulkInsert.Checked = True
		Me.chkBulkInsert.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkBulkInsert.Location = New System.Drawing.Point(382, 188)
		Me.chkBulkInsert.Name = "chkBulkInsert"
		Me.chkBulkInsert.Size = New System.Drawing.Size(76, 17)
		Me.chkBulkInsert.TabIndex = 51
		Me.chkBulkInsert.Text = "Bulk Insert"
		Me.chkBulkInsert.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(1019, 368)
		Me.Controls.Add(Me.chkBulkInsert)
		Me.Controls.Add(Me.chkScriptToFile)
		Me.Controls.Add(Me.chkNvarChar)
		Me.Controls.Add(Me.chkRec_id)
		Me.Controls.Add(Me.btnUncheckAll)
		Me.Controls.Add(Me.btnCheckAll)
		Me.Controls.Add(Me.btnCheckNew)
		Me.Controls.Add(Me.btnShrinkTables)
		Me.Controls.Add(Me.btnDeleteFolderCache)
		Me.Controls.Add(Me.chkShrinkTable)
		Me.Controls.Add(Me.selDelimiter)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.chkHideNotSelected)
		Me.Controls.Add(Me.dgTables)
		Me.Controls.Add(Me.btnStop)
		Me.Controls.Add(Me.ProgressBar2)
		Me.Controls.Add(Me.chkDropTable)
		Me.Controls.Add(Me.btnConnect2)
		Me.Controls.Add(Me.txtConnectTo)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.lbCount)
		Me.Controls.Add(Me.chkDeleteData)
		Me.Controls.Add(Me.txtLog)
		Me.Controls.Add(Me.chkCreateTable)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.ProgressBar1)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnConnect1)
		Me.Controls.Add(Me.txtFolderPath)
		Me.Controls.Add(Me.btnCopy)
		Me.Name = "Form1"
		Me.Text = "Copy tables from CSV to SQL Server"
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

#End Region

	Private Sub frmExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

		txtFolderPath.Text = oAppSetting.GetSetting("FolderPath")
		txtConnectTo.Text = oAppSetting.GetSetting("ConnectTo")
		oSelTables = GetSelectedTablesFromReg()

		chkCreateTable.Checked = GetBoolSetting("chkCreateTable", True)
		chkDropTable.Checked = GetBoolSetting("chkDropTable", True)

		chkRec_id.Checked = GetBoolSetting("chkRec_id", True)
		chkDeleteData.Checked = GetBoolSetting("chkDeleteData", True)

		chkHideNotSelected.Checked = GetBoolSetting("chkHideNotSelected", False)
		chkShrinkTable.Checked = GetBoolSetting("chkShrinkTable", True)

		chkBulkInsert.Checked = GetBoolSetting("chkBulkInsert", False)
		chkScriptToFile.Checked = GetBoolSetting("chkScriptToFile", False)

		chkCreateTable_CheckedChanged()

		Dim sDelimiter As String = oAppSetting.GetSetting("Delimiter")
		If sDelimiter <> "" Then
			Try
				selDelimiter.SelectedIndex = selDelimiter.FindString(sDelimiter)
			Catch ex As Exception
				'Ignore
			End Try
		End If

		If selDelimiter.SelectedIndex = -1 Then
			selDelimiter.SelectedIndex = 0
		End If

		ToolTip1.SetToolTip(chkScriptToFile, "Files will be created in CSV folder")
		ToolTip1.SetToolTip(chkRec_id, "Cannot use BULK INSERT")
		'Bulk Insert

	End Sub

	Private Function GetBoolSetting(ByVal sKey As String, ByVal bDefault As Boolean) As Boolean
		Dim s As String = oAppSetting.GetSetting(sKey)
		If s = "1" Then Return True
		If s = "0" Then Return False
		Return bDefault
	End Function

	Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
		If System.IO.Directory.Exists(txtFolderPath.Text) = False Then
			txtFolderPath.Text = ""
		End If

		Windows.Forms.Application.DoEvents()
		SetTableGrid(False)

		If dgTables.Rows.Count > 0 Then
			Dim sSortedColumn As String = oAppSetting.GetSetting("SortedColumn")
			If sSortedColumn <> "" Then
				Dim sSortOrder As String = oAppSetting.GetSetting("SortOrder")
				If sSortOrder = "Ascending" Then
					dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Ascending)
				Else
					dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Descending)
				End If
			End If
		End If

	End Sub

	Private Function GetCsvDelimeter2() As String
		Dim sFieldTerm As String = selDelimiter.Text

		Select Case Trim(sFieldTerm)
			Case "Tab" : Return "\t"
			Case Else : Return CChar(sFieldTerm) '|
		End Select
	End Function

	Private Function GetCsvDelimeter() As Char
		Dim sFieldTerm As String = selDelimiter.Text

		Select Case Trim(sFieldTerm)
			Case "Tab" : Return CChar(vbTab)
				'Case "c", "" : Return CChar(",")
			Case Else : Return CChar(sFieldTerm) '|
		End Select
	End Function

	Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

		Dim oHash As New Hashtable
		oHash("FolderPath") = txtFolderPath.Text
		oHash("ConnectTo") = txtConnectTo.Text
		oHash("Delimiter") = selDelimiter.SelectedText

		oHash("chkCreateTable") = IIf(chkCreateTable.Checked, "1", "0").ToString()
		oHash("chkDropTable") = IIf(chkDropTable.Checked, "1", "0").ToString()

		oHash("chkRec_id") = IIf(chkRec_id.Checked, "1", "0").ToString()
		oHash("chkDeleteData") = IIf(chkDeleteData.Checked, "1", "0").ToString()

		oHash("chkHideNotSelected") = IIf(chkHideNotSelected.Checked, "1", "0").ToString()
		oHash("chkShrinkTable") = IIf(chkShrinkTable.Checked, "1", "0").ToString()

		oHash("chkBulkInsert") = IIf(chkBulkInsert.Checked, "1", "0").ToString()
		oHash("chkScriptToFile") = IIf(chkScriptToFile.Checked, "1", "0").ToString()

		If Not dgTables.SortedColumn Is Nothing Then
			oHash("SortedColumn") = dgTables.SortedColumn.Name
			oHash("SortOrder") = dgTables.SortOrder.ToString()
		End If

		Dim oTables As List(Of String) = GetSelectedTables()
		Dim sTables As String = String.Join(",", oTables.ToArray)

		oHash("SelectedTables") = sTables

		oAppSetting.SaveSettings(oHash)
	End Sub

	Private Function GetSelectedTablesFromReg() As Hashtable

		Dim oSelTables As New Hashtable
		Dim sSelectedTables As String = oAppSetting.GetSetting("SelectedTables")
		If sSelectedTables <> "" Then
			Dim oSelectedTables As String() = Split(sSelectedTables, ",")
			For i As Integer = 0 To oSelectedTables.Length - 1
				Dim sTable As String = oSelectedTables(i)
				oSelTables(sTable) = True
			Next
		End If
		Return oSelTables
	End Function

	Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect1.Click

		FolderBrowserDialog1.SelectedPath = txtFolderPath.Text
		FolderBrowserDialog1.ShowDialog()

		If FolderBrowserDialog1.SelectedPath = "" Then
			dgTables.Rows.Clear()
			Exit Sub
		End If

		If FolderBrowserDialog1.SelectedPath = "" Then
			Exit Sub
		End If

		txtFolderPath.Text = FolderBrowserDialog1.SelectedPath

		SetTableGrid(True)
	End Sub

	Sub SetTableGrid(ByVal bRefresh As Boolean)

		If txtFolderPath.Text = "" Then
			Exit Sub
		End If

		Dim sSortOrder As String = ""
		Dim sSortedColumn As String = ""
		If bRefresh Then
			'Update oSelTables by selected tables from grid
			UpdateSelectedTables()

			If Not dgTables.SortedColumn Is Nothing Then
				sSortedColumn = dgTables.SortedColumn.Name
				sSortOrder = dgTables.SortOrder.ToString()
			End If
		End If

		Dim dStart As DateTime = Now
		Dim oSqlServerTables As System.Data.DataTable = GetSqlServerTables()

		If (dStart - Now).TotalSeconds > 10 Then
			Log("GetSqlServerTables " & (dStart - Now).TotalSeconds)
		End If

		dStart = Now

		Dim oTable As Data.DataTable = GetFilesTable(txtFolderPath.Text)

		If (dStart - Now).TotalSeconds > 10 Then
			Log("GetFilesTable " & (dStart - Now).TotalSeconds)
		End If

		If oTable Is Nothing Then
			Exit Sub
		End If

		'Update Checked, DestRowCount
		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim sTableName As String = oTable.Rows(iRow)("Name").ToString()

			oTable.Rows(iRow)("Checked") = oSelTables.ContainsKey(sTableName)

			If Not oSqlServerTables Is Nothing Then
				Dim oRows As Data.DataRow() = oSqlServerTables.Select("Name='" & sTableName & "'")
				If oRows.Length > 0 Then
					oTable.Rows(iRow)("DestRowCount") = oRows(0)("Rows")
				End If
			End If
		Next

		dgTables.DataSource = oTable
		dgTables.Update()

		Dim oCol As DataGridViewCheckBoxColumn = DirectCast(dgTables.Columns("Checked"), DataGridViewCheckBoxColumn)
		oCol.TrueValue = True
		oCol.SortMode = DataGridViewColumnSortMode.Automatic
		oCol.Width = 35
		oCol.HeaderText = ""

		dgTables.Columns("DestRowCount").Visible = Not oSqlServerTables Is Nothing

		UpdateDataColumn("DateModified", "", "Date Modified")
		UpdateDataColumn("Size", "#,#", "Size")
		UpdateDataColumn("RowCount", "#,#", "Src Row Count")
		UpdateDataColumn("DestRowCount", "#,#", "Dest Row Count")

		SetupBackground()

		If sSortedColumn <> "" Then
			If sSortOrder = "Ascending" Then
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Ascending)
			Else
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Descending)
			End If
		End If

	End Sub

	Private Sub UpdateSelectedTables()

		oSelTables = New Hashtable
		Dim oTables As List(Of String) = GetSelectedTables()

		For i As Integer = 0 To oTables.Count - 1
			Dim sTable As String = oTables(i).ToString
			oSelTables(sTable) = True
		Next
	End Sub

	Private Sub SetupBackground()
		For iRow = 0 To dgTables.RowCount - 1
			Dim sSrcCount As String = dgTables.Rows(iRow).Cells("RowCount").Value.ToString()
			Dim sDstCount As String = dgTables.Rows(iRow).Cells("DestRowCount").Value.ToString()
			If sSrcCount <> "" AndAlso sDstCount <> "" Then
				If CInt(sSrcCount) = CInt(sDstCount) Then
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightBlue
				Else
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightPink
				End If
			Else
				dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.White
			End If
		Next
	End Sub

	Private Sub CheckCompare(sCol1 As String, sCol2 As String)
		For iRow = 0 To dgTables.RowCount - 1
			Dim sColVal1 As String = dgTables.Rows(iRow).Cells(sCol1).Value.ToString()
			Dim sColVal2 As String = dgTables.Rows(iRow).Cells(sCol2).Value.ToString()
			If sColVal1 <> "" AndAlso sColVal2 <> "" AndAlso sColVal1 <> sColVal2 Then
				dgTables.Rows(iRow).Cells("Checked").Value = True
			Else
				dgTables.Rows(iRow).Cells("Checked").Value = False
			End If
		Next
	End Sub

	Private Function GetFilesTable(ByVal sFolderPath As String) As Data.DataTable
		If chkHideNotSelected.Checked = False Then
			Return GetFilesTable2(sFolderPath)
		End If

		Dim oTable As Data.DataTable = GetFilesTable2(sFolderPath)
		For i As Integer = oTable.Rows.Count - 1 To 0 Step -1
			Dim sTableName As String = oTable.Rows(i)("Name").ToString()
			If oSelTables.ContainsKey(sTableName) = False Then
				oTable.Rows(i).Delete()
			End If
		Next

		Return oTable
	End Function

	Private Function GetFilesTable2(ByVal sFolderPath As String) As Data.DataTable

		'Try to get list if files from cache
		Dim sTempFilePath As String = GetTempFileName(sFolderPath, "CsvCopyXml")
		Dim ds As New System.Data.DataSet()
		If IO.File.Exists(sTempFilePath) Then
			Dim oFileInfo As New IO.FileInfo(sTempFilePath)
			If DateTime.Now.Subtract(oFileInfo.LastWriteTime).Hours > 2 Then
				'File is 2 hours old - delete
				System.IO.File.Delete(sTempFilePath)
			Else
				btnDeleteFolderCache.Visible = True
				txtFolderPath.Width = txtConnectTo.Width - 100
				ds.ReadXml(sTempFilePath)
				Return ds.Tables(0)
			End If
		End If

		txtFolderPath.Width = txtConnectTo.Width
		btnDeleteFolderCache.Visible = False

		Dim dStart As DateTime = DateTime.Now

		Dim oTable As New Data.DataTable
		oTable.Columns.Add(New Data.DataColumn("Checked", System.Type.GetType("System.Boolean"))) '<--
		oTable.Columns.Add(New Data.DataColumn("Name"))
		oTable.Columns.Add(New Data.DataColumn("DateModified", System.Type.GetType("System.DateTime")))
		oTable.Columns.Add(New Data.DataColumn("Size", System.Type.GetType("System.Int64")))
		oTable.Columns.Add(New Data.DataColumn("RowCount", System.Type.GetType("System.Int64")))
		oTable.Columns.Add(New Data.DataColumn("DestRowCount", System.Type.GetType("System.Int64"))) '<--

		Dim oFiles As String()

		Try
			oFiles = System.IO.Directory.GetFiles(sFolderPath)
		Catch ex As Exception
			MsgBox(ex.Message)
			Return Nothing
		End Try

		For i As Integer = 0 To oFiles.Length - 1
			Dim sFilePath As String = oFiles(i)
			Dim oFileInfo As New IO.FileInfo(sFilePath)
			Dim sTableName As String = IO.Path.GetFileNameWithoutExtension(sFilePath)
			If oFileInfo.Extension.ToLower() = ".csv" Then
				Dim iRowCount As Integer = GetRecCount(sFolderPath, sTableName)

				If iRowCount > 0 Then
					'remove header row from the row count
					iRowCount += -1
				End If

				Dim oDataRow As DataRow = oTable.NewRow()
				oDataRow("Name") = sTableName
				oDataRow("DateModified") = oFileInfo.LastWriteTime
				oDataRow("Size") = oFileInfo.Length

				If iRowCount > 0 Then
					oDataRow("RowCount") = iRowCount
				End If

				oTable.Rows.Add(oDataRow)
			End If
		Next

		Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
		If oDuration.Seconds > 5 Then
			'Save to Cache if it took more than 5 seconds to get the list of tables
			ds.Tables.Add(oTable)
			ds.WriteXml(sTempFilePath, XmlWriteMode.WriteSchema)

			btnDeleteFolderCache.Visible = True
			txtFolderPath.Width = txtConnectTo.Width - 100
		End If

		Return oTable
	End Function

	Private oRowCount As New Hashtable()

	Private Function GetRecCount(ByVal sFolderPath As String, ByVal sTableName As String) As Integer
		Dim sFilePath As String = IO.Path.Combine(sFolderPath, sTableName & ".csv")
		Return CountLinesInFile(sFilePath)
	End Function

	Public Function CountLinesInFile(sFilePath As String) As Integer

		If IO.File.Exists(sFilePath) = False Then
			Return 0
		End If

		Dim i As Integer = 0
		'Dim oTextFieldParser As New Microsoft.VisualBasic.FileIO.TextFieldParser(sFilePath)
		'oTextFieldParser.TextFieldType = FileIO.FieldType.Delimited
		'oTextFieldParser.SetDelimiters(GetCsvDelimeter())

		'While Not oTextFieldParser.EndOfData
		'	i += 1
		'	Dim oFields As String() = oTextFieldParser.ReadFields()
		'End While
		'oTextFieldParser.Close()

		Using reader As New System.IO.StreamReader(sFilePath)
			While reader.ReadLine() IsNot Nothing
				i += 1
			End While
		End Using

		Return i
	End Function

	Private Function GetTempFileName(ByVal sKey As String, sExt As String) As String
		Dim oRegex As New Regex(String.Format("[{0}]", Regex.Escape(New String(IO.Path.GetInvalidFileNameChars()))), RegexOptions.Compiled)
		Dim sFileName As String = oRegex.Replace(sKey, "-") & "." & sExt
		Return IO.Path.Combine(GetTempFolderPath(), sFileName)
	End Function

	Private Function GetSqlServerTables() As Data.DataTable

		If txtConnectTo.Text = "" Then
			Return Nothing
		End If

		Dim cn As SqlConnection = New SqlConnection(txtConnectTo.Text)

		Try
			cn.Open()
		Catch ex As Exception
			MsgBox(ex.Message)
			Return Nothing
		End Try

		Dim sSql As String = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
		Dim oTable As Data.DataTable = GetTable(cn, sSql)

		Dim oRetTable As New Data.DataTable
		oRetTable.Columns.Add(New Data.DataColumn("Name"))
		oRetTable.Columns.Add("Rows", System.Type.GetType("System.Int64"))

		For i As Long = 0 To oTable.Rows.Count - 1
			Dim sSchema As String = oTable.Rows(i)("TABLE_SCHEMA") & ""
			Dim sTName As String = oTable.Rows(i)("TABLE_NAME") & ""
			Dim sKey As String = sSchema & "." & sTName

			If sSchema <> "sys" Then
				If sSchema = "" Or sSchema = "dbo" Then
					sKey = sTName
				End If

				Try
					Dim cmd As New SqlCommand("sp_MStablespace '" & sKey & "'", cn)
					Dim dr As SqlDataReader = cmd.ExecuteReader()
					If dr.Read Then
						Dim iRowCount As Integer = CInt(dr.GetValue(dr.GetOrdinal("Rows")))
						If iRowCount > 0 Then
							Dim oDataRow As DataRow = oRetTable.NewRow()
							oDataRow("Name") = sKey
							oDataRow("Rows") = iRowCount
							oRetTable.Rows.Add(oDataRow)
						End If
					End If
					dr.Close()
				Catch ex As Exception
					'Do Nothing
				End Try
			End If

		Next

		cn.Close()
		Return oRetTable
	End Function

	Private Function GetFolderPath() As String
		Return txtFolderPath.Text
	End Function

	Private Sub UpdateDataColumn(sColName As String, sFormat As String, sHeaderText As String)
		Dim oCol As DataGridViewColumn = dgTables.Columns(sColName)
		If sFormat <> "" Then oCol.DefaultCellStyle.Format = sFormat
		If sHeaderText <> "" Then oCol.HeaderText = sHeaderText
	End Sub

	Protected Function EditConnectionString(ByVal sConnectionString As String) As String
		Try
			Dim oDataLinks As Object = CreateObject("DataLinks")
			Dim cn As Object = CreateObject("ADODB.Connection")

			cn.ConnectionString = sConnectionString
			oDataLinks.hWnd = Me.Handle

			If Not oDataLinks.PromptEdit(cn) Then
				'User pressed cancel button
				Return ""
			End If

			cn.Open()

			Return cn.ConnectionString

		Catch ex As Exception
			MsgBox(ex.Message)
			Return ""
		End Try
	End Function

	Function GetSelectedTables() As List(Of String)
		Dim oRet As New List(Of String)

		For Each oRow As DataGridViewRow In dgTables.Rows
			Dim oCheckbox As DataGridViewCheckBoxCell = DirectCast(oRow.Cells.Item(0), DataGridViewCheckBoxCell)

			If oCheckbox.Value.ToString = oCheckbox.TrueValue.ToString() Then
				Dim sName As String = oRow.Cells(1).Value.ToString()
				oRet.Add(sName)
			End If
		Next

		Return oRet
	End Function


	Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click

		If dgTables.Rows.Count = 0 Then
			MsgBox("Please connect to the source database.")
			Exit Sub
		End If

		Dim oTables As List(Of String) = GetSelectedTables()
		If oTables.Count = 0 Then
			MsgBox("Please select tables to copy.")
			Exit Sub
		End If

		If oTables.Count = 0 Then
			Exit Sub
		End If

		Dim cn As SqlConnection = New SqlConnection(txtConnectTo.Text)

		Try
			cn.Open()
		Catch ex As Exception
			MsgBox(ex.Message)
			Exit Sub
		End Try

		ProgressBar2.Maximum = oTables.Count
		bStop = False

		For i As Integer = 0 To oTables.Count - 1
			ProgressBar2.Value = i
			ProgressBar2.Refresh()
			Windows.Forms.Application.DoEvents()

			Dim sTable As String = oTables(i).ToString
			If Len(sTable) > 128 Then
				Log("Table name " & sTable & " is " & Len(sTable) & " characters long. It can be 128 max.")
			Else
				CopyTable(sTable, cn)

				If chkShrinkTable.Checked AndAlso chkScriptToFile.Checked = False Then
					Dim dStart As DateTime = DateTime.Now
					Try
						Shrink(cn, sTable)
						Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
						Log("Shrunk table " & sTable & " in " & oDuration.Seconds & " seconds")
					Catch ex As Exception
						Log("Could not shrink table " & sTable & ". " & ex.Message)
					End Try
				End If

				If bStop Then
					Exit For
				End If
			End If
		Next

		ProgressBar2.Value = 0
		cn.Close()
		SetTableGrid(True)
	End Sub

	Private Sub btnShrinkTables_Click(sender As Object, e As EventArgs) Handles btnShrinkTables.Click

		Dim cn As SqlConnection = New SqlConnection(txtConnectTo.Text)

		Try
			cn.Open()
		Catch ex As Exception
			MsgBox(ex.Message)
			Exit Sub
		End Try

		For iRow = 0 To dgTables.RowCount - 1
			Dim bChecked As Boolean = CBool(dgTables.Rows(iRow).Cells("Checked").Value)
			If bChecked Then
				Dim dStart As DateTime = DateTime.Now
				Dim sTable As String = dgTables.Rows(iRow).Cells("Name").Value.ToString()

				If chkScriptToFile.Checked Then
					'ExecuteCommand will log to file instead of executing the SQL

					Dim sScriptFilePath As String = System.IO.Path.Combine(txtFolderPath.Text, sTable & "_Shrink.sql")

					If IO.File.Exists(sScriptFilePath) Then
						IO.File.Delete(sScriptFilePath)
					End If

					oSqlScriptWriter = New System.IO.StreamWriter(sScriptFilePath)
				Else
					oSqlScriptWriter = Nothing
				End If

				Try
					Shrink(cn, sTable)
					Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
					Log("Shrunk table " & sTable & " in " & oDuration.Seconds & " seconds")
				Catch ex As Exception
					Log("Could not shrink table " & sTable & ". " & ex.Message)
				End Try

				If oSqlScriptWriter IsNot Nothing Then
					'Cleanup log file
					oSqlScriptWriter.Close()
					oSqlScriptWriter = Nothing
				End If

			End If
		Next

		cn.Close()
	End Sub

	Private Sub OpenConnections(ByRef cn As SqlConnection)
		If cn.State <> ConnectionState.Open Then
			cn.Open()
		End If
	End Sub

	Private Sub CopyTable(ByVal sTableName As String, ByRef cn As SqlConnection)

		Dim dStart As DateTime = DateTime.Now
		Dim bDestTableExists As Boolean = False
		Dim iDestRecCount As Integer = 0

		If chkScriptToFile.Checked Then
			'ExecuteCommand will log to file instead of executing the SQL

			Dim sScriptFilePath As String = System.IO.Path.Combine(txtFolderPath.Text, sTableName & ".sql")

			If IO.File.Exists(sScriptFilePath) Then
				IO.File.Delete(sScriptFilePath)
			End If

			oSqlScriptWriter = New System.IO.StreamWriter(sScriptFilePath)
		Else
			oSqlScriptWriter = Nothing
		End If

		Try
			Dim cm As New SqlCommand("SELECT Count(*) FROM " & PadColumnName(sTableName), cn)
			iDestRecCount = Integer.Parse(cm.ExecuteScalar().ToString())
			bDestTableExists = True
		Catch ex As Exception
			'Ignore - assume table dos not exist
		End Try

		Dim bDropTable As Boolean = chkCreateTable.Checked AndAlso chkDropTable.Checked AndAlso bDestTableExists

		If chkDeleteData.Checked AndAlso iDestRecCount > 0 AndAlso bDropTable = False Then
			Log("Deleteting data from table: " & sTableName)

			OpenConnections(cn)
			Dim sSql As String = "DELETE FROM " & PadColumnName(sTableName)

			Try
				ExecuteCommand(cn, sSql)
			Catch ex As Exception
				Log(ex.Message & vbTab & "SQL: " & sSql)
			End Try
		End If

		Dim sFolderPath As String = GetFolderPath()
		Dim sFilePath As String = IO.Path.Combine(sFolderPath, sTableName & ".csv")
		Dim iLineCount As Integer = GetRecCount(sFolderPath, sTableName)
		If iLineCount = 0 Then
			'Nothing to copy - Exit
			Exit Sub
		End If

		If bDropTable Then
			Log("Drop table: " & sTableName)

			Dim sSql As String = "DROP TABLE " & PadColumnName(sTableName)

			Try
				ExecuteCommand(cn, sSql)
				bDestTableExists = False
			Catch ex As Exception
				Log("Could not drop table: " & sTableName & ", " & ex.Message & vbTab)
			End Try
		End If

		Log("Copying " & iLineCount & " rows from table: " & sTableName)

		If chkBulkInsert.Checked Then
			BulkInsertFromFile(sFilePath, sTableName, cn, iLineCount, chkCreateTable.Checked And bDestTableExists = False)
		Else
			InsertFromFile(sFilePath, sTableName, cn, iLineCount, chkCreateTable.Checked And bDestTableExists = False)
		End If

		Log("Copied table " & sTableName & vbTab & " in " & GetDuration(dStart))

		If oSqlScriptWriter IsNot Nothing Then
			'Cleanup log file
			oSqlScriptWriter.Close()
			oSqlScriptWriter = Nothing
		End If

	End Sub

	Sub CreateTableFromDataTable(ByRef cn As SqlConnection, dt As DataTable, ByVal sTableName As String)
		Dim sCreateColumns As String = ""

		If chkRec_id.Checked Then
			sCreateColumns += "rec_id int not null primary key clustered identity(1,1)"
		End If

		For iCol = 0 To dt.Columns.Count - 1
			Dim sCol As String = dt.Columns(iCol).ColumnName
			If sCreateColumns <> "" Then sCreateColumns += ", "
			sCreateColumns += PadColumnName(sCol) & " " & GetNvarChar() & "(max) NULL"
		Next

		Dim sSql As String = "create table " & PadColumnName(sTableName) & " (" & sCreateColumns & ")"
		ExecuteCommand(cn, sSql)
	End Sub

	Private Sub BulkInsertFromFile(ByVal csvPath As String,
								   ByVal destinationTableName As String,
								   ByRef cn As SqlConnection,
								   ByVal iLineCount As Integer,
								   ByVal bCreateTable As Boolean,
								   Optional batchSize As Integer = 50000,
								   Optional bulkCopyTimeoutSeconds As Integer = 0,
								   Optional delimiter As String = ",",
								   Optional hasHeaders As Boolean = True,
								   Optional keepNullsAsDBNull As Boolean = True)


		ProgressBar1.Value = 0
		ProgressBar1.Maximum = iLineCount
		lbCount.Visible = False
		lbCount.Text = ""
		btnStop.Visible = False

		Dim swTotal As Stopwatch = Stopwatch.StartNew()
		Dim totalRead As Long = 0
		Dim totalInserted As Long = 0
		Dim batchNumber As Integer = 0

		Using parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(csvPath)
			parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
			parser.SetDelimiters(delimiter)
			parser.HasFieldsEnclosedInQuotes = True
			parser.TrimWhiteSpace = False

			Dim columnNames As String() = Nothing
			Dim firstRow As String() = Nothing

			If parser.EndOfData Then
				Exit Sub
			End If

			If hasHeaders Then
				columnNames = parser.ReadFields()
				If columnNames Is Nothing OrElse columnNames.Length = 0 Then
					Exit Sub
				End If
			Else
				firstRow = parser.ReadFields()
				If firstRow Is Nothing OrElse firstRow.Length = 0 Then
					Exit Sub
				End If

				ReDim columnNames(firstRow.Length - 1)
				For i = 0 To columnNames.Length - 1
					columnNames(i) = "Col" & (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
				Next
			End If

			' Prepare bulk copy once
			Using bulk As New SqlBulkCopy(cn, SqlBulkCopyOptions.TableLock, Nothing)
				bulk.DestinationTableName = destinationTableName
				bulk.BatchSize = batchSize
				If bulkCopyTimeoutSeconds > 0 Then bulk.BulkCopyTimeout = bulkCopyTimeoutSeconds

				' Column mappings by name (CSV header must match destination columns)
				For Each col In columnNames
					bulk.ColumnMappings.Add(col, col)
				Next

				' Now stream rows in batches
				Dim dt As DataTable = CreateBatchTable(columnNames)

				' If we already consumed the first row (no headers), add it first
				If firstRow IsNot Nothing Then
					AddRowToBatch(dt, firstRow, keepNullsAsDBNull)
				End If

				If bCreateTable Then
					CreateTableFromDataTable(cn, dt, destinationTableName)
				End If

				Dim swBatch As New Stopwatch()

				While Not parser.EndOfData
					Dim fields As String()

					Try
						fields = parser.ReadFields()
					Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
						' You can log ex.LineNumber / ex.Message and continue, or rethrow.
						Throw New Exception($"Malformed CSV at line {ex.LineNumber}.", ex)
					End Try

					If fields Is Nothing Then Continue While

					AddRowToBatch(dt, fields, keepNullsAsDBNull)

					totalRead += 1
					ProgressBar1.Value = CInt(totalRead)

					If dt.Rows.Count >= batchSize Then
						batchNumber += 1
						bulk.WriteToServer(dt)
						totalInserted += dt.Rows.Count

						swBatch.Stop()
						Log($"Batch {batchNumber:N0} | Rows: {dt.Rows.Count:N0} | " &
							$"Batch time: {swBatch.Elapsed.TotalSeconds:N2}s | " &
							$"Total rows: {totalInserted:N0}")
						swBatch.Restart()

						dt.Clear() ' keep schema, release rows
					End If
				End While

				' Flush remaining rows
				If dt.Rows.Count > 0 Then
					bulk.WriteToServer(dt)
					totalInserted += dt.Rows.Count
					dt.Clear()
				End If
			End Using
		End Using

		ProgressBar1.Value = 0
		lbCount.Visible = False
		lbCount.Text = ""
		btnStop.Visible = False

		swTotal.Stop()

		Debug.WriteLine(
			$"CSV load complete | File: {csvPath} | " &
			$"Rows inserted: {totalInserted:N0} | " &
			$"Total time: {swTotal.Elapsed.TotalMinutes:N2} minutes"
		)

	End Sub

	' Creates a DataTable for one batch (all string columns by default)
	Private Function CreateBatchTable(columnNames As String()) As DataTable
		Dim dt As New DataTable()
		For Each Name As String In columnNames
			Dim colName = If(String.IsNullOrWhiteSpace(Name), "UnnamedColumn", Name.Trim())
			' Defaulting to String keeps it fast and avoids type conversion memory overhead.
			dt.Columns.Add(colName, GetType(String))
		Next
		Return dt
	End Function

	' Adds one CSV record into the batch table (pads/trim fields to match column count)
	Private Sub AddRowToBatch(dt As DataTable, fields As String(), keepNullsAsDBNull As Boolean)
		Dim colCount = dt.Columns.Count
		Dim row = dt.NewRow()

		For i = 0 To colCount - 1
			Dim value As String = Nothing
			If i < fields.Length Then value = fields(i)

			If keepNullsAsDBNull AndAlso (value Is Nothing OrElse value.Length = 0) Then
				row(i) = DBNull.Value
			Else
				row(i) = value
			End If
		Next

		dt.Rows.Add(row)
	End Sub

	Private Sub InsertFromFile(ByVal sFilePath As String,
							   ByVal sTableName As String,
							   ByRef cn As SqlConnection,
							   ByVal iLineCount As Integer,
							   ByVal bCreateTable As Boolean)

		ProgressBar1.Maximum = iLineCount
		lbCount.Visible = True
		btnStop.Visible = True

		Dim swTotal As Stopwatch = Stopwatch.StartNew()
		Dim sCreateColumns As String = ""

		If chkRec_id.Checked Then
			sCreateColumns += "rec_id int not null primary key clustered identity(1,1)"
		End If

		Dim sSqlHeader As String = ""
		Dim oColumns As New Hashtable
		Dim iRow As Long = 0
		Dim oTextFieldParser As New Microsoft.VisualBasic.FileIO.TextFieldParser(sFilePath)
		oTextFieldParser.TextFieldType = FileIO.FieldType.Delimited
		oTextFieldParser.SetDelimiters(GetCsvDelimeter())

		While Not oTextFieldParser.EndOfData
			iRow += 1

			Dim sValues As String = ""

			Dim oFields As String() = oTextFieldParser.ReadFields()
			For iCol As Integer = 0 To oFields.Length - 1
				Dim sCol As String = Trim(oFields(iCol) & "")

				If iRow = 1 Then
					oColumns(iCol) = sCol
					If sSqlHeader <> "" Then sSqlHeader += ", "
					sSqlHeader += PadColumnName(sCol)

					If sCreateColumns <> "" Then sCreateColumns += ", "
					sCreateColumns += PadColumnName(sCol) & " " & GetNvarChar() & "(max) NULL"
				Else
					If iRow = 2 AndAlso iCol = 0 Then

						If bCreateTable Then
							Dim sSql As String = "create table " & PadColumnName(sTableName) & " (" & sCreateColumns & ")"
							ExecuteCommand(cn, sSql)
						End If
					End If

					If oColumns.ContainsKey(iCol) Then
						'Dim sColumn As String = oColumns(iCol).ToString()
						If sValues <> "" Then sValues += ", "
						If Trim(sCol) = "" Then
							sValues += "null"
						Else
							If chkNvarChar.Checked Then
								sValues += "N"
							End If

							sValues += "'" & (sCol & "").Replace("'", "''") & "'"
						End If

					End If
				End If
			Next

			If iRow > 1 Then
				Dim sSql As String = "insert into [" & sTableName & "] (" & sSqlHeader & ")" &
					" values (" & sValues & ")"
				ExecuteCommand(cn, sSql)

				If bStop Then
					Exit While
				End If
			End If

			ProgressBar1.Value = CInt(iRow)
			lbCount.Text = iRow.ToString()
			lbCount.Refresh()
			Windows.Forms.Application.DoEvents()
		End While

		oTextFieldParser.Close()

		ProgressBar1.Value = 0
		lbCount.Visible = False
		lbCount.Text = ""
		btnStop.Visible = False

		swTotal.Stop()

		Debug.WriteLine(
			$"CSV load complete | File: {sFilePath} | " &
			$"Rows inserted: {iRow:N0} | " &
			$"Total time: {swTotal.Elapsed.TotalMinutes:N2} minutes"
		)
	End Sub



	Function GetNvarChar() As String

		If chkNvarChar.Checked Then
			Return "nvarchar"
		Else
			Return "varchar"
		End If

	End Function

	Private Function CreateDataTable(ByRef oTable As System.Data.DataTable,
									 ByVal sTableName As String,
									 ByRef cn As SqlConnection) As String

		If oTable Is Nothing Then
			Return ""
		End If

		Dim sCreateColumns As String = ""

		If chkRec_id.Checked Then
			sCreateColumns += "rec_id int not null primary key clustered identity(1,1)"
		End If

		For i As Integer = 0 To oTable.Columns.Count - 1
			Dim sCol As String = oTable.Columns(i).ColumnName
			If sCreateColumns <> "" Then sCreateColumns += ", "
			sCreateColumns += PadColumnName(sCol) & " " & GetNvarChar() & "(max) NULL"
		Next

		Dim sSql As String = "create table " & PadColumnName(sTableName) & " (" & sCreateColumns & ")"
		ExecuteCommand(cn, sSql)

		Return ""
	End Function
	Private Function GetDuration(ByVal dStart As DateTime) As String
		Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
		Return (New DateTime(oDuration.Ticks)).ToString("HH 'hrs' mm 'mins' ss 'secs'").Replace("00 hrs", "").Replace("00 mins", "").Trim()
	End Function

	Private Function ContainsClosingDoubleQuote(ByVal oLine As String()) As Boolean
		For i As Integer = 0 To oLine.Length - 1
			Dim sVal As String = Trim(oLine(i) & "")
			If Microsoft.VisualBasic.Left(sVal, 1) <> """" AndAlso Microsoft.VisualBasic.Right(sVal, 1) = """" Then
				Return True
			End If
		Next
		Return False
	End Function

	Private Sub Log(s As String)

		If txtLog.Text = "" Then
			txtLog.Text = s
		Else
			txtLog.AppendText(vbCrLf & s)
		End If

		txtLog.Visible = True
		txtLog.ScrollToCaret()
		txtLog.Refresh()
	End Sub


	Private Function GetTempFolderPath() As String
		Dim sFolder As String = Application.StartupPath()
		Dim sXmlFolder As String = System.IO.Path.Combine(sFolder, "CsvCopyXml")

		If Not System.IO.Directory.Exists(sXmlFolder) Then
			System.IO.Directory.CreateDirectory(sXmlFolder)
		End If

		Return sXmlFolder
	End Function

	Private Sub btnConnect2_Click(sender As Object, e As EventArgs) Handles btnConnect2.Click
		Dim sConnectionString As String = txtConnectTo.Text

		If sConnectionString = "" Then
			sConnectionString = "Provider=SQLOLEDB.1"

		ElseIf sConnectionString.ToLower().IndexOf("provider=") = -1 Then
			sConnectionString = "Provider=SQLOLEDB.1;" & sConnectionString
		End If

		sConnectionString = EditConnectionString(sConnectionString)

		If sConnectionString = "" Then
			Exit Sub
		End If

		txtConnectTo.Text = RemoveProvider(sConnectionString)

		SetTableGrid(True)
	End Sub

	Private Function RemoveProvider(ByVal s As String) As String
		Dim oList As String() = Split(s, ";")
		Dim sRet As String = ""

		For i As Integer = 0 To oList.Length - 1
			Dim sItem As String = oList(i)
			Dim oItem As String() = Split(sItem, "=")
			Dim sKey As String = LCase(oItem(0))
			If "data source" = sKey Or
			   "initial catalog" = sKey Or
			   "user id" = sKey Or
			   "integrated security" = sKey Or
			   "password" = sKey Then

				sRet += sItem & ";"
			End If
		Next

		Return sRet
	End Function

	Private Sub btnCheckAll_Click(sender As Object, e As EventArgs) Handles btnCheckAll.Click
		For iRow = 0 To dgTables.RowCount - 1
			dgTables.Rows(iRow).Cells("Checked").Value = True
		Next
	End Sub

	Private Sub btnUncheckAll_Click(sender As Object, e As EventArgs) Handles btnUncheckAll.Click
		For iRow = dgTables.RowCount - 1 To 0 Step -1
			dgTables.Rows(iRow).Cells("Checked").Value = False
		Next
	End Sub

	Private Sub btnCheckNew_Click(sender As Object, e As EventArgs) Handles btnCheckNew.Click
		CheckCompare("DestRowCount", "RowCount")
	End Sub

	Private Sub dgTables_Sorted(sender As Object, e As EventArgs) Handles dgTables.Sorted
		SetupBackground()
	End Sub

	Private Sub txtFolderPath_KeyUp(sender As Object, e As KeyEventArgs) Handles txtFolderPath.KeyUp
		If e.KeyCode = Keys.Enter Then
			SetTableGrid(True)
		End If
	End Sub

	Private Sub txtConnectTo_KeyUp(sender As Object, e As KeyEventArgs) Handles txtConnectTo.KeyUp
		If e.KeyCode = Keys.Enter Then
			SetTableGrid(True)
		End If
	End Sub

	Private Sub chkCreateTable_CheckedChanged(sender As Object, e As EventArgs) Handles chkCreateTable.CheckedChanged
		chkCreateTable_CheckedChanged()
	End Sub

	Private Sub chkCreateTable_CheckedChanged()
		chkDropTable.Visible = chkCreateTable.Checked
		chkRec_id.Visible = chkCreateTable.Checked
		chkNvarChar.Visible = chkCreateTable.Checked
	End Sub

	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub

	Private Sub btnStop_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles btnStop.LinkClicked
		bStop = True
	End Sub

	Private Sub chkCopyFiles_CheckedChanged(sender As Object, e As EventArgs)
		SetTableGrid(True)
	End Sub

	Private Sub chkHideSelected_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideNotSelected.CheckedChanged
		SetTableGrid(True)
	End Sub

	Private Sub btnDeleteFolderCache_Click(sender As Object, e As EventArgs) Handles btnDeleteFolderCache.Click
		Dim sTempFilePath As String = GetTempFileName(txtFolderPath.Text, "CsvCopyXml")

		If IO.File.Exists(sTempFilePath) = False Then
			Exit Sub
		End If

		If MsgBox("Delete cache? Did you add a file to the folder within two hours? The cache file is used for two hours if it takes 5 seconds or more to get the list of tables.", vbYesNo) = vbNo Then
			Exit Sub
		End If

		IO.File.Delete(sTempFilePath)
		Windows.Forms.Application.DoEvents()
		SetTableGrid(False)

		txtFolderPath.Width = txtConnectTo.Width
		btnDeleteFolderCache.Visible = False
	End Sub

	Friend Sub Shrink(ByRef cn As SqlConnection, ByVal sTable As String)
		Dim oTable As DataTable = GetColumnsDataTable(cn, sTable)

		If oTable.Rows.Count = 0 Then
			Exit Sub
		End If

		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim oRow As DataRow = oTable.Rows(iRow)
			Dim sName As String = oRow("Name") & ""
			Dim sDataType As String = oRow("DataType") & ""
			Dim sColumnSize As String = oRow("ColumnSize") & ""
			Dim sNumeric As String = oRow("Numeric") & ""
			Dim sPeriod As String = oRow("Period") & ""
			Dim sNull As String = oRow("Null") & ""
			Dim sStartsWithZero As String = oRow("StartsWithZero") & ""
			Dim sDate As String = oRow("Date") & ""

			Dim sLength As String = oRow("Length") & ""
			Dim iLength As Integer = 0
			If sLength <> "" Then
				iLength = CInt(sLength)
			End If

			Dim sNewDataType As String = ""

			If sDate = "Y" Then
				If iLength > 10 Then
					sNewDataType = "datetime"
				Else
					sNewDataType = "date"
				End If

			ElseIf sNumeric = "Y" AndAlso ((sPeriod <> "Y" AndAlso sStartsWithZero = "Y") = False) Then

				If sPeriod = "Y" Then
					sNewDataType = "decimal(10,2)"
				Else
					If iLength > 9 Then
						sNewDataType = "bigint"
					Else
						sNewDataType = "int"
					End If

				End If

			ElseIf sLength <> "" AndAlso sColumnSize <> sLength Then
				sNewDataType = sDataType & "(" & RoundUp(sLength) & ")"
			End If

			If sNewDataType <> "" Then
				Dim sSql As String = "alter table " & PadTableName(sTable) & " alter column [" & sName & "] " & sNewDataType
				Dim sError As String = ExecuteCommand(cn, sSql)
				If sError <> "" Then
					If sLength <> "" AndAlso sColumnSize <> sLength Then
						sNewDataType = sDataType & "(" & RoundUp(sLength) & ")"
						sSql = "alter table " & PadTableName(sTable) & " alter column [" & sName & "] " & sNewDataType
						sError = ExecuteCommand(cn, sSql)
					End If
				End If

				If sError <> "" Then
					Log(sSql & vbCrLf & vbTab & sError)
				End If
			End If
		Next

	End Sub

	Function GetTable(ByRef cn As SqlConnection, ByVal sSql As String) As System.Data.DataTable
		Dim ds As DataSet = New DataSet
		Dim ad As New SqlDataAdapter(sSql, cn)
		ad.Fill(ds)
		Return ds.Tables(0)
	End Function

	Public Function PadColumnName(ByVal s As String) As String
		Return "[" & s & "]"
	End Function

	Function PadTableName(ByVal s As String) As String
		Return "[" & s & "]"
	End Function

	Function PadQuotes(ByVal s As String) As String
		If s = "" Then
			Return ""
		End If
		Return (s & "").Replace("'", "''")
	End Function

	Friend Function GetColumnsDataTable(ByRef cn As SqlConnection, ByVal sTableName As String) As DataTable
		Dim oDataTable As New DataTable

		oDataTable.Columns.Add(New DataColumn("Name", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("DataType", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("ColumnSize", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("Length", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("Numeric", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("Date", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("Period", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("Null", System.Type.[GetType]("System.String")))
		oDataTable.Columns.Add(New DataColumn("StartsWithZero", System.Type.[GetType]("System.String")))

		If Trim(sTableName) = "" Then
			Return oDataTable
		End If

		Dim sSql As String = ""

		If sTableName.IndexOf(".") <> -1 Then
			sSql = "select * from INFORMATION_SCHEMA.COLUMNS where TABLE_SCHEMA + '.' + TABLE_NAME = '" & PadQuotes(sTableName) & "'"
		Else
			sSql = "select * from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & PadQuotes(sTableName) & "' order by ordinal_position"
		End If

		Dim oComputedColumns As Hashtable = GetComputedColumns(cn, sTableName)
		Dim oTable As DataTable = GetTable(cn, sSql)

		Dim oColumnNulls As Hashtable = GetColumnFunc(cn, sTableName, oTable, "NULL", CombineHash(oComputedColumns))
		Dim oExcludeCols As Hashtable = CombineHash(oComputedColumns, oColumnNulls) 'exlude Computed and Null Columns

		Dim oColumnLength As Hashtable = GetColumnLength(cn, sTableName, oTable, oExcludeCols)
		Dim oColumnNumeric As Hashtable = GetColumnFunc(cn, sTableName, oTable, "ISNUMERIC", oExcludeCols)
		Dim oColumnDate As Hashtable = GetColumnFunc(cn, sTableName, oTable, "ISDATE", CombineHash(oExcludeCols, oColumnNumeric))
		Dim oColumnPeriod As Hashtable = GetColumnFunc(cn, sTableName, oTable, "Period", Nothing, oColumnNumeric)
		Dim oColumnStartsWith0 As Hashtable = GetColumnFunc(cn, sTableName, oTable, "StartsWith0", Nothing, oColumnNumeric)

		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim sColumn As String = oTable.Rows(iRow)("COLUMN_NAME")

			If oComputedColumns.ContainsKey(sColumn) = False Then
				Dim sDataType As String = oTable.Rows(iRow)("DATA_TYPE") & ""
				'Dim bAllowDBNull As Boolean = oTable.Rows(iRow)("IS_NULLABLE") & "" = "YES"
				Dim sColumnSize As String = oTable.Rows(iRow)("CHARACTER_MAXIMUM_LENGTH") & ""

				If sDataType = "decimal" OrElse sDataType = "numeric" Then
					Dim sPrecision As String = oTable.Rows(iRow)("NUMERIC_PRECISION") & ""
					Dim sScale As String = oTable.Rows(iRow)("NUMERIC_SCALE") & ""
					sDataType += " (" & sPrecision & ", " & sScale & ")"

				ElseIf sDataType = "text" OrElse sDataType = "image" Then
					sColumnSize = ""

				ElseIf sColumnSize = "-1" Then
					sColumnSize = "max"
				End If

				Dim sCol As String = sColumn.ToLower()

				Dim oDataRow As DataRow = oDataTable.NewRow()


				oDataRow("Name") = sColumn
				oDataRow("DataType") = sDataType
				oDataRow("ColumnSize") = sColumnSize
				oDataRow("Length") = oColumnLength(sCol)
				oDataRow("Numeric") = oColumnNumeric(sCol)
				oDataRow("Date") = oColumnDate(sCol)
				oDataRow("Period") = oColumnPeriod(sCol)
				oDataRow("Null") = oColumnNulls(sCol)
				oDataRow("StartsWithZero") = oColumnStartsWith0(sCol)

				oDataTable.Rows.Add(oDataRow)
			End If
		Next

		Return oDataTable
	End Function

	Private Function CombineHash(ByVal a As Hashtable, ByVal b As Hashtable) As Hashtable

		For Each o As DictionaryEntry In b
			Dim sKey As String = o.Key
			a(sKey.ToLower()) = o.Value
		Next

		Return a
	End Function

	Private Function CombineHash(ByVal a As Hashtable) As Hashtable

		Dim oRet As New Hashtable

		For Each o As DictionaryEntry In a
			Dim sKey As String = o.Key
			oRet(sKey.ToLower()) = o.Value
		Next

		Return oRet
	End Function

	Private Function GetColumnLength(ByRef cn As SqlConnection,
									 ByVal sTableName As String,
									 ByRef oTable As DataTable,
									 ByRef oExcludeColumns As Hashtable) As Hashtable

		Dim sColumns As String = ""

		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim sColumn As String = oTable.Rows(iRow)("COLUMN_NAME")
			If oExcludeColumns.ContainsKey(sColumn.ToLower()) = False Then
				Dim sDataType As String = LCase(oTable.Rows(iRow)("DATA_TYPE") & "")
				If sDataType = "nvarchar" OrElse sDataType = "varchar" OrElse sDataType = "char" Then
					If sColumns <> "" Then sColumns += ", "
					sColumns += "max(len([" & sColumn & "])) as [" & sColumn & "]"
				End If
			End If
		Next

		Dim oColumns As New Hashtable

		If sColumns = "" Then
			Return oColumns
		End If

		Dim sColSql As String = "select " & sColumns & " from " & PadTableName(sTableName)
		Dim tdColumns As DataTable = GetTable(cn, sColSql)
		If tdColumns.Rows.Count > 0 Then
			For iCol As Integer = 0 To tdColumns.Columns.Count - 1
				Dim sCol As String = tdColumns.Columns(iCol).ColumnName.ToLower()
				Dim sVal As String = tdColumns.Rows(0)(iCol) & ""
				oColumns(sCol) = sVal
			Next
		End If

		Return oColumns
	End Function

	Private Function GetColumnFunc(ByRef cn As SqlConnection,
								   ByVal sTableName As String,
								   ByRef oTable As DataTable,
								   ByVal sFunc As String,
								   ByVal oExcludeColumns As Hashtable,
								   Optional ByVal oIncludeColumns As Hashtable = Nothing) As Hashtable

		If oIncludeColumns IsNot Nothing AndAlso oIncludeColumns.Count = 0 Then
			Return New Hashtable
		End If

		Dim sColumns As String = ""

		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim sColumn As String = oTable.Rows(iRow)("COLUMN_NAME") & ""

			If (oExcludeColumns IsNot Nothing AndAlso oExcludeColumns.ContainsKey(sColumn.ToLower()) = False) OrElse
			   (oIncludeColumns IsNot Nothing AndAlso oIncludeColumns.ContainsKey(sColumn.ToLower()) = True) Then

				Dim sDataType As String = LCase(oTable.Rows(iRow)("DATA_TYPE") & "")
				If sDataType = "nvarchar" OrElse sDataType = "varchar" OrElse sDataType = "char" Then
					If sColumns <> "" Then sColumns += ", "

					If sFunc = "Period" Then
						'check if any record has period
						sColumns += "max(case when CHARINDEX('.', [" & sColumn & "]) <> 0 then 1 else 0 end) as [" & sColumn & "]"

					ElseIf sFunc = "StartsWith0" Then
						'at least one record starts with zero
						sColumns += "max(case when SUBSTRING(isnull([" & sColumn & "],''), 1, 1) = '0' then 1 else 0 end) as [" & sColumn & "]"

					ElseIf sFunc = "NULL" Then
						'all rows are null
						sColumns += "min(case when [" & sColumn & "] IS NULL then 1 else 0 end) as [" & sColumn & "]"

					Else
						'make sure all rows are nulls or numeric
						sColumns += "min(case when [" & sColumn & "] IS NULL OR " & sFunc & "([" & sColumn & "]) = 1 then 1 else 0 end) as [" & sColumn & "]"
					End If

				End If
			End If
		Next

		Dim oColumns As New Hashtable

		If sColumns = "" Then
			Return oColumns
		End If

		Dim sColSql As String = "select " & sColumns & " from " & PadTableName(sTableName)

		'sMsg += "<p>" & sColSql & "</p>" & vbCrLf

		Dim tdColumns As DataTable = GetTable(cn, sColSql)
		If tdColumns.Rows.Count > 0 Then
			For iCol As Integer = 0 To tdColumns.Columns.Count - 1
				Dim sCol As String = tdColumns.Columns(iCol).ColumnName.ToLower()
				Dim sVal As String = tdColumns.Rows(0)(iCol) & ""
				If sVal = "1" Then
					oColumns(sCol) = "Y"
				End If
			Next
		End If

		Return oColumns
	End Function

	Protected Function GetComputedColumns(ByRef cn As SqlConnection,
										  ByVal sTableName As String,
										  Optional ByVal sColumnName As String = "") As Hashtable
		Dim oRet As New Hashtable
		Dim sSql As String = "SELECT name, definition FROM sys.computed_columns "

		If sTableName.IndexOf(".") = -1 Then
			sSql += " WHERE OBJECT_NAME(object_id) = '" & PadQuotes(sTableName) & "'"
		Else
			sSql += " WHERE OBJECT_SCHEMA_NAME(object_id) + '.' + OBJECT_NAME(object_id) = '" & PadQuotes(sTableName) & "'"
		End If

		If sColumnName <> "" Then
			sSql += " AND name = '" & PadQuotes(sColumnName) & "'"
		End If

		Try
			Dim oTable As DataTable = GetTable(cn, sSql)
			For iRow As Integer = 0 To oTable.Rows.Count - 1
				Dim sColumn As String = oTable.Rows(iRow)("name") & ""
				Dim sDef As String = oTable.Rows(iRow)("definition") & ""
				oRet(sColumn) = sDef
			Next

		Catch ex As Exception
			'Ignore becase of earlier versions of SQL Server that did not suppot computed columns
		End Try

		Return oRet
	End Function

	Private Function ExecuteCommand(ByRef cn As SqlConnection, ByVal sSql As String) As String

		If oSqlScriptWriter IsNot Nothing Then
			oSqlScriptWriter.WriteLine(sSql)
		Else
			Dim cm As New SqlCommand(sSql, cn)
			Try
				cm.ExecuteNonQuery()
			Catch ex As Exception
				Return ex.Message
			End Try
		End If

		Return ""
	End Function

	Private Function ExecuteCommand(ByRef cn As SqlConnection, ByVal sSql As String, iCommandTimeout As Integer) As String

		If oSqlScriptWriter IsNot Nothing Then
			oSqlScriptWriter.WriteLine(sSql)
		Else
			Dim cm As New SqlCommand(sSql, cn)
			cm.CommandTimeout = iCommandTimeout
			Try
				cm.ExecuteNonQuery()
			Catch ex As Exception
				Return ex.Message
			End Try
		End If

		Return ""
	End Function

	Private Function RoundUp(ByVal s As String) As String
		Dim i As Integer = RoundUp2(s)

		If chkNvarChar.Checked Then
			If i > 4000 Then
				Return "max"
			End If
		Else
			If i > 8000 Then
				Return "max"
			End If
		End If

		Return i.ToString()
	End Function

	Private Function RoundUp2(ByVal s As String) As Integer
		Dim num As Integer = CInt(s)

		If num = 0 Then
			'Nulls or dats in columns
			Return 100
		End If

		If (num < 10) Then Return Math.Ceiling(num / 10) * 10
		If (num < 100) Then Return Math.Ceiling(num / 100) * 100
		If (num < 1000) Then Return Math.Ceiling(num / 1000) * 1000

		Return num * 2
	End Function

	Private Sub chkScriptToFile_CheckedChanged(sender As Object, e As EventArgs) Handles chkScriptToFile.CheckedChanged
		If chkBulkInsert.Checked And chkScriptToFile.Checked Then
			MsgBox("Buld insert cannot be scripted to file")
			chkScriptToFile.Checked = False
		End If
	End Sub

	Private Sub chkBulkInsert_CheckedChanged(sender As Object, e As EventArgs) Handles chkBulkInsert.CheckedChanged
		If chkBulkInsert.Checked And chkScriptToFile.Checked Then
			MsgBox("Buld insert cannot be scripted to file")
			chkScriptToFile.Checked = False
		End If
	End Sub
End Class
