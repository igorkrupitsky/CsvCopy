Public Class AppSetting

	Public Sub SaveSettings(oHash As Hashtable)

		Dim oTable As New Data.DataTable

		For Each oEntry As Collections.DictionaryEntry In oHash
			Dim sColumnName As String = oEntry.Key.ToString()
			oTable.Columns.Add(New Data.DataColumn(sColumnName))
		Next


		Dim oDataRow As DataRow = oTable.NewRow()

		For Each oEntry As Collections.DictionaryEntry In oHash
			Dim sColumnName As String = oEntry.Key.ToString()
			oDataRow(sColumnName) = oEntry.Value
		Next

		oTable.Rows.Add(oDataRow)

		Dim ds As New Data.DataSet()
		ds.Tables.Add(oTable)
		ds.WriteXml(GetFilePath())

	End Sub

	Public Function GetSetting(ByVal sKey As String) As String
		Dim sFilePath As String = GetFilePath()

		If System.IO.File.Exists(sFilePath) = False Then
			Return ""
		End If

		Dim ds As New Data.DataSet()
		ds.ReadXml(sFilePath)
		If ds.Tables.Count = 0 Then
			Return ""
		End If

		Dim oTable As Data.DataTable = ds.Tables(0)
		If oTable.Columns.Contains(sKey) = False Then
			Return ""
		End If

		Return oTable.Rows(0)(sKey).ToString()
	End Function


	Private Function GetFilePath() As String
		Dim sFolder As String = Application.StartupPath()
		Dim sXmlFolder As String = System.IO.Path.Combine(sFolder, "CsvCopyXml")

		If Not System.IO.Directory.Exists(sXmlFolder) Then
			System.IO.Directory.CreateDirectory(sXmlFolder)
		End If

		Return System.IO.Path.Combine(sXmlFolder, "CsvCopySettings.xml")
	End Function


End Class