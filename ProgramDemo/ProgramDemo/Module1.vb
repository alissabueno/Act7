Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop.Excel
Module Module1
    Public Sub BackupTableToExcel(ByVal tableName As String, ByVal excelTemplatePath As String, ByVal excelOutputPath As String)

        ' Open MySQL connection
        Dim conn As MySqlConnection = New MySqlConnection("Server=127.0.0.1;Database=trialdb;Uid=root;Pwd=alissamoresbueno;")
        conn.Open()

        ' Create SQL command to select data from table
        Dim cmd As MySqlCommand = New MySqlCommand(String.Format("SELECT * FROM users", tableName), conn)

        ' Create MySQL data adapter
        Dim adapter As MySqlDataAdapter = New MySqlDataAdapter(cmd)

        ' Create new dataset
        Dim dataset As DataSet = New DataSet()

        ' Fill dataset with data from MySQL
        adapter.Fill(dataset, "MyTable")

        ' Close MySQL connection
        conn.Close()

        ' Open Excel template
        Dim excelApp As Application = New Application()
        Dim excelWorkbook As Workbook = excelApp.Workbooks.Open(excelTemplatePath)

        ' Get Excel worksheet
        Dim excelWorksheet As Worksheet = CType(excelWorkbook.Sheets("Sheet1"), Worksheet)

        ' Write dataset to Excel worksheet
        Dim rows As DataRowCollection = dataset.Tables("MyTable").Rows
        For i As Integer = 0 To rows.Count - 1
            Dim row As DataRow = rows(i)
            For j As Integer = 0 To row.ItemArray.Length - 1
                Dim value As Object = row.ItemArray(j)
                excelWorksheet.Cells(i + 10, j + 3) = value
            Next
        Next

        ' Save Excel workbook
        excelWorkbook.SaveAs(excelOutputPath)

        ' Close Excel workbook and application
        excelWorkbook.Close()
        excelApp.Quit()

        ' Show a message box to indicate success
        MessageBox.Show("Backup created successfully!")

    End Sub
End Module
