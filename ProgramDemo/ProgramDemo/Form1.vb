Imports System.Data.Common
Imports Excel = Microsoft.Office.Interop.Excel
Imports MySql.Data.MySqlClient
Public Class Form1
    Dim connectionString As String = "Server=127.0.0.1;Database=trialdb;Uid=root;Pwd=alissamoresbueno;"
    Dim connection As MySqlConnection = New MySqlConnection(connectionString)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Module1.BackupTableToExcel("MyTable", "C:\Users\User\OneDrive\Documents\template.xlsx", "C:\Users\User\Desktop\MyBackup.xlsx")

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim query As String = "SELECT * FROM users"
        Dim adapter As MySqlDataAdapter = New MySqlDataAdapter(query, connection)
        Dim table As DataTable = New DataTable()
        adapter.Fill(table)
        DataGridView1.DataSource = table
    End Sub
End Class
