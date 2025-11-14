Imports System.Data.OleDb

Public Class frmEmployees
    Private conn As OleDbConnection
    Private da As OleDbDataAdapter
    Private ds As DataSet

    Private Sub frmEmployees_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=EmployeeDB.accdb;"
        conn = New OleDbConnection(connStr)

        da = New OleDbDataAdapter("SELECT * FROM Employees", conn)
        Dim cb As New OleDbCommandBuilder(da)

        ds = New DataSet()
        da.Fill(ds, "Employees")

        dgvEmployees.DataSource = ds.Tables("Employees")
        ClearFields()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If Not ValidateEmployee() Then Return

        Dim row = ds.Tables("Employees").NewRow()
        row("FirstName") = txtFirst.Text.Trim()
        row("LastName") = txtLast.Text.Trim()
        row("Department") = txtDept.Text.Trim()

        ds.Tables("Employees").Rows.Add(row)
        da.Update(ds, "Employees")

        lblStatus.Text = "Employee added."
        ClearFields()
        RefreshGrid()
    End Sub

    Private Sub RefreshGrid()
        ds.Tables("Employees").Clear()
        da.Fill(ds, "Employees")
    End Sub

    Private Sub ClearFields()
        txtID.Clear()
        txtFirst.Clear()
        txtLast.Clear()
        txtDept.Clear()
        txtFirst.Focus()
    End Sub

    Private Function ValidateEmployee() As Boolean
        If txtFirst.Text.Trim() = "" Or txtLast.Text.Trim() = "" Or txtDept.Text.Trim() = "" Then
            MessageBox.Show("All fields are required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        Return True
    End Function

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If conn.State = ConnectionState.Open Then conn.Close()
        Me.Close()
    End Sub
End Class
