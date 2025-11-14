Imports System.Data.OleDb

Public Class frmPayroll
    Private conn As OleDbConnection
    Private employeeID As Integer = -1   ' selected from frmEmployees

    Private Sub frmPayroll_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=EmployeeDB.accdb;"
        conn = New OleDbConnection(connStr)
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Using f As New frmEmployees()
            f.StartPosition = FormStartPosition.CenterParent
            AddHandler f.EmployeeSelected, AddressOf EmployeeSelectedHandler
            f.ShowDialog()
        End Using
    End Sub

    Private Sub EmployeeSelectedHandler(empID As Integer, name As String)
        employeeID = empID
        txtName.Text = name
        txtHours.Focus()
    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click
        If Not ValidatePayroll() Then Return

        Dim hours As Double = CDbl(txtHours.Text)
        Dim rate As Double = CDbl(txtRate.Text)

        Dim ot As Boolean = chkOvertime.Checked And hours > 40
        Dim total As Double

        If ot Then
            Dim regular = 40 * rate
            Dim otHours = hours - 40
            Dim otPay = otHours * rate * 1.5
            total = regular + otPay
        Else
            total = hours * rate
        End If

        lblTotal.Text = "Total Pay: " & total.ToString("C2")

        ' ---- Save to Payroll table ----
        Using cmd As New OleDbCommand(
            "INSERT INTO Payroll (EmployeeID, HoursWorked, HourlyRate, TotalPay, Overtime) " &
            "VALUES (@eid, @hrs, @rate, @total, @ot)", conn)

            cmd.Parameters.AddWithValue("@eid", employeeID)
            cmd.Parameters.AddWithValue("@hrs", hours)
            cmd.Parameters.AddWithValue("@rate", rate)
            cmd.Parameters.AddWithValue("@total", total)
            cmd.Parameters.AddWithValue("@ot", If(ot, 1, 0))

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End Using

        MessageBox.Show("Payroll saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtName.Clear()
        txtHours.Clear()
        txtRate.Clear()
        lblTotal.Text = "Total Pay:"
        chkOvertime.Checked = False
        employeeID = -1
    End Sub

    Private Function ValidatePayroll() As Boolean
        If employeeID = -1 Then
            MessageBox.Show("Select an employee first.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Dim h, r As Double
        If Not Double.TryParse(txtHours.Text, h) Or h < 0 Then
            MessageBox.Show("Hours Worked must be a positive number.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If Not Double.TryParse(txtRate.Text, r) Or r <= 0 Then
            MessageBox.Show("Hourly Rate must be a positive number.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        Return True
    End Function
End Class
