Imports Excel = Microsoft.Office.Interop.Excel
Public Class Export

    Private Sub Export_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        choice()
        toShow()
        cmbMonth.SelectedIndex = DateAndTime.Now.Month - 1
        nUpDay.Value = DateAndTime.Now.Day
        nUpDown.Value = DateAndTime.Now.Year

    End Sub

    Private Sub choice()
        If exportType.SelectedIndex = -1 Then
            nUpDay.Visible = True
            cmbMonth.Visible = True
            nUpDown.Visible = True
        ElseIf exportType.SelectedIndex = 0 Then
            nUpDay.Visible = True
            cmbMonth.Visible = True
            nUpDown.Visible = True
        ElseIf exportType.SelectedIndex = 1 Then
            nUpDay.Visible = False
            cmbMonth.Visible = True
            nUpDown.Visible = True
        ElseIf exportType.SelectedIndex = 2 Then
            nUpDay.Visible = False
            cmbMonth.Visible = False
            nUpDown.Visible = True
        End If
    End Sub

    Private Sub toShow()

        If nUpDay.Visible = True And cmbMonth.Visible = True And nUpDown.Visible = True Then
            viewTableDay()
        ElseIf nUpDay.Visible = False And cmbMonth.Visible = True And nUpDown.Visible = True Then
            viewTableMonth()
        ElseIf nUpDay.Visible = False And cmbMonth.Visible = False And nUpDown.Visible = True Then
            viewTableYear()
        End If

    End Sub

    Private Sub exportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles exportType.SelectedIndexChanged, exportType.SelectedIndexChanged
        choice()
        toShow()
    End Sub

    Private Sub cmbMonth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMonth.SelectedIndexChanged
        toShow()
    End Sub

    Private Sub viewTableDay()
        Dim i As Integer = 0
        Dim sum As Integer = 0
        Dim sum1 As Integer = 0
        Dim sum2 As Integer = 0
        Dim mvalue As Integer = 0

        openConn()
        table.Reset()
        table = New DataTable
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE YEAR = '" & nUpDown.Value & "' AND MONTH = '" & cmbMonth.SelectedIndex + 1 & "' AND DAY = '" & nUpDay.Value & "'"
        adapter.SelectCommand = cmd
        table.Clear()
        adapter.Fill(table)
        DataGridView1.DataSource = table
        conn.Close()
    End Sub

    Private Sub viewTableMonth()
        Dim i As Integer = 0
        Dim sum As Integer = 0
        Dim sum1 As Integer = 0
        Dim sum2 As Integer = 0
        Dim mvalue As Integer = 0

        openConn()
        table.Reset()
        table = New DataTable
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE YEAR = '" & nUpDown.Value & "' AND MONTH = '" & cmbMonth.SelectedIndex + 1 & "'"
        adapter.SelectCommand = cmd
        table.Clear()
        adapter.Fill(table)
        DataGridView1.DataSource = table
        conn.Close()
    End Sub

    Private Sub viewTableYear()
        Dim i As Integer = 0
        Dim sum As Integer = 0
        Dim sum1 As Integer = 0
        Dim sum2 As Integer = 0
        Dim mvalue As Integer = 0

        openConn()
        table.Reset()
        table = New DataTable
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE YEAR = '" & nUpDown.Value & "'"
        adapter.SelectCommand = cmd
        table.Clear()
        adapter.Fill(table)
        DataGridView1.DataSource = table
        conn.Close()
    End Sub

    Private Sub nUpDay_ValueChanged(sender As Object, e As EventArgs) Handles nUpDay.ValueChanged
        toShow()
    End Sub

    Private Sub nUpDown_ValueChanged(sender As Object, e As EventArgs) Handles nUpDown.ValueChanged
        toShow()
    End Sub

    Private Sub ExportToExcel()
        ' Create a new instance of Excel
        Dim excelApp As Excel.Application = New Excel.Application()

        ' Create a new workbook and sheet
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim sheet As Excel.Worksheet = workbook.Sheets(1)

        ' Copy data from DataGridView to sheet
        Dim dgv As DataGridView = DataGridView1
        Dim rowCount As Integer = dgv.Rows.Count
        Dim colCount As Integer = dgv.Columns.Count

        ' Get the column headers from the DataGridView

        ' Add header row
        For j As Integer = 0 To colCount - 1
            sheet.Cells(1, j + 1) = dgv.Columns(j).HeaderText
        Next

        ' Add data rows
        For i As Integer = 0 To rowCount - 1
            For j As Integer = 0 To colCount - 1
                If dgv.Rows(i).Cells(j).Value IsNot Nothing Then
                    sheet.Cells(i + 2, j + 1) = dgv.Rows(i).Cells(j).Value.ToString()
                End If
            Next
        Next


        ' Prompt the user to choose where to save the file
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        saveFileDialog.FilterIndex = 1
        saveFileDialog.RestoreDirectory = True
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            ' Save the workbook to the selected location
            workbook.SaveAs(saveFileDialog.FileName, FileFormat:=Excel.XlFileFormat.xlWorkbookDefault, Password:="mypassword", WriteResPassword:="mypassword", ReadOnlyRecommended:=True)
            ' Close the workbook and Excel
            workbook.Close()
            excelApp.Quit()
            ' Release the objects from memory
            ReleaseObject(sheet)
            ReleaseObject(workbook)
            ReleaseObject(excelApp)
            MsgBox("Data exported to Excel successfully!")
        End If
    End Sub


    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Guna2Button6_Click(sender As Object, e As EventArgs) Handles Guna2Button6.Click
        ExportToExcel()
    End Sub

    Private Sub Guna2Button5_Click(sender As Object, e As EventArgs) Handles Guna2Button5.Click
        Monthly.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        priceUpdate.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Manage.Show()
        Me.Close()
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        sellerMode.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Annual.show
        Me.Close()
    End Sub



End Class