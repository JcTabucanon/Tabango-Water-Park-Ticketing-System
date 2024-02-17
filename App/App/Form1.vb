Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        ' Set the column headers
        DataGridView1.Columns.Add("Description", "Description")
        DataGridView1.Columns.Add("Quantity", "Quantity")
        DataGridView1.Columns.Add("PricePerPerson", "Price/Person")
        DataGridView1.Columns.Add("EnvironmentalFee", "Environmental Fee")
        DataGridView1.Columns.Add("SubTotal", "SubTotal")

        ' Adjust column widths as needed
        DataGridView1.Columns("Description").Width = 200
        DataGridView1.Columns("Quantity").Width = 80
        DataGridView1.Columns("PricePerPerson").Width = 120
        DataGridView1.Columns("EnvironmentalFee").Width = 150
        DataGridView1.Columns("SubTotal").Width = 100

        ' Make the DataGridView read-only
        DataGridView1.ReadOnly = True

        ' Example data
        Dim salesReportData As String(,) = {
            {"TABANGOHANON", "1", "20", "", "30"},
            {"TABANGOHANON (CHILD/PWD/SENIOR)", "2", "15", "", "50"},
            {"NON-TABANGOHANON", "3", "25", "", "105"},
            {"NON-TABANGOHANON (CHILD/PWD/SENIOR)", "1", "20", "", "30"},
            {"OTHRES", "1", "", "", "30"},
            {"TICKETS SOLD", "10", "", "10", "100"}
}

        ' Add rows to the DataGridView
        For i As Integer = 0 To salesReportData.GetLength(0) - 1
            DataGridView1.Rows.Add(salesReportData(i, 0), salesReportData(i, 1), salesReportData(i, 2), salesReportData(i, 3), salesReportData(i, 4))
        Next

        ' Set the total value
        DataGridView1.Rows.Add("", "", "", "TOTAL :", "205")
        DataGridView1.Rows.Add("", "", "", "", "")

    End Sub

    Private Sub ExportToExcel()
        ' Create a new instance of Excel
        Dim excelApp As Excel.Application = New Excel.Application()

        ' Create a new workbook and sheet
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim sheet As Excel.Worksheet = workbook.Sheets(1)

        ' Copy data from DataGridView to sheet
        Dim dgv As DataGridView = DataGridView1
        Dim rowCount As Integer = dgv.Rows.Count - 2
        Dim colCount As Integer = dgv.Columns.Count

        ' FOR THE TITLE SIDE ------------------------------------------------------------------
        Dim range As Excel.Range = sheet.Range("B2:F2") ' Specify the range of cells to merge
        range.Merge() ' Merge the cells
        range.Value = "REPUBLIC OF THE PHILIPPINES" ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range.Interior.Color = System.Drawing.Color.DodgerBlue
        range.Font.Size = 12
        range.Font.Color = System.Drawing.Color.White
        range.Font.Bold = True
        range.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        range.WrapText = True

        ' Autofit the row height
        range.Rows.AutoFit()

        ' -------------------------------------------------------------------------------------
        Dim range1 As Excel.Range = sheet.Range("B3:F3") ' Specify the range of cells to merge
        range1.Merge() ' Merge the cells
        range1.Value = "MUNICIPALITY OF TABANGO" ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range1.Interior.Color = System.Drawing.Color.DodgerBlue
        range1.Font.Size = 12
        range1.Font.Color = System.Drawing.Color.White
        range1.Font.Bold = True
        range1.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        range1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        range1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        range1.WrapText = True

        ' Autofit the row height
        range1.Rows.AutoFit()

        '--------------------------------------------------------------------------------------

        Dim range2 As Excel.Range = sheet.Range("B4:F4") ' Specify the range of cells to merge
        range2.Merge() ' Merge the cells
        range2.Value = "PROVINCE OF LEYTE" ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range2.Interior.Color = System.Drawing.Color.DodgerBlue
        range2.Font.Size = 12
        range2.Font.Color = System.Drawing.Color.White
        range2.Font.Bold = True
        range2.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        range2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        range2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        range2.WrapText = True

        ' Autofit the row height
        range2.Rows.AutoFit()

        Dim range3 As Excel.Range = sheet.Range("B6:F6") ' Specify the range of cells to merge
        range3.Merge() ' Merge the cells
        range3.Value = " DAILY SALES REPORT " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range3.Interior.Color = System.Drawing.Color.DodgerBlue
        range3.Font.Size = 12
        range3.Font.Color = System.Drawing.Color.White
        range3.Font.Bold = True
        range3.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        range3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        range3.WrapText = True

        ' Autofit the row height
        range3.Rows.AutoFit()

        ' FOR THE TITLE SIDE--------------------------------------------------------------------

        Dim header1 As Excel.Range = sheet.Range("B7") ' Specify the range of cells to merge
        header1.Value = " Description " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header1.Interior.Color = System.Drawing.Color.DodgerBlue
        header1.Font.Size = 10
        header1.Font.Color = System.Drawing.Color.White
        header1.Font.Bold = True
        header1.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Autofit the row height
        header1.Rows.AutoFit()

        Dim header2 As Excel.Range = sheet.Range("C7") ' Specify the range of cells to merge
        header2.Value = " Quantity " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header2.Interior.Color = System.Drawing.Color.DodgerBlue
        header2.Font.Size = 10
        header2.Font.Color = System.Drawing.Color.White
        header2.Font.Bold = True
        header2.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Autofit the row height
        header2.Rows.AutoFit()

        Dim header3 As Excel.Range = sheet.Range("D7") ' Specify the range of cells to merge
        header3.Value = " PricePerPerson " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header3.Interior.Color = System.Drawing.Color.DodgerBlue
        header3.Font.Size = 10
        header3.Font.Color = System.Drawing.Color.White
        header3.Font.Bold = True
        header3.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Autofit the row height
        header3.Rows.AutoFit()

        Dim header4 As Excel.Range = sheet.Range("E7") ' Specify the range of cells to merge
        header4.Value = " EnvironmentalFee " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header4.Interior.Color = System.Drawing.Color.DodgerBlue
        header4.Font.Size = 10
        header4.Font.Color = System.Drawing.Color.White
        header4.Font.Bold = True
        header4.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header4.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Autofit the row height
        header4.Rows.AutoFit()

        Dim header5 As Excel.Range = sheet.Range("F7") ' Specify the range of cells to merge
        header5.Value = " SubTotal " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header5.Interior.Color = System.Drawing.Color.DodgerBlue
        header5.Font.Size = 10
        header5.Font.Color = System.Drawing.Color.White
        header5.Font.Bold = True
        header5.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header5.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Autofit the row height
        header5.Rows.AutoFit()

        'HEADER --------------------------------------------------------------------------------

        Dim header6 As Excel.Range = sheet.Range("E18:F18") ' Specify the range of cells to merge
        header6.Merge() ' Merge the cells
        header6.Value = " JESSIE T. CUNA " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header6.Interior.Color = System.Drawing.Color.White
        header6.Font.Size = 10
        header6.Font.Color = System.Drawing.Color.Black
        header6.Font.Bold = True
        header6.Font.Italic = False
        header6.Font.Underline = True

        ' Set the horizontal and vertical alignment to center
        header6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header6.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        header6.WrapText = True

        ' Autofit the row height
        header6.Rows.AutoFit()

        Dim header7 As Excel.Range = sheet.Range("E19:F19") ' Specify the range of cells to merge
        header7.Merge() ' Merge the cells
        header7.Value = " Prepared By " ' Set the cell value

        ' Set the background color, font size, font color, and font style
        header7.Interior.Color = System.Drawing.Color.White
        header7.Font.Size = 10
        header7.Font.Color = System.Drawing.Color.Black
        header7.Font.Bold = True
        header7.Font.Italic = False

        ' Set the horizontal and vertical alignment to center
        header7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        header7.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' Apply wrap text
        header7.WrapText = True

        ' Autofit the row height
        header7.Rows.AutoFit()


        'FOR SIGNATURE -------------------------------------------------------------------------


        ' Add data rows
        For i As Integer = 0 To rowCount - 1
            For j As Integer = 0 To colCount - 1
                If dgv.Rows(i).Cells(j).Value IsNot Nothing Then
                    sheet.Cells(i + 8, j + 2) = dgv.Rows(i).Cells(j).Value.ToString()
                    Dim column = sheet.Columns(j + 1)
                    column.AutoFit()
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
            'workbook.SaveAs(saveFileDialog.FileName, FileFormat:=Excel.XlFileFormat.xlWorkbookDefault, Password:="mypassword", WriteResPassword:="mypassword", ReadOnlyRecommended:=True)
            workbook.SaveAs(saveFileDialog.FileName, FileFormat:=Excel.XlFileFormat.xlWorkbookDefault, ReadOnlyRecommended:=True)
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

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        ExportToExcel()
    End Sub
End Class