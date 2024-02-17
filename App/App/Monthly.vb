Imports LiveCharts
Imports LiveCharts.Defaults
Imports LiveCharts.Wpf
Imports Microsoft.Office.Interop
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel
Imports SeriesCollection = LiveCharts.SeriesCollection
Imports Axis = LiveCharts.Wpf.Axis

Public Class Monthly
    Public Sub dataPrint()
        ' Set the column headers
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear
        DataGridView2.Columns.Add("Description", "Description")
        DataGridView2.Columns.Add("Quantity", "Quantity")
        DataGridView2.Columns.Add("PricePerPerson", "Price/Person")
        DataGridView2.Columns.Add("EnvironmentalFee", "Environmental Fee")
        DataGridView2.Columns.Add("SubTotal", "SubTotal")

        ' Adjust column widths as needed
        DataGridView2.Columns("Description").Width = 200
        DataGridView2.Columns("Quantity").Width = 80
        DataGridView2.Columns("PricePerPerson").Width = 120
        DataGridView2.Columns("EnvironmentalFee").Width = 150
        DataGridView2.Columns("SubTotal").Width = 100

        ' Make the DataGridView read-only
        DataGridView2.ReadOnly = True

        ' Example data
        Dim salesReportData As String(,) = {
            {"TABANGOHANON", TNorm.Text, lblPrice.Text, "", LBLT.Text},
            {"TABANGOHANON (CHILD/PWD/SENIOR)", TABANGOHANONWITHDISCOUNT.Text, TwDISCOUNT.Text, "", Guna2HtmlLabel27.Text},
            {"NON-TABANGOHANON", NONTABANGOHANON.Text, NT.Text, "", Guna2HtmlLabel31.Text},
            {"NON-TABANGOHANON (CHILD/PWD/SENIOR)", NONTABANGOHANONDISCOUNT.Text, NTW.Text, "", Guna2HtmlLabel35.Text},
            {"OTHERS", Guna2HtmlLabel19.Text, "", "", Guna2HtmlLabel21.Text},
            {"ENVIRONMENTAL FEE TICKET", tSold.Text, "", envi.Text, Guna2HtmlLabel22.Text}
}

        ' Add rows to the DataGridView
        For i As Integer = 0 To salesReportData.GetLength(0) - 1
            DataGridView2.Rows.Add(salesReportData(i, 0), salesReportData(i, 1), salesReportData(i, 2), salesReportData(i, 3), salesReportData(i, 4))
        Next

        ' Set the total value
        DataGridView2.Rows.Add("", "", "", "TOTAL :", TOTAL.Text)
        DataGridView2.Rows.Add("", "", "", "", "")
    End Sub
    Private Sub ExportToExcel()
        dataPrint()
        ' Create a new instance of Excel
        Dim excelApp As Excel.Application = New Excel.Application()

        ' Create a new workbook and sheet
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim sheet As Excel.Worksheet = workbook.Sheets(1)

        ' Copy data from DataGridView to sheet
        Dim dgv As DataGridView = DataGridView2
        Dim rowCount As Integer = dgv.Rows.Count - 2
        Dim colCount As Integer = dgv.Columns.Count

        ' FOR THE TITLE SIDE ------------------------------------------------------------------
        Dim range As Excel.Range = sheet.Range("B2:F2") ' Specify the range of cells to merge
        range.Merge() ' Merge the cells
        range.Value = "REPUBLIC OF THE PHILIPPINES" ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range.Interior.Color = System.Drawing.Color.White
        range.Font.Size = 12
        range.Font.Color = System.Drawing.Color.Black
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
        range1.Interior.Color = System.Drawing.Color.White
        range1.Font.Size = 12
        range1.Font.Color = System.Drawing.Color.Black
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
        range2.Interior.Color = System.Drawing.Color.White
        range2.Font.Size = 12
        range2.Font.Color = System.Drawing.Color.Black
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
        range3.Value = salesTxt.Text & " " & "(" & dataDatee.Text & " " & ")" ' Set the cell value

        ' Set the background color, font size, font color, and font style
        range3.Interior.Color = System.Drawing.Color.White
        range3.Font.Size = 12
        range3.Font.Color = System.Drawing.Color.Black
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

        ' Autofit the row heightn        header5.Rows.AutoFit()

        'HEADER --------------------------------------------------------------------------------

        Dim header6 As Excel.Range = sheet.Range("E18:F18") ' Specify the range of cells to merge
        header6.Merge() ' Merge the cells
        header6.Value = Fullname ' Set the cell value

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
    Private Sub Guna2Button6_Click(sender As Object, e As EventArgs) Handles Guna2Button6.Click
        ExportToExcel()
    End Sub

    Private Sub Monthly_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView2.Visible = False
        cmbMonth.SelectedIndex = DateTime.Now.Month - 1
        nUpDown.Value = DateTime.Now.Year
        SwitchView()
        Nday.Value = DateAndTime.Now.Day
        Nday.Visible = False
        exportType.Visible = False
        price()

    End Sub

    Private Sub Btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click

        sellerMode.Show()
        Me.Hide()

    End Sub

    Private Sub Chart()
        price()
        Dim i As Integer = 0
        Dim sum As Integer = 0
        Dim T = 0, V = 0, TC = 0, Tb = 0, Nt = 0, O = 0
        Dim sum1 As Integer = 0
        Dim sum2 As Integer = 0
        Dim sum3 As Integer = 0
        Dim dataDay As Integer = 1
        Dim series As New SeriesCollection
        Dim selectedMonth As Integer
        Dim year As Integer = nUpDown.Value

        If cmbMonth.StartIndex = -1 Then
            selectedMonth = 1
        Else
            selectedMonth = cmbMonth.SelectedIndex
        End If
        CartesianChart1.Series = series

        Dim salesData As New ChartValues(Of Decimal)({})

        Dim seriesSales As New LineSeries With {
            .Title = cmbMonth.SelectedItem,
            .Values = salesData
        }
        series.Add(seriesSales)

        ' Add X Axis (Days)
        Dim xAxis As New Axis
        xAxis.Labels = New List(Of String)()
        For d As Integer = 1 To DateTime.DaysInMonth(nUpDown.Value, selectedMonth)
            xAxis.Labels.Add(d.ToString())
        Next
        CartesianChart1.AxisX.Clear()
        CartesianChart1.AxisX.Add(xAxis)

        ' Add Y Axis (Sales)
        Dim yAxis As New Axis
        yAxis.Title = "Total Sales"
        CartesianChart1.AxisY.Clear()
        CartesianChart1.AxisY.Add(yAxis)

        salesData.Clear()
        T = 0
        V = 0
        TC = 0
        While DateTime.DaysInMonth(year, selectedMonth) >= dataDay
            openConn()
            cmd.Connection = conn
            cmd.CommandText = "SELECT * FROM TRANSACTION where MONTH = '" & cmbMonth.SelectedIndex + 1 & "' AND  YEAR = '" & nUpDown.Value & "' AND DAY = '" & dataDay & "'"
            reader = cmd.ExecuteReader
            While reader.Read
                sum += Val(reader.GetString("TOTAL")) + Val(reader.GetString("OTHERTOTAL"))
                sum1 += Val(reader.GetString("TABANGOHANON")) + Val(reader.GetString("TABANGOHANONWITHDISCOUNT"))
                sum2 += Val(reader.GetString("NONTABANGOHANON")) + Val(reader.GetString("NONTABANGOHANONDISCOUNT"))
                sum3 += Val(reader.GetString("OTHERS"))
                Tb += Val(reader.GetString("TABANGOHANON")) + Val(reader.GetString("TABANGOHANONWITHDISCOUNT"))
                Nt += Val(reader.GetString("NONTABANGOHANON")) + Val(reader.GetString("NONTABANGOHANONDISCOUNT"))
                O += Val(reader.GetString("OTHERS"))
                TC += Val(reader.GetString("TOTALTICKETS"))
            End While

            ' Add Sales Data to Chart
            salesData.Add(sum)
            T += sum
            V += sum1 + sum2 + sum3
            dataDay += 1
            sum = 0
            sum1 = 0
            sum2 = 0
            sum3 = 0

            reader.Close()
            conn.Close()
        End While

        CartesianChart1.Invalidate()
        Pie(Tb, Nt, O)
        totalSales.Text = T
        totalVisitors.Text = V
        ticketSold.Text = TC

    End Sub

    Private Sub Pie(n As Integer, t As Integer, O As Integer)
        PieChart1.Series.Clear()
        Dim series1 As PieSeries = New PieSeries With {
            .Title = "Tabangohanon",
            .Values = New ChartValues(Of Double)({n}),
            .DataLabels = True,
            .LabelPoint = Function(p) String.Format("{1:P}", p, p.Participation)
        }
        Dim series2 As PieSeries = New PieSeries With {
            .Title = "Non-Tabangohanon",
            .Values = New ChartValues(Of Double)({t}),
            .DataLabels = True,
            .LabelPoint = Function(p) String.Format("{1:P}", p, p.Participation)
        }
        Dim series3 As PieSeries = New PieSeries With {
            .Title = "OTHERS",
            .Values = New ChartValues(Of Double)({O}),
            .DataLabels = True,
            .LabelPoint = Function(p) String.Format("{1:P}", p, p.Participation)
        }
        'set chart values
        PieChart1.LegendLocation = LegendLocation.Right
        PieChart1.Series.Add(series1)
        PieChart1.Series.Add(series2)
        PieChart1.Series.Add(series3)


    End Sub

    Private Sub SwitchView()
        If switch.Text <> "View Sales Report" Then
            Chart()
            switch.Text = "View Sales Report"
            DataGridView1.Visible = False
            Nday.Visible = False
            exportType.Visible = False
            Guna2Panel3.Visible = False
            Guna2GradientPanel1.Visible = True
            CartesianChart1.Visible = True
            Guna2GradientPanel4.Visible = True
            TotalSaleMonth()
        Else
            switch.Text = "View Dashboard"
            choice()
            exportType.Visible = True
            CartesianChart1.Visible = False
            DataGridView1.Visible = True
            Guna2Panel3.Visible = True
            Guna2GradientPanel1.Visible = False
            Guna2GradientPanel4.Visible = False
            TotalSaleMonth()
        End If
    End Sub
    Private Sub choice() '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If exportType.SelectedIndex = -1 Then
            Nday.Visible = True
            cmbMonth.Visible = True
            nUpDown.Visible = True
            salesTxt.Text = "DAILY SALES REPORT"
            TotalSaleDay()
        ElseIf exportType.SelectedIndex = 0 Then
            Nday.Visible = True
            cmbMonth.Visible = True
            nUpDown.Visible = True
            salesTxt.Text = "DAILY SALES REPORT"
            TotalSaleDay()
        ElseIf exportType.SelectedIndex = 1 Then
            Nday.Visible = False
            cmbMonth.Visible = True
            nUpDown.Visible = True
            salesTxt.Text = "MONTHLY SALES REPORT"
            TotalSaleMonth()
        ElseIf exportType.SelectedIndex = 2 Then
            Nday.Visible = False
            cmbMonth.Visible = False
            nUpDown.Visible = True
        End If
    End Sub

    Private Sub NUpDown_ValueChanged(sender As Object, e As EventArgs) Handles nUpDown.ValueChanged
        Chart()
        TotalSaleMonth()
        choice()
    End Sub

    Private Sub CmbMonth_SelectedIndexChanged(sender As Object, e As EventArgs)
        Chart()
        TotalSaleMonth()
    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Manage.Show()
        Me.Hide()
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        priceUpdate.Show()
        Me.Hide()
    End Sub

    Private Sub Monthly_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Login.Close()
    End Sub

    Private Sub Switch_Click(sender As Object, e As EventArgs) Handles switch.Click
        SwitchView()
    End Sub




    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Export.Show()
        Me.Hide()
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Annual.Show()
        Me.Hide()
    End Sub

    Private Sub cmbMonth_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cmbMonth.SelectedIndexChanged
        Chart()
        choice()
    End Sub

    Private Sub Nday_ValueChanged(sender As Object, e As EventArgs) Handles Nday.ValueChanged
        choice()
    End Sub

    Sub price()
        openConn()
        Try
            cmd.Connection = conn
            cmd.CommandText = "SELECT * FROM PRICE"
            adapter.SelectCommand = cmd
            data.Clear()
            adapter.Fill(data, "List")

            lblPrice.DataBindings.Add("Text", data, "List.TNORMAL")
            lblPrice.DataBindings.Clear()
            'p1 = Val(txtPriceTN.Text)
            lblPrice.Text = lblPrice.Text + ".00"

            TwDISCOUNT.DataBindings.Add("Text", data, "List.TWDISCOUNT")
            TwDISCOUNT.DataBindings.Clear()
            'p2 = Val(txtPriceNN.Text)
            TwDISCOUNT.Text = TwDISCOUNT.Text + ".00"


            NT.DataBindings.Add("Text", data, "List.NNORMAL")
            NT.DataBindings.Clear()
            'p1 = Val(txtPriceTN.Text)
            NT.Text = NT.Text + ".00"

            NTW.DataBindings.Add("Text", data, "List.NWDISCOUNT")
            NTW.DataBindings.Clear()
            'p2 = Val(txtPriceNN.Text)
            NTW.Text = NTW.Text + ".00"

            envi.DataBindings.Add("Text", data, "List.EnvironmentalFee")
            envi.DataBindings.Clear()
            'p2 = Val(txtPriceNN.Text)
            envi.Text = envi.Text + ".00"

            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Function TotalSaleMonth()
        Dim sum As Decimal = 0
        Dim sum1 As Decimal = 0
        Dim sum2 As Decimal = 0
        Dim sum3 As Decimal = 0
        Dim sum4 As Decimal = 0
        Dim sum5 As Decimal = 0
        Dim sum6 As Decimal = 0
        Dim sum7 As Decimal = 0
        Dim sum8 As Decimal = 0
        Dim datadate As Decimal = 0
        Dim i As Integer = 2
        openConn()
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE MONTH = '" & cmbMonth.SelectedIndex + 1 & "' AND YEAR = '" & nUpDown.Value & "'"
        reader = cmd.ExecuteReader
        While reader.Read()
            datadate = Val(reader.GetString("TRANSACTIONDATE"))
            sum1 += Val(reader.GetString("TABANGOHANON"))
            sum2 += Val(reader.GetString("NONTABANGOHANON"))
            sum5 += Val(reader.GetString("TABANGOHANONWITHDISCOUNT"))
            sum6 += Val(reader.GetString("NONTABANGOHANONDISCOUNT"))
            sum4 += Val(reader.GetString("OTHERS"))
            If Val(reader.GetString("OTHERS")) > 0 Then
                sum8 += 1
            End If
            sum += Val(reader.GetString("TOTAL"))
            sum3 += Val(reader.GetString("TOTALTICKETS"))
            sum7 += Val(reader.GetString("OTHERTOTAL")) 
        End While
        reader.Close()
        conn.Close()

        price()
        Guna2HtmlLabel22.Text = sum3 * Val(envi.Text)
        dataDatee.Text = cmbMonth.SelectedItem & " " & nUpDown.Value
        TNorm.Text = sum1
        TABANGOHANONWITHDISCOUNT.Text = sum5
        NONTABANGOHANONDISCOUNT.Text = sum6
        NONTABANGOHANON.Text = sum2
        tSold.Text = sum3
        'envi.Text
        LBLT.Text = sum1 * Val(lblPrice.Text)
        Guna2HtmlLabel21.Text = sum7 - (sum8 * Val(envi.Text))
        Guna2HtmlLabel27.Text = sum5 * Val(TwDISCOUNT.Text)
        Guna2HtmlLabel31.Text = sum2 * Val(NT.Text)
        Guna2HtmlLabel35.Text = sum6 * Val(NTW.Text)
        Guna2HtmlLabel19.Text = Val(sum4)
        TOTAL.Text = Val(Guna2HtmlLabel35.Text) + Val(Guna2HtmlLabel31.Text) + Val(Guna2HtmlLabel27.Text) + Val(LBLT.Text) + Val(Guna2HtmlLabel22.Text) + Val(Guna2HtmlLabel21.Text)


        Return sum

    End Function

    Function TotalSaleDay()
        Dim sum As Decimal = 0
        Dim sum1 As Decimal = 0
        Dim sum2 As Decimal = 0
        Dim sum3 As Decimal = 0
        Dim sum4 As Decimal = 0
        Dim sum5 As Decimal = 0
        Dim sum6 As Decimal = 0
        Dim sum7 As Decimal = 0
        Dim sum8 As Decimal = 0
        Dim datadate As Decimal = 0
        Dim i As Integer = 2
        openConn()
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE DAY = '" & Nday.Value & "' AND MONTH = '" & cmbMonth.SelectedIndex + 1 & "' AND YEAR = '" & nUpDown.Value & "'"
        reader = cmd.ExecuteReader
        While reader.Read()
            datadate = Val(reader.GetString("TRANSACTIONDATE"))
            sum1 += Val(reader.GetString("TABANGOHANON"))
            sum2 += Val(reader.GetString("NONTABANGOHANON"))
            sum5 += Val(reader.GetString("TABANGOHANONWITHDISCOUNT"))
            sum6 += Val(reader.GetString("NONTABANGOHANONDISCOUNT"))
            sum4 += Val(reader.GetString("OTHERS"))
            sum += Val(reader.GetString("TOTAL"))
            If Val(reader.GetString("OTHERS")) > 0 Then
                sum8 += 1
            End If
            sum3 += Val(reader.GetString("TOTALTICKETS"))
            sum7 += Val(reader.GetString("OTHERTOTAL"))
        End While
        reader.Close()
        conn.Close()


        price()
        dataDatee.Text = cmbMonth.SelectedItem & " " & Nday.Value & " " & nUpDown.Value
        TNorm.Text = sum1
        TABANGOHANONWITHDISCOUNT.Text = sum5
        NONTABANGOHANONDISCOUNT.Text = sum6
        NONTABANGOHANON.Text = sum2
        tSold.Text = sum3
        'envi.Text
        LBLT.Text = sum1 * Val(lblPrice.Text)
        Guna2HtmlLabel22.Text = sum3 * Val(envi.Text)

        Guna2HtmlLabel21.Text = sum7 - (sum8 * Val(envi.Text))
        Guna2HtmlLabel27.Text = sum5 * Val(TwDISCOUNT.Text)
        Guna2HtmlLabel31.Text = sum2 * Val(NT.Text)
        Guna2HtmlLabel35.Text = sum6 * Val(NTW.Text)
        Guna2HtmlLabel19.Text = Val(sum4)
        TOTAL.Text = Val(Guna2HtmlLabel35.Text) + Val(Guna2HtmlLabel31.Text) + Val(Guna2HtmlLabel27.Text) + Val(LBLT.Text) + Val(Guna2HtmlLabel22.Text) + Val(Guna2HtmlLabel21.Text)

        Return sum

    End Function

    Private Sub exportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles exportType.SelectedIndexChanged
        choice()
    End Sub

End Class