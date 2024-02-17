Imports LiveCharts
Imports LiveCharts.Defaults
Imports LiveCharts.Wpf
Public Class Annual

    Private Sub Monthly_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SwitchView()
    End Sub

    Private Sub Btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click

        sellerMode.Show()
        Me.Close()

    End Sub

    Private Sub Chart()
        Dim i As Integer = 0
        Dim sum As Integer = 0
        Dim sum1 As Integer = 0
        Dim sum2 As Integer = 0
        Dim sum3 As Integer = 0
        Dim dataMonth As Integer = 1
        Dim T = 0, V = 0, TC = 0, Tb = 0, Nt = 0, O = 0
        Dim series As New SeriesCollection
        Dim year As Integer = nUpDown.Value
        CartesianChart1.Series = series

        Dim salesData As New ChartValues(Of Decimal)({})

        Dim seriesSales As New LineSeries With {
            .Title = nUpDown.Value,
            .Values = salesData
        }
        series.Add(seriesSales)

        ' Add X Axis (Days)
        Dim xAxis, yAxis As New Axis
        xAxis.Labels = New List(Of String)()
        For d As Integer = 1 To 12
            xAxis.Labels.Add(d.ToString())
        Next
        CartesianChart1.AxisX.Clear()
        CartesianChart1.AxisX.Add(xAxis)

        ' Add Y Axis (Sales)
        yAxis.Title = "Total Sales"
        CartesianChart1.AxisY.Clear()
        CartesianChart1.AxisY.Add(yAxis)

        salesData.Clear()
        T = 0
        V = 0
        TC = 0
        While 12 >= dataMonth
            openConn()
            cmd.Connection = conn
            cmd.CommandText = "SELECT * FROM TRANSACTION where YEAR = '" & nUpDown.Value & "' AND MONTH = '" & dataMonth & "'"
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
            dataMonth += 1
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

    Private Sub NUpDown_ValueChanged(sender As Object, e As EventArgs) Handles nUpDown.ValueChanged
        Chart()
        ViewTable()
    End Sub

    Private Sub CmbMonth_SelectedIndexChanged(sender As Object, e As EventArgs)
        Chart()
        totalSaleMonth()
        viewTable()
    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Manage.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        priceUpdate.Show()
        Me.Close()
    End Sub

    Function TotalSaleMonth()
        Dim sum As Decimal = 0
        Dim sum1 As Decimal = 0
        Dim sum2 As Decimal = 0
        Dim sum3 As Decimal = 0
        Dim sum4 As Decimal = 0
        Dim i As Integer = 2
        openConn()
        cmd.Connection = conn
        cmd.CommandText = "SELECT * FROM TRANSACTION WHERE YEAR = '" & nUpDown.Value & "'"
        reader = cmd.ExecuteReader
        While reader.Read()
            sum += Val(reader.GetString("TOTAL"))
            sum1 += Val(reader.GetString("TABANGOHANON"))
            sum2 += Val(reader.GetString("NONTABANGOHANON"))
            sum3 += Val(reader.GetString("TOTALTICKETS"))
            sum4 += Val(reader.GetString("OTHERS"))
        End While
        reader.Close()
        conn.Close()

        Return sum

    End Function

    Private Sub Monthly_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Login.Close()
    End Sub

    Private Sub Switch_Click(sender As Object, e As EventArgs) Handles switch.Click
        switchView()
    End Sub

    Private Sub ViewTable()
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

    Private Sub SwitchView()
        If switch.Text <> "View as Table" Then
            Chart()
            TotalSaleMonth()
            switch.Text = "View as Table"
            DataGridView1.Visible = False
            CartesianChart1.Visible = True
        Else
            switch.Text = "View as Graph"
            TotalSaleMonth()
            ViewTable()
            CartesianChart1.Visible = False
            DataGridView1.Visible = True
        End If
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Export.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button5_Click(sender As Object, e As EventArgs) Handles Guna2Button5.Click
        Monthly.Show()
        Me.Close()
    End Sub

End Class