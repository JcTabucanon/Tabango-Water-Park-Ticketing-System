Imports System.ComponentModel
Public Class sellerMode
    Dim LIMIT As Integer = 1
    Dim TransType As Integer
    Dim who As String
    Dim totalsssss As Integer
    Dim sub_Total As Integer
    Dim total As Integer
    Dim TNTotal As Integer = 0
    Dim NTNTotal As Integer = 0
    Dim TWDTotal As Integer = 0
    Dim NTWDTotal As Integer = 0
    Dim totalTickets As Integer
    Dim p1, p2, p3, p4, Eprice As Integer
    Dim change As Integer
    Dim currentDate As Date = DateTime.Now.Date


    Private Sub SellerMode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Alternative.Visible = False

        ticketLayout.Visible = False
        Timer1.Start()
        subTotaltxt.Text = "0.00"
        txtChange.Text = 0

        openConn()
        Try
            cmd.Connection = conn
            cmd.CommandText = "SELECT * FROM PRICE"
            adapter.SelectCommand = cmd
            data.Clear()
            adapter.Fill(data, "List")

            txtPriceTN.DataBindings.Add("Text", data, "List.TNORMAL")
            txtPriceTN.DataBindings.Clear()
            p1 = Val(txtPriceTN.Text)
            txtPriceTN.Text = txtPriceTN.Text + ".00"

            txtPriceNN.DataBindings.Add("Text", data, "List.NNORMAL")
            txtPriceNN.DataBindings.Clear()
            p2 = Val(txtPriceNN.Text)
            txtPriceNN.Text = txtPriceNN.Text + ".00"

            txtPriceTW.DataBindings.Add("Text", data, "List.TWDISCOUNT")
            txtPriceTW.DataBindings.Clear()
            p3 = Val(txtPriceTW.Text)
            txtPriceTW.Text = txtPriceTW.Text + ".00"

            txtPriceNW.DataBindings.Add("Text", data, "List.NWDISCOUNT")
            txtPriceNW.DataBindings.Clear()
            p4 = Val(txtPriceNW.Text)
            txtPriceNW.Text = txtPriceNW.Text + ".00"

            lblEnvi.DataBindings.Add("Text", data, "List.EnvironmentalFee")
            lblEnvi.DataBindings.Clear()
            Eprice = Val(lblEnvi.Text)
            lblEnvi.Text = lblEnvi.Text + ".00"

            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub


    Private Sub LogoutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Login.Show()

        Me.Close()
    End Sub


    Function Subtract(bal As Integer, ByRef txt As Object, S As String, p As Integer)
        Dim subtotal As Integer = 0
        If bal > 0 Then
            If S = "W" Then
                subtotal = 0
                bal -= 1
                txt.Text = bal
                subtotal = p
            ElseIf S = "N" Then
                subtotal = 0
                bal -= 1
                txt.Text = bal
                subtotal = p
            End If
        End If

        total -= subtotal
        subTotaltxt.Text = total
        subTotaltxt.Text = subTotaltxt.Text & ".00"
        txtTotal.Text = total + Eprice
        txtTotal.Text = txtTotal.Text & ".00"
        Return bal

    End Function

    Function Add(bal As Integer, ByRef txt As Object, S As String, p As Integer)

        Dim subtotal As Integer = 0
        If S = "W" Then
            bal += 1
            txt.Text = bal
            subtotal = p
        ElseIf S = "N" Then
            bal += 1
            txt.Text = bal
            subtotal = p
        End If
        total += subtotal
        subTotaltxt.Text = total
        subTotaltxt.Text = subTotaltxt.Text + ".00"
        txtTotal.Text = total + Eprice
        txtTotal.Text = txtTotal.Text & ".00"
        Return bal

    End Function

    Private Sub TxtEnterMoney_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEnterMoney.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub SellerMode_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing


        ' MessageBox.Show("You will be logout !!", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)

    End Sub

    Private Sub SellerMode_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'MessageBox.Show("You will be logout !!", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If cUser = "Admin" Then
            Monthly.Show()
        End If
        Login.Close()
    End Sub

    Private Sub TxtEnterMoney_KeyUp(sender As Object, e As KeyEventArgs) Handles txtEnterMoney.KeyUp
        If txtEnterMoney.Text = "" Then
            txtChange.Text = "0.00"
        Else
            Dim a As Integer = 0
            Dim b As Integer = 0
            change = 0
            a = Val(txtEnterMoney.Text)
            b = Val(subTotaltxt.Text)

            change = (a - b)
            change = change - Val(lblEnvi.Text)

            txtChange.Text = change
            txtChange.Text = txtChange.Text + ".00"

        End If
    End Sub

    Private Sub BtnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click

        If TransType = 1 Then
            SaveTrans()
        ElseIf TransType = 2 Then
            SaveTrans1()
        End If
        'AddData()



    End Sub

    'printing ---------------------------------------------------------------------------------------------------------
    Private Sub PrintPageHandler(ByVal sender As Object, ByVal args As Printing.PrintPageEventArgs)
        Dim text1 As String = title1.Text
        args.Graphics.DrawString(text1, title1.Font, Brushes.Black, 3, 0)

        Dim text2 As String = lblDate.Text
        args.Graphics.DrawString(text2, lblDate.Font, Brushes.Black, 20, 15)
        Dim text3 As String = tDate.Text
        args.Graphics.DrawString(text3, tDate.Font, Brushes.Black, 48, 15)
        '-------------- Code number---------------------------------------------

        Dim cod As String = code.Text
        args.Graphics.DrawString(cod, code.Font, Brushes.Black, 20, 25)
        Dim codee As String = txtCode.Text
        args.Graphics.DrawString(codee, txtCode.Font, Brushes.Black, 115, 25)
        '--------------No Discount---------------------------------------------
        Dim nodisLine As String = nodiscount.Text
        args.Graphics.DrawString(nodisLine, nodiscount.Font, Brushes.Black, 0, 40)

        Dim txtTN As String = TabangohanonN.Text
        args.Graphics.DrawString(txtTN, TabangohanonN.Font, Brushes.Black, 0, 55)
        Dim txtTotalTN As String = TpN.Text
        args.Graphics.DrawString(txtTotalTN, TpN.Font, Brushes.Black, 120, 55)

        Dim txtNTN As String = nonTabangohanonN.Text
        args.Graphics.DrawString(txtNTN, nonTabangohanonN.Font, Brushes.Black, 0, 70)
        Dim txtTotalNTN As String = NTpN.Text
        args.Graphics.DrawString(txtTotalNTN, NTpN.Font, Brushes.Black, 120, 70)
        '------------------ With Discount ---------------------------------------
        Dim wdisLine As String = wDiscount.Text
        args.Graphics.DrawString(wdisLine, wDiscount.Font, Brushes.Black, 0, 85)

        Dim txtTD As String = TabangohanonD.Text
        args.Graphics.DrawString(txtTD, TabangohanonD.Font, Brushes.Black, 0, 100)
        Dim txtTotalTD As String = TpWd.Text
        args.Graphics.DrawString(txtTotalTD, TpWd.Font, Brushes.Black, 120, 100)

        Dim txtNTD As String = nonTabangohanonD.Text
        args.Graphics.DrawString(txtNTD, nonTabangohanonD.Font, Brushes.Black, 0, 115)
        Dim txtTotalNTD As String = NTpWd.Text
        args.Graphics.DrawString(txtTotalNTD, NTpWd.Font, Brushes.Black, 120, 115)
        '------------------------- Others  ---------------------------------------

        Dim O As String = Guna2HtmlLabel15.Text
        args.Graphics.DrawString(O, Guna2HtmlLabel15.Font, Brushes.Black, 0, 130)
        Dim OT As String = OtherPrint.Text
        args.Graphics.DrawString(OT, OtherPrint.Font, Brushes.Black, 120, 130)
        '------------------------- Total  ---------------------------------------
        Dim txtSub As String = txtSubtotal.Text
        args.Graphics.DrawString(txtSub, txtSubtotal.Font, Brushes.Black, 0, 145)
        Dim subTtal As String = subTotal.Text
        args.Graphics.DrawString(subTtal, subTotal.Font, Brushes.Black, 120, 145)

        Dim enVI As String = txtEnviFee.Text
        args.Graphics.DrawString(enVI, txtEnviFee.Font, Brushes.Black, 0, 165)
        Dim enVip As String = enviFee.Text
        args.Graphics.DrawString(enVip, enviFee.Font, Brushes.Black, 120, 165)

        Dim totl As String = txtTotall.Text
        args.Graphics.DrawString(totl, txtTotall.Font, Brushes.Black, 0, 185)
        Dim totll As String = totall.Text
        args.Graphics.DrawString(totll, totall.Font, Brushes.Black, 120, 185)

        Dim textkwarta As String = txtkwarta.Text
        args.Graphics.DrawString(textkwarta, txtkwarta.Font, Brushes.Black, 0, 205)
        Dim kwartaa As String = kwarta.Text
        args.Graphics.DrawString(kwartaa, kwarta.Font, Brushes.Black, 120, 205)

        Dim textsukli As String = txtsukli.Text
        args.Graphics.DrawString(textsukli, txtsukli.Font, Brushes.Black, 0, 225)
        Dim suklii As String = sukli.Text
        args.Graphics.DrawString(suklii, sukli.Font, Brushes.Black, 120, 225)
        '------------------------- Message  ---------------------------------------
        Dim m1 As String = Gm1.Text
        args.Graphics.DrawString(m1, Gm1.Font, Brushes.Black, 20, 245)

        Dim m2 As String = Gm2.Text
        args.Graphics.DrawString(m2, Gm2.Font, Brushes.Black, 100, 245)

        Dim m3 As String = Gm3.Text
        args.Graphics.DrawString(m3, Gm1.Font, Brushes.Crimson, 0, 265)



    End Sub


    'printing ---------------------------------------------------------------------------------------------------------



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        clock.Text = TimeOfDay
    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        Alternative.Visible = True
    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click

        TransType = 2
        subTotaltxt.Text = Guna2TextBox1.Text
        txtTotal.Text = Val(subTotaltxt.Text) + Val(lblEnvi.Text)

    End Sub

    Private Sub BtnSum1_Click(sender As Object, e As EventArgs) Handles BtnSum1.Click

        TransType = 1

        Dim total1 As Integer = 0
        Dim total2 As Integer = 0
        Dim total3 As Integer = 0
        Dim total4 As Integer = 0
        Dim total5 As Integer = 0

        total1 = Guna2NumericUpDown1.Value * Val(txtPriceTN.Text)
        total2 = Guna2NumericUpDown2.Value * Val(txtPriceNN.Text)
        total3 = Guna2NumericUpDown3.Value * Val(txtPriceTW.Text)
        total4 = Guna2NumericUpDown4.Value * Val(txtPriceNW.Text)

        subTotaltxt.Text = (total1 + total2 + total3 + total4).ToString
        txtTotal.Text = Val(subTotaltxt.Text) + Val(lblEnvi.Text)


    End Sub

    Private Sub Guna2HtmlLabel5_Click(sender As Object, e As EventArgs) Handles Guna2HtmlLabel5.Click

    End Sub

    Sub PrintConfirm()
        Dim printDoc As New Printing.PrintDocument
        AddHandler printDoc.PrintPage, AddressOf PrintPageHandler
        printDoc.PrintController = New Printing.StandardPrintController()
        printDoc.Print()
    End Sub

    Sub SaveTrans()
        Dim sum As Integer = 0

        Dim currentDate As Date = DateTime.Now.Date

            openConn()
            Try
                If Val(subTotaltxt.Text) = 0 Then

                    MessageBox.Show("No ticket to be print  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Else

                    If String.IsNullOrWhiteSpace(txtEnterMoney.Text) Then
                        MessageBox.Show("Please enter money  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Else
                        Dim eFee = Eprice
                        Dim money As Integer = 0
                        Dim total As Integer = 0
                        money = Val(txtEnterMoney.Text)
                        total = Val(subTotaltxt.Text) + eFee
                        If money < total Then
                            MessageBox.Show("Insuffient money  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            Beep()
                            cmd.Connection = conn
                            cmd.CommandText = "INSERT INTO TRANSACTION ( `TRANSACTIONDATE`, `TABANGOHANON`, `TABANGOHANONWITHDISCOUNT`, `NONTABANGOHANON`, `NONTABANGOHANONDISCOUNT`, `TOTAL`, `TOTALTICKETS`, `DAY`, `MONTH`, `YEAR`) VALUES('" &
                                currentDate.Date & "', '" & Guna2NumericUpDown1.Value & "', '" & Guna2NumericUpDown3.Value & "', '" & Guna2NumericUpDown2.Value & "', '" & Guna2NumericUpDown4.Value & "', '" & txtTotal.Text & "', '" & 1 & "', '" & currentDate.Day & "', '" & currentDate.Month & "', '" & currentDate.Year & "')"

                            tDate.Text = currentDate.Month & "-" & currentDate.Day & "-" & currentDate.Year & " " & "at" & " " & clock.Text
                            TpN.Text = Guna2NumericUpDown1.Value & " " & "X" & " " & txtPriceTN.Text
                            NTpN.Text = Guna2NumericUpDown2.Value & " " & "X" & " " & txtPriceNN.Text
                            TpWd.Text = Guna2NumericUpDown3.Value & " " & "X" & " " & txtPriceTW.Text
                            NTpWd.Text = Guna2NumericUpDown4.Value & " " & "X" & " " & txtPriceNW.Text

                            OtherPrint.Text = Guna2TextBox1.Text & " " & "/" & " " & Guna2TextBox2.Text

                            subTotal.Text = subTotaltxt.Text
                            enviFee.Text = eFee
                            totall.Text = total
                            kwarta.Text = txtEnterMoney.Text
                            sukli.Text = txtChange.Text
                            cmd.ExecuteNonQuery()
                            conn.Close()
                            openConn()
                            cmd.Connection = conn
                            cmd.CommandText = "SELECT * FROM TRANSACTION "
                            reader = cmd.ExecuteReader
                            While reader.Read()
                                sum = Val(reader.GetString("ID"))
                            End While
                            reader.Close()
                            conn.Close()
                            txtCode.Text = sum
                            MessageBox.Show("Click enter to print!!", "Confirm?", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            LIMIT -= 1
                            PrintConfirm()
                        End If

                    End If

                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

    End Sub

    Sub SaveTrans1()
        Dim currentDate As Date = DateTime.Now.Date

        openConn()
            Try
                If Val(subTotaltxt.Text) = 0 Then

                    MessageBox.Show("No ticket to be print  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Else

                    If String.IsNullOrWhiteSpace(txtEnterMoney.Text) Then
                        MessageBox.Show("Please enter money  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Else
                        Dim eFee = Eprice
                        Dim money As Integer = 0
                        Dim total As Integer = 0
                        money = Val(txtEnterMoney.Text)
                        total = Val(subTotaltxt.Text) + eFee
                        If money < total Then
                            MessageBox.Show("Insuffient money  !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            Beep()
                            cmd.Connection = conn
                        cmd.CommandText = "INSERT INTO TRANSACTION ( `TRANSACTIONDATE`,`OTHERS`, `OTHERTOTAL`, `TOTALTICKETS`, `DAY`, `MONTH`, `YEAR`) VALUES('" &
                                currentDate.Date & "', '" & Guna2TextBox2.Text & "', '" & txtTotal.Text & "', '" & 1 & "', '" & currentDate.Day & "', '" & currentDate.Month & "', '" & currentDate.Year & "')"

                        tDate.Text = currentDate.Month & "-" & currentDate.Day & "-" & currentDate.Year & " " & "at" & " " & clock.Text
                            TpN.Text = Guna2NumericUpDown1.Value & " " & "X" & " " & txtPriceTN.Text
                            NTpN.Text = Guna2NumericUpDown2.Value & " " & "X" & " " & txtPriceNN.Text
                            TpWd.Text = Guna2NumericUpDown3.Value & " " & "X" & " " & txtPriceTW.Text
                            NTpWd.Text = Guna2NumericUpDown4.Value & " " & "X" & " " & txtPriceNW.Text
                            OtherPrint.Text = Guna2TextBox1.Text & " " & "/" & " " & Guna2TextBox2.Text

                            subTotal.Text = subTotaltxt.Text
                            enviFee.Text = eFee
                            totall.Text = total
                            kwarta.Text = txtEnterMoney.Text
                            sukli.Text = txtChange.Text
                            cmd.ExecuteNonQuery()
                            conn.Close()
                            MessageBox.Show("Click enter to print!!", "Confirm?", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            LIMIT -= 1
                            PrintConfirm()
                        End If

                    End If

                End If

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


    End Sub


    Sub AddData()
        Dim TbT As Integer
        Dim NTbT As Integer
        Dim money As Integer = 100
        Dim sg As Integer = 5
        Dim total As Integer = 300
        totalTickets = 5
        TbT = 3
        NTbT = 2

        Try
            While totalsssss < 5000
                totalsssss += 1
                openConn()
                TbT += 1
                NTbT += 1
                totalTickets += 1


                money += 100
                total += 100
                cmd.Connection = conn
                cmd.CommandText = "INSERT INTO TRANSACTION ( `TRANSACTIONDATE`, `TABANGOHANON`, `NONTABANGOHANON`, `TOTAL`, `TOTALTICKETS`, `DAY`, `MONTH`, `YEAR`) VALUES('" &
                            currentDate.Date & "', '" & TbT & "', '" & NTbT & "', '" & money & "', '" & totalTickets & "', '" & currentDate.Day & "', '" & currentDate.Month & "', '" & currentDate.Year & "')"
                cmd.ExecuteNonQuery()
                conn.Close()
                If totalsssss = sg Then
                    currentDate = currentDate.AddDays(1)
                    sg += 10
                End If

            End While

            Beep()
            MessageBox.Show("Successfully Added !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class