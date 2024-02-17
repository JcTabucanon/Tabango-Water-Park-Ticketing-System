Imports System.ComponentModel

Public Class priceUpdate
    Private Sub priceUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        price()
        loadprice(txtClassification.Text)

    End Sub

    Sub price()
        openConn()
        Try
            cmd.Connection = conn
            cmd.CommandText = "SELECT * FROM PRICE"
            adapter.SelectCommand = cmd
            data.Clear()
            adapter.Fill(data, "List")

            lblRegular.DataBindings.Add("Text", data, "List.TNORMAL")
            lblRegular.DataBindings.Clear()
            'p1 = Val(txtPriceTN.Text)
            lblRegular.Text = lblRegular.Text + ".00"

            lblWdiscount.DataBindings.Add("Text", data, "List.TWDISCOUNT")
            lblWdiscount.DataBindings.Clear()
            'p2 = Val(txtPriceNN.Text)
            lblWdiscount.Text = lblWdiscount.Text + ".00"

            lblEnvi.DataBindings.Add("Text", data, "List.EnvironmentalFee")
            lblEnvi.DataBindings.Clear()
            'p2 = Val(txtPriceNN.Text)
            lblEnvi.Text = lblEnvi.Text + ".00"
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub tglClassification_CheckedChanged(sender As Object, e As EventArgs) Handles tglClassification.CheckedChanged

        If tglClassification.Checked = True Then
            txtClassification.Text = "NON-TABANGOHANON"
            txtPriceboard.Text = "NON-TABANGOHANON"

            loadprice(txtClassification.Text)

        Else
            txtClassification.Text = "TABANGOHANON"
            txtPriceboard.Text = "TABANGOHANON"

            loadprice(txtClassification.Text)

        End If

    End Sub



    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        prs(txtClassification.Text)
    End Sub

    Sub prs(v As String)
        Dim ID As Integer
        If v = "TABANGOHANON" Then
            openConn()
            Try
                If String.IsNullOrWhiteSpace(txtN.Text) Or String.IsNullOrWhiteSpace(txtN.Text) Then
                    MessageBox.Show("Please fill all form !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Dim N As Integer = 0
                    Dim D As Integer = 0
                    Dim E As Integer = 0
                    E = Val(txtEnvi.Text)
                    N = Val(txtN.Text)
                    D = Val(txtD.Text)
                    ID = 1
                    openConn()

                    cmd.Connection = conn
                    cmd.CommandText = "UPDATE PRICE SET TNORMAL = '" & N & "', TWDISCOUNT = '" & D & "', EnvironmentalFee = '" & E & "'  WHERE ID = '" & ID & "'"
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    Beep()
                    MessageBox.Show("Successfully Added !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        ElseIf v = "NON-TABANGOHANON" Then
            Try
                If String.IsNullOrWhiteSpace(txtN.Text) Or String.IsNullOrWhiteSpace(txtN.Text) Then
                    MessageBox.Show("Please fill all form !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Dim N As Integer = 0
                    Dim D As Integer = 0
                    Dim E As Integer = 0
                    ID = 1
                    N = Val(txtN.Text)
                    D = Val(txtD.Text)
                    E = Val(txtEnvi.Text)
                    openConn()

                    cmd.Connection = conn
                    cmd.CommandText = "UPDATE PRICE SET NNORMAL = '" & N & "', NWDISCOUNT = '" & D & "', EnvironmentalFee = '" & E & "' WHERE ID = '" & ID & "'"
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    Beep()
                    MessageBox.Show("Successfully Added !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)

                End If
            Catch ex As Exception
                MsgBox(ex.ToString)

            End Try
        End If
        loadprice(txtClassification.Text)
    End Sub
    Sub loadprice(v As String)


        If v = "TABANGOHANON" Then
            openConn()
            Try
                cmd.Connection = conn
                cmd.CommandText = "SELECT * FROM PRICE WHERE ID = '" & 1 & "'"
                adapter.SelectCommand = cmd
                data.Clear()
                adapter.Fill(data, "List")

                lblRegular.DataBindings.Add("Text", data, "List.TNORMAL")
                lblRegular.DataBindings.Clear()

                lblWdiscount.DataBindings.Add("Text", data, "List.TWDISCOUNT")
                lblWdiscount.DataBindings.Clear()

                lblEnvi.DataBindings.Add("Text", data, "List.EnvironmentalFee")
                lblEnvi.DataBindings.Clear()

                conn.Close()
            Catch ex As Exception

            End Try
        ElseIf v = "NON-TABANGOHANON" Then
            openConn()
            Try
                cmd.Connection = conn
                cmd.CommandText = "SELECT * FROM PRICE WHERE ID = '" & 1 & "'"
                adapter.SelectCommand = cmd
                data.Clear()
                adapter.Fill(data, "List")

                lblRegular.DataBindings.Add("Text", data, "List.NNORMAL")
                lblRegular.DataBindings.Clear()

                lblWdiscount.DataBindings.Add("Text", data, "List.NWDISCOUNT")
                lblWdiscount.DataBindings.Clear()

                lblEnvi.DataBindings.Add("Text", data, "List.EnvironmentalFee")
                lblEnvi.DataBindings.Clear()

                conn.Close()
            Catch ex As Exception

            End Try
        End If


    End Sub

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        Export.Show()
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
        Annual.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button5_Click(sender As Object, e As EventArgs) Handles Guna2Button5.Click
        Monthly.Show()
        Me.Close()
    End Sub
End Class