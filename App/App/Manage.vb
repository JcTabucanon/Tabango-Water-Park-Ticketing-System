Imports System.ComponentModel

Public Class Manage
    Dim userId As String
    Dim fn As String
    Dim ln As String
    Dim jb As String
    Dim g As String
    Dim bd As Date
    Dim usr As String

    Private Sub Manage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SelectedChoice()
        clr()
        comUsername.Visible = False

        'ViewDeletePanel.Visible = False
        btnSave.Visible = False
        btnEdit.Visible = False
        deletePopup.Visible = False
        'ViewDeletePanel.Visible = False
        editPopup.Visible = False
        Guna2HtmlLabel3.Visible = False
        btnShowAccount.Visible = False
        lblAdd.Visible = False
        btnCancel.Visible = False

    End Sub

    Public Sub SelectedChoice()
        If Choice.SelectedIndex = -1 Then

        ElseIf Choice.SelectedIndex = 0 Then

            ViewDeletePanel.Visible = True
            btnDelete.Visible = False
            btnRefresh.Visible = True
            editPopup.Visible = False
            loadTable()
            Guna2HtmlLabel3.Visible = False
            btnShowAccount.Visible = False
            btnCancel.Visible = False
            lblAdd.Visible = False

        ElseIf Choice.SelectedIndex = 1 Then

            clr()
            btnSave.Visible = True
            lblAdd.Visible = True
            ViewDeletePanel.Visible = False
            btnEdit.Visible = False
            editPopup.Visible = False
            Guna2HtmlLabel3.Visible = False
            btnShowAccount.Visible = False
            btnCancel.Visible = False
            btnRefresh.Visible = False
            btnDelete.Visible = False
            Guna2HtmlLabel1.Visible = False

        ElseIf Choice.SelectedIndex = 2 Then

            clr()
            ViewDeletePanel.Visible = False
            btnSave.Visible = False
            btnEdit.Visible = True
            lblAdd.Visible = False
            editPopup.Visible = False
            btnCancel.Visible = False
            Guna2HtmlLabel3.Visible = True
            btnShowAccount.Visible = True
            loadEdit()

        ElseIf Choice.SelectedIndex = 3 Then

            ViewDeletePanel.Visible = True
            btnDelete.Visible = True
            btnRefresh.Visible = False
            editPopup.Visible = False
            loadTable()
            Guna2HtmlLabel3.Visible = False
            btnShowAccount.Visible = False
            btnCancel.Visible = True
            lblAdd.Visible = False

        End If
    End Sub

    Private Sub userName_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub passWord_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub txtRetype_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ViewDeletePanel.Visible = False
        clr()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        comUsername.Text = ""

        If String.IsNullOrWhiteSpace(txtFname.Text) Or String.IsNullOrWhiteSpace(txtLname.Text) Or String.IsNullOrWhiteSpace(cmbJbtitle.Text
            ) Or String.IsNullOrWhiteSpace(cmbGender.Text) Or String.IsNullOrWhiteSpace(userName.Text
            ) Or String.IsNullOrWhiteSpace(passWord.Text) Or String.IsNullOrWhiteSpace(txtRetype.Text) Then

            Beep()
            MessageBox.Show("Please fill all !!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else

            If passWord.Text <> txtRetype.Text Then

                MessageBox.Show("Password dont match !!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else

                openConn()
                Try
                    cmd.Connection = conn
                    cmd.CommandText = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & userName.Text & "'"
                    adapter.SelectCommand = cmd
                    data.Clear()
                    adapter.Fill(data, "List")

                    comUsername.DataBindings.Add("Text", data, "List.USERNAME")
                    comUsername.DataBindings.Clear()

                    conn.Close()

                    If comUsername.Text = userName.Text Then

                        MessageBox.Show("Username already taken !!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Else

                        openConn()
                        Try
                            Dim SALT As String = Now.Ticks 'secret

                            cmd.Connection = conn
                            cmd.CommandText = "INSERT INTO ACCOUNT (`FIRSTNAME`, `LASTNAME`, `JOBTITLE`, `GENDER`, `BIRTHDAY`, `USERNAME`, `PASSWORD`, `SALT`) VALUES('" &
                                txtFname.Text & "', '" & txtLname.Text & "', '" & cmbJbtitle.SelectedItem & "', '" & cmbGender.SelectedItem & "', '" &
                                dtpbday.Value.Date & "', '" & userName.Text & "', '" & AES_Encrypt(passWord.Text, SALT) & "', '" & SALT & "')"
                            cmd.ExecuteNonQuery()
                            conn.Close()
                            Beep()
                            MessageBox.Show("Successfully Added !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            clr()

                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try

                    End If

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If
        End If

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        loadTable()
    End Sub

    Private Sub dtView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtView.CellClick
        userId = dtView.Item("USERNAME", dtView.CurrentRow.Index).Value
        dataId.Text = userId + " from list ?"
    End Sub

    Private Sub btnDalete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Beep()
        If userId = "" Then
            deleteWarningmsg.Text = "Please select from list !!"
            dataId.Text = "No selected item..."
            deletePopup.Visible = True
        Else

            deleteWarningmsg.Text = "Confirm to delete from list !!"
            deletePopup.Visible = True

        End If

    End Sub
    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        openConn()
        Try

            cmd.Connection = conn
            cmd.CommandText = "DELETE from ACCOUNT WHERE USERNAME = '" & userId & "'"
            cmd.ExecuteNonQuery()
            conn.Close()
            loadTable()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        userId = ""
        deletePopup.Visible = False
        loadTable()
    End Sub

    Private Sub canCel_Click(sender As Object, e As EventArgs) Handles canCel.Click
        deletePopup.Visible = False
    End Sub
    Sub clr()
        txtFname.Clear()
        txtLname.Clear()
        cmbJbtitle.SelectedIndex = -1
        cmbGender.SelectedIndex = -1
        dtpbday.Value = Today
        userName.Clear()
        passWord.Clear()
        txtRetype.Clear()
    End Sub

    Private Sub dtViewEdit_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtViewEdit.CellClick

        fn = dtViewEdit.Item("FIRSTNAME", dtViewEdit.CurrentRow.Index).Value
        ln = dtViewEdit.Item("LASTNAME", dtViewEdit.CurrentRow.Index).Value
        jb = dtViewEdit.Item("JOBTITLE", dtViewEdit.CurrentRow.Index).Value
        g = dtViewEdit.Item("GENDER", dtViewEdit.CurrentRow.Index).Value
        bd = dtViewEdit.Item("BIRTHDAY", dtViewEdit.CurrentRow.Index).Value
        usr = dtViewEdit.Item("USERNAME", dtViewEdit.CurrentRow.Index).Value


    End Sub
    Private Sub btnConfirmEdit_Click(sender As Object, e As EventArgs) Handles btnConfirmEdit.Click
        txtFname.Text = fn
        txtLname.Text = ln
        cmbJbtitle.Text = jb
        cmbGender.Text = g
        dtpbday.Text = bd
        userName.Text = usr
        editPopup.Hide()

    End Sub
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim idNum As Integer

        dataId.Text = ""
        comUsername.Text = ""

        If String.IsNullOrWhiteSpace(txtFname.Text) Or String.IsNullOrWhiteSpace(txtLname.Text) Or String.IsNullOrWhiteSpace(cmbJbtitle.Text
            ) Or String.IsNullOrWhiteSpace(cmbGender.Text) Or String.IsNullOrWhiteSpace(userName.Text
            ) Or String.IsNullOrWhiteSpace(passWord.Text) Or String.IsNullOrWhiteSpace(txtRetype.Text) Then

            Beep()
            MessageBox.Show("Please fill all !!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else
            MsgBox(passWord.Text)
            MsgBox(txtRetype.Text)

            If passWord.Text <> txtRetype.Text Then

                MessageBox.Show("Password dont match !!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else

                openConn()
                Try
                    cmd.Connection = conn
                    cmd.CommandText = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & userName.Text & "'"
                    adapter.SelectCommand = cmd
                    data.Clear()
                    adapter.Fill(data, "List")

                    comUsername.DataBindings.Add("Text", data, "List.USERNAME")
                    comUsername.DataBindings.Clear()

                    dataId.DataBindings.Add("Text", data, "List.ID")
                    dataId.DataBindings.Clear()

                    Int32.TryParse(dataId.Text, idNum)

                    MsgBox(idNum)
                    conn.Close()

                    openConn()
                    Try
                        Dim SALT As String = Now.Ticks 'secret

                        cmd.Connection = conn

                        cmd.CommandText = "UPDATE ACCOUNT SET FIRSTNAME = '" & txtFname.Text & "', LASTNAME = '" & txtLname.Text & "', JOBTITLE = '" & cmbJbtitle.SelectedItem & "', GENDER = '" &
                                cmbGender.SelectedItem & "', BIRTHDAY = '" & dtpbday.Value.Date & "', USERNAME = '" & userName.Text & "', PASSWORD = '" & AES_Encrypt(passWord.Text, SALT) & "', SALT = '" & SALT & "' WHERE ID = '" & idNum & "'"
                        cmd.ExecuteNonQuery()
                        conn.Close()
                        Beep()
                        MessageBox.Show("Successfully Added !!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        cmbJbtitle.Text = ""
                        cmbGender.Text = ""
                        clr()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If

        End If

    End Sub

    Sub loadTable()
        openConn()
        Try
            table.Reset()
            table = New DataTable
            cmd.Connection = conn
            cmd.CommandText = "SELECT `USERNAME`,`FIRSTNAME`, `LASTNAME`, `JOBTITLE`, `GENDER`, `BIRTHDAY`  FROM `account`"
            adapter.SelectCommand = cmd
            adapter.Fill(table)
            dtView.DataSource = table
            conn.Close()

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Sub loadEdit()
        openConn()
        Try
            table.Reset()
            table = New DataTable
            cmd.Connection = conn
            cmd.CommandText = "SELECT `USERNAME`,`FIRSTNAME`, `LASTNAME`, `JOBTITLE`, `GENDER`, `BIRTHDAY`  FROM `account`"
            adapter.SelectCommand = cmd
            adapter.Fill(table)
            dtViewEdit.DataSource = table
            conn.Close()
        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles btnRefreshEdit.Click
        loadEdit()
    End Sub

    Private Sub btnShowAccount_Click(sender As Object, e As EventArgs) Handles btnShowAccount.Click
        If editPopup.Visible = True Then
            editPopup.Visible = False
        Else
            editPopup.Visible = True
        End If
    End Sub

    Private Sub btnhide_Click(sender As Object, e As EventArgs) Handles btnhide.Click
        If editPopup.Visible = True Then
            editPopup.Visible = False
        Else
            editPopup.Visible = True
        End If
    End Sub

    Private Sub passWord_IconRightClick(sender As Object, e As EventArgs) Handles passWord.IconRightClick
        If passWord.PasswordChar = "●" Then
            passWord.PasswordChar = ""
        Else
            passWord.PasswordChar = "●"
        End If
    End Sub


    Private Sub txtRetype_IconRightClick(sender As Object, e As EventArgs) Handles txtRetype.IconRightClick
        If txtRetype.PasswordChar = "●" Then
            txtRetype.PasswordChar = ""
        Else
            txtRetype.PasswordChar = "●"
        End If
    End Sub

    Private Sub Manage_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Monthly.Show()
    End Sub

    Private Sub Choice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Choice.SelectedIndexChanged
        SelectedChoice()
    End Sub

    Private Sub Guna2Button6_Click(sender As Object, e As EventArgs) Handles Guna2Button6.Click
        Monthly.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Annual.Show()
        Me.Close()
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        sellerMode.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button2_Click_1(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        priceUpdate.Show()
        Me.Close()
    End Sub

    Private Sub Guna2Button5_Click(sender As Object, e As EventArgs) Handles Guna2Button5.Click
        Export.Show()
        Me.Close()
    End Sub
End Class