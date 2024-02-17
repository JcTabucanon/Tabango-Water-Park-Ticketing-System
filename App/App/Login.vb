Imports System.ComponentModel

Public Class Login
    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        comId.Visible = False
        comUsername.Visible = False
        comUserpassword.Visible = False
        comsalt.Visible = False
        comDecrypt.Visible = False
        jbTitle.Visible = False
        firstname.Visible = False
        lastname.Visible = False

    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub txtUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUsername.KeyDown
        If e.KeyCode = Keys.Space Then
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub txtUsername_MouseClick(sender As Object, e As MouseEventArgs) Handles txtUsername.MouseClick
        txtUsername.PlaceholderText = ""
        txtUsername.Font = New Font("century gothic", 14)
        txtUsername.BorderColor = Color.Blue
    End Sub

    Private Sub txtPassword_MouseClick(sender As Object, e As MouseEventArgs) Handles txtPassword.MouseClick
        txtPassword.PlaceholderText = ""
        txtPassword.Font = New Font("century gothic", 14)
        txtPassword.BorderColor = Color.Blue
    End Sub

    Private Sub txtPassword_IconRightClick(sender As Object, e As EventArgs) Handles txtPassword.IconRightClick

        If txtPassword.PasswordChar = "●" Then
            txtPassword.PasswordChar = ""
        Else
            txtPassword.PasswordChar = "●"
        End If

    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click

        comId.Text = ""
        comUsername.Text = ""
        comUserpassword.Text = ""
        comsalt.Text = ""
        comDecrypt.Text = ""

        If String.IsNullOrWhiteSpace(txtUsername.Text) Then

            If String.IsNullOrWhiteSpace(txtPassword.Text) Then
                txtIsEmpty("all", txtUsername, txtPassword)
            Else
                txtIsEmpty("username", txtUsername, txtPassword)
            End If

        ElseIf String.IsNullOrWhiteSpace(txtPassword.Text) Then

            txtIsEmpty("password", txtUsername, txtPassword)

        Else

            openConn()
            Try
                cmd.Connection = conn
                cmd.CommandText = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & txtUsername.Text & "'"
                adapter.SelectCommand = cmd
                data.Clear()
                adapter.Fill(data, "List")

                comId.DataBindings.Add("Text", data, "List.ID")
                comId.DataBindings.Clear()

                jbTitle.DataBindings.Add("Text", data, "List.JOBTITLE")
                jbTitle.DataBindings.Clear()

                comUsername.DataBindings.Add("Text", data, "List.USERNAME")
                comUsername.DataBindings.Clear()

                comUserpassword.DataBindings.Add("Text", data, "List.PASSWORD")
                comUserpassword.DataBindings.Clear()

                comsalt.DataBindings.Add("Text", data, "List.SALT")
                comsalt.DataBindings.Clear()

                firstname.DataBindings.Add("Text", data, "List.FIRSTNAME")
                firstname.DataBindings.Clear()

                lastname.DataBindings.Add("Text", data, "List.LASTNAME")
                lastname.DataBindings.Clear()

                comDecrypt.Text = AES_Decrypt(comUserpassword.Text, comsalt.Text)

                conn.Close()

                If comUsername.Text <> txtUsername.Text Then

                    Beep()
                    MessageBox.Show("INVALID USERNAME !!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)

                ElseIf comUsername.Text = txtUsername.Text And comDecrypt.Text <> txtPassword.Text Then

                    Beep()
                    MessageBox.Show("WRONG PASSWORD !!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Else

                    If jbTitle.Text = "Seller" Then

                        Me.Hide()
                        cUser = "Seller"
                        sellerMode.Show()

                    ElseIf jbTitle.Text = "Admin" Then
                        cUser = "Admin"
                        Fullname = firstname.Text & " " & lastname.Text
                        Monthly.Show()
                        Me.Hide()
                    End If
                    'Me.Hide()
                    'Manage.Show()

                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

    End Sub

    Sub txtIsEmpty(whereError As String, ByRef txtBox As Object, ByRef txtBox1 As Object)

        Beep()
        If whereError = "username" Then
            txtBox.PlaceholderText = "Please provide username or email !!"
            txtBox.Font = New Font("century gothic", 12)
            txtBox.BorderColor = Color.Red
        ElseIf whereError = "password" Then
            txtBox1.PlaceholderText = "Please provide password !!"
            txtBox1.Font = New Font("century gothic", 12)
            txtBox1.BorderColor = Color.Red
        ElseIf whereError = "all" Then
            txtBox.PlaceholderText = "Please provide username or email !!"
            txtBox1.PlaceholderText = "Please provide password !!"
            txtBox.Font = New Font("century gothic", 12)
            txtBox.BorderColor = Color.Red
            txtBox1.Font = New Font("century gothic", 12)
            txtBox1.BorderColor = Color.Red
        End If

    End Sub

End Class
