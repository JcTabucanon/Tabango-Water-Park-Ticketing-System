Imports MySql.Data.MySqlClient
Module database


    'For my server
    Public ipAddress As String = "127.0.0.1"
    Public userPassword As String = ""

    'mysqlCommand or Syntax
    Public cmd As New MySqlCommand

    'To fetch record and select cmd or select query
    Public adapter As New MySqlDataAdapter

    'If  mysqldataAdatepter is use to fetch data on textbox
    Public data As New DataSet

    'use for datagrid view
    Public table As New DataTable

    'reader
    Public reader As MySqlDataReader

    'mysqlConnection
    Public conn As MySqlConnection

    Public server As String = ipAddress
    Public userId As String = "root"
    Public password As String = userPassword
    Public myDatabase As String = "TWPTS"

    Public Function openConn()

        conn = New MySqlConnection()

        conn.ConnectionString = "server=" & server & ";" _
        & "user id=" & userId & ";" _
        & "password=" & password & ";" _
        & "database=" & myDatabase

        Try
            conn.Open()
            Return True

        Catch ex As Exception
            MsgBox("SERVER IS NOT CONNECTED, PLEASE CHECK THE NETWORK CONNECTION..!: " _
        & vbNewLine & ex.Message)
            Return False
            End

        End Try

    End Function


End Module

