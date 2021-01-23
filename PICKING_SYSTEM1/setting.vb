Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms.Form
Imports System
Public Class setting
    Dim reader As SqlDataReader
    Public myConn = "NOO"
    Private Sub setting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim connect_db = New connect()
        myConn = connect_db.conn()
        load_data_printer()
        load_data_HANDHELD()
        ComboBox2.Enabled = False
    End Sub
    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.ParentChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Panel2.Visible = False
        'MsgBox("999")
        Me.Close()
        'MsgBox("----")
        If Module1.check_page <> "NO_DATA" Then
            '   MsgBox("1")
            part_detail.Show()
            '  MsgBox("2")
        Else
            main.Show()
        End If
    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub
    Public Sub load_data_printer()
        ComboBox1.Items.Clear()
        Dim default_printer As String = "NO_DATA"
        Dim str_get As String = "select * from DEVICE_MASTER DM , DEVICE_SETTING_LOG DSL  where DM.DEVICE_IP = '" & main.ip_address & "' and DM.DEVICE_NAME = DSL.DEVICE_PAIR1 and DM.DEVICE_STATUS = '1' and DSL.STATUS = '1'"
        Dim command2 As SqlCommand = New SqlCommand(str_get, myConn)
        reader = command2.ExecuteReader()
        Dim DEVICE_NAME As String = "NO_DTAA"
        Dim DEVICE_PIN As String = "NO_DTAA"
        Dim DEVICE_BT As String = "NO_DTAA"
        Dim DEVICE_PAIR1 As String = "NO_DATA"
        Dim DEVICE_PAIR2 As String = "NO_DATA"
        Dim status As Integer = 0
        Dim id_History_device As String = "NO_DATA"
        If reader.Read() Then
            default_printer = reader("DEVICE_PAIR2").ToString()
        Else
            MsgBox("เครื่องยิงไม่มีข้อมูลในระบบ")
        End If
        reader.Close()
        Dim str As String = "select * from DEVICE_MASTER where DEVICE_TYPE = '2' and DEVICE_STATUS = '1'"
        Dim command As SqlCommand = New SqlCommand(str, myConn)
        reader = command.ExecuteReader()
        Dim i As Integer = 0
        Dim index As Integer = 0
        Do While reader.Read()
            ComboBox1.Items.Add(reader("DEVICE_NAME").ToString())
            If reader("DEVICE_NAME") = default_printer Then
                index = i
            Else
                i = i + 1
            End If
        Loop
        reader.Close()
        ComboBox1.SelectedIndex = index
    End Sub

    Public Sub load_data_HANDHELD()
        ComboBox2.Items.Clear()
        Dim str As String = "select * from DEVICE_MASTER where DEVICE_TYPE = '1' and DEVICE_STATUS = '1' and DEVICE_IP = '" & main.ip_address & "' "
        Dim command As SqlCommand = New SqlCommand(str, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read()
            ComboBox2.Items.Add(reader("DEVICE_NAME").ToString())
        Loop
        reader.Close()
        ComboBox2.SelectedIndex = 0

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim term_id As String = main.scan_terminal_id
        Dim str_get As String = "select * from DEVICE_MASTER DM , DEVICE_SETTING_LOG DSL  where DM.DEVICE_IP = '" & main.ip_address & "' and DM.DEVICE_NAME = DSL.DEVICE_PAIR1 and DM.DEVICE_STATUS = '1' and DSL.STATUS = '1'"
        Dim command As SqlCommand = New SqlCommand(str_get, myConn)
        reader = command.ExecuteReader()
        Dim DEVICE_NAME As String = "NO_DTAA"
        Dim DEVICE_PIN As String = "NO_DTAA"
        Dim DEVICE_BT As String = "NO_DTAA"
        Dim DEVICE_PAIR1 As String = "NO_DATA"
        Dim DEVICE_PAIR2 As String = "NO_DATA"
        Dim status As Integer = 0
        Dim id_History_device As String = "NO_DATA"
        If reader.Read() Then
            DEVICE_PAIR2 = reader("DEVICE_PAIR2").ToString()
            DEVICE_PAIR1 = reader("DEVICE_PAIR1").ToString()
            id_History_device = reader("DEVICE_ID").ToString()
        Else
            MsgBox("เครื่องยิงไม่มีข้อมูลในระบบ")
        End If
        reader.Close()
        Dim get_bt As String = "select * from DEVICE_MASTER  where  DEVICE_NAME = '" & ComboBox1.Text & "' and  DEVICE_STATUS = '1'"
        ' MsgBox(get_bt)
        Dim command2 As SqlCommand = New SqlCommand(get_bt, myConn)
        reader = command2.ExecuteReader()
        If reader.Read() Then
            DEVICE_BT = reader("DEVICE_BT").ToString()
            'MsgBox("DEVICE_BT = " & DEVICE_BT)
            main.pin_printer = reader("DEVICE_PIN").ToString()
            status = 1
        Else
            MsgBox("เครื่องยิงไม่มีข้อมูลในระบบ")
        End If
        reader.Close()
        If status = 1 Then
            main.scan_terminal_id = DEVICE_NAME
            main.pin_printer = DEVICE_PIN
            main.number_printter_bt = DEVICE_BT
            Update_History_device(id_History_device)
            Insert_History_device(id_History_device, term_id)
            MsgBox("SUCCESS")
            status = 0
            id_History_device = "NO_DATA"
        End If


    End Sub
    Public Sub Update_History_device(ByVal id_History_device As String)
        Try
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim str_update As String = "update DEVICE_SETTING_LOG set STATUS = '0'  , UPDATED_DATE = '" & date_now & "' , UPDATED_BY = '" & main.code_id_user.Substring(7) & "' where DEVICE_ID = '" & id_History_device & "'"
            ' MsgBox("str_update = " & str_update)
            Dim command As SqlCommand = New SqlCommand(str_update, myConn)
            reader = command.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("ERROR Update_History_device")
        End Try

    End Sub
    Public Sub Insert_History_device(ByVal id_History_device As String, ByVal term_id As String)
        Try
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim str_insert As String = "insert into DEVICE_SETTING_LOG  (DEVICE_PAIR1, DEVICE_PAIR2, STATUS , CREATED_DATE , CREATED_BY , UPDATED_DATE , UPDATED_BY) VALUES ('" & ComboBox2.Text & "' , '" & ComboBox1.Text & "' , '1' ,  '" & date_now & "' , '" & main.code_id_user.Substring(7) & "' , '" & date_now & "' , '" & main.code_id_user.Substring(7) & "')"
            'MsgBox("str_update = " & str_insert)
            Dim command As SqlCommand = New SqlCommand(str_insert, myConn)
            reader = command.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("ERROR Insert_History_device")
        End Try
    End Sub

    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class