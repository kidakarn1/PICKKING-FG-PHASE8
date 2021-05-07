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
Public Class main
    Public scan_terminal_id = "NO_DATA" 'GetIPAddress()
    Public pin_printer = "NO_DATA" 'GetIPAddress()
    Public number_printter_bt = "NO_DATA" 'GetIPAddress()
    Public ip_address As String = GetIPAddress()
    'Dim myConn As SqlConnection
    Dim path As String
    Public Str As String
    Public count_emp_id As Integer
    Dim imagefile As String
    Public pd_of_user As String
    Public passToanofrm As String
    Public empToanofrm As String
    Public code_id_user As String = "NO"
    Public pd_user As String
    Public valSel As Array
    Public strData(,) As String
    Public miscData() As Object = {}
    Public ml As Integer = 0
    Public count_time As Integer = 0
    Public status As Integer = 0
    Dim re_data As ArrayList = New ArrayList()
    Public myConn = "NOO"
    Public count_net As Integer = 0
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
reconnect:
            If Api.check_net() = True Then
                Dim connect_db = New connect()
                myConn = connect_db.conn()
                set_data_handheld()
            Else
                MsgBox("กรุณา รอ Internet")
                GoTo reconnect
            End If
        Finally
            Panel1.Show()
            PictureBox8.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            'Panel1.Visible = True
            PictureBox4.Visible = False
            PictureBox20.Visible = False
            'PictureBox9.Visible = False
            Label7.Visible = False
            Me.emp_cd.Focus()
            ' get_image_user()
            Panel2.Visible = False
            setting.Visible = False
        End Try
    End Sub
    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub Label1_ParentChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.TextChanged
        Label2.Text = "DATE" + ": " + DateTime.Now.ToString("dd-MM-yyyy")
    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Panel1.Parent = PictureBox1
        Panel1.BackColor = Color.Transparent
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub

    Dim reader As SqlDataReader
    Dim dat As String = String.Empty

    Private Sub emp_cd_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles emp_cd.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                Str = ""
                Str = emp_cd.Text
                ' If Str.Length() = 5 Then
                query_user()
                'End If
        End Select
    End Sub

    Public Function show_code_id_user() As String
        Return code_id_user
    End Function

    Public Sub query_user()
        Dim strCommand As String = "NOOO"
        Try
            'ProgressBar1.Show()
            'Dim strCommand As String = "SELECT * FROM sys_users"
            strCommand = "SELECT * FROM sys_users WHERE emp_id = " & "'" & Str & "' and enable = '1' "
            ' MsgBox(strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            Do While reader.Read()
                dat = reader.Item(2) & " " & reader.Item(3)
                passToanofrm = dat
                'System.Console.WriteLine("===>" + reader.Item(1))'
                Module1.A_USER_ID = reader.Item(1)
                code_id_user = "CODE : " + reader.Item(1)
                empToanofrm = "Name:  " + reader("firstname").ToString & " " & reader("lastname").ToString
                Module1.Fullname = empToanofrm
                count_emp_id = 1
                pd_user = reader("dep_id").ToString()
                Module1.PHASE = reader("phase").ToString()
            Loop
            reader.Close()
            emp_cd.Text = ""
            If count_emp_id = 0 Then
                PictureBox8.Visible = True
                Dim stBuz As New Bt.LibDef.BT_BUZZER_PARAM()
                Dim stVib As New Bt.LibDef.BT_VIBRATOR_PARAM()
                Dim stLed As New Bt.LibDef.BT_LED_PARAM()
                stBuz.dwOn = 200
                stBuz.dwOff = 100
                stBuz.dwCount = 2
                stBuz.bVolume = 3
                stBuz.bTone = 1
                stVib.dwOn = 200
                stVib.dwOff = 100
                stVib.dwCount = 2
                stLed.dwOn = 200
                stLed.dwOff = 100
                stLed.dwCount = 2
                stLed.bColor = Bt.LibDef.BT_LED_MAGENTA
                Bt.SysLib.Device.btBuzzer(1, stBuz)
                Bt.SysLib.Device.btVibrator(1, stVib)
                Bt.SysLib.Device.btLED(1, stLed)
                TextBox1.Focus()
                'MessageBox.Show(strCommand)
            ElseIf count_emp_id = 1 Then
                Fullname = empToanofrm
                Panel1.Hide()
                Label4.Text = code_id_user
                Label5.Text = empToanofrm
                Label4.Visible = True
                Label5.Visible = True
                PictureBox1.Visible = True
                PictureBox2.Visible = True
                PictureBox3.Visible = True
                PictureBox4.Visible = True
                PictureBox18.Visible = True
                Label9.Visible = True
                Label8.Visible = True
                setting.Visible = True
                setting.Visible = True
                Label7.Visible = True
                PictureBox9.Visible = False
                Label11.Visible = False
                PictureBox19.Visible = True
                Label12.Visible = True
                PictureBox3.Visible = False
                Label6.Visible = False
                PictureBox20.Visible = False
                PictureBox2.Visible = False
                Label8.Visible = False
                PictureBox21.Visible = True
                Label13.Visible = True
                'get_image_user()
            End If
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status")
        End Try
    End Sub
    Public Sub p_img(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        PictureBox8.Visible = False
        Str = ""
        emp_cd.Text = ""
        emp_cd.Focus()
    End Sub
    Public Sub get_image_user()
        Dim test_img1 As img_user = New img_user()
        test_img1.Show()
    End Sub

    Public Function show_empToanofrm() As String
        Return empToanofrm
    End Function

    Private Sub Label1_ParentChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WebBrowser2_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        'WebBrowser2_DocumentCompleted = "https://gfycat.com/stickers/search/b%C3%ACnh+d%C6%B0%C6%A1ng"
    End Sub

    Private Sub ProgressBar1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        code_id_user = ""
        empToanofrm = ""
        Module1.Fullname = ""
        count_emp_id = ""
        pd_user = ""
        Panel1.Show()

    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Module1.MENU_ID = "1"
        Timer1.Enabled = True
        loader()
        Application.DoEvents()
        Try
            Dim PD As Select_PD = New Select_PD()
            PD.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("load faill")
        End Try

    End Sub

    Private Sub PictureBox4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        setting.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        pd_user = ""
        emp_cd.Text = ""
        count_emp_id = 0
        Panel1.Show()
        Panel1.Visible = True
        emp_cd.Focus()
        Me.emp_cd.Focus()
        PictureBox1.Visible = False
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox20.Visible = False
        'PictureBox9.Visible = False
        Label7.Visible = False
        dat = ""
        passToanofrm = dat
        'System.Console.WriteLine("===>" + reader.Item(1))'
        code_id_user = ""
        empToanofrm = ""
        Module1.Fullname = ""
        pd_user = ""
        Fullname = ""
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)

    End Sub

    Private Sub Panel3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub emp_cd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel3_GotFocus_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel3.GotFocus

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label4_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.ParentChanged

    End Sub


    Private Sub Label5_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.ParentChanged

    End Sub

    Private Sub Label3_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label7_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.ParentChanged

    End Sub

    Private Sub Label6_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.ParentChanged

    End Sub
    Public Sub loader()
        Panel2.Visible = True
    End Sub
    Private Sub PictureBox3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Module1.MENU_ID = "5"
        Timer1.Enabled = True
        loader()
        Application.DoEvents()
        Try
            Dim reprint As reprint = New reprint()
            reprint.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("error next page")
        End Try

    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.PictureBox1.Image = Image.FromFile("C:\Users\Me\Pictures\myanimatedimage.gif")
    End Sub

    Private Sub PictureBox10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ml += 1
        status += 5
recheck_net:
        If count_net = 5000 Then
            If Api.check_net <> True Then
                MsgBox("อินเตอร์เน็ตไม่เสถียร กรุณา กด ENT เพื่อ รอ INTERNET")
                GoTo recheck_net
            Else
                count_net = 0
            End If
        Else
            count_net += 1
        End If
        If ml <= 1 Then
            PictureBox10.Visible = True
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 2 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = True
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 3 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = True
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 4 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = True
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 5 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = True
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 6 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = True
            PictureBox16.Visible = False
            PictureBox17.Visible = False
        ElseIf ml = 7 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = True
            PictureBox17.Visible = False
        ElseIf ml = 8 Then
            PictureBox10.Visible = False
            PictureBox11.Visible = False
            PictureBox12.Visible = False
            PictureBox13.Visible = False
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = True
            ml = 0
        End If

    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel2_GotFocus_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel2.GotFocus

    End Sub
    Private Function GetIPAddress()
         If Api.check_net() = True Then 
            Dim strHostName As String

            Dim strIPAddress As String

            strHostName = System.Net.Dns.GetHostName()

            strIPAddress = System.Net.Dns.GetHostByName(strHostName).AddressList(1).ToString()
            'MsgBox("re_turn strIPAddress" & strIPAddress)

            Return strIPAddress
        End If
        Return 0
        'MessageBox.Show("Host Name: " & strHostName & "; IP Address: " & strIPAddress)
    End Function
    Public Sub set_data_handheld()

        Dim ip_han As String = GetIPAddress()
        Dim str_get As String = "select * from DEVICE_MASTER DM , DEVICE_SETTING_LOG DSL  where DM.DEVICE_IP = '" & ip_han & "' and DM.DEVICE_NAME = DSL.DEVICE_PAIR1 and DM.DEVICE_STATUS = '1' and DSL.STATUS = '1'"
        Dim command As SqlCommand = New SqlCommand(str_get, myConn)
        reader = command.ExecuteReader()
        Dim DEVICE_NAME As String = "NO_DTAA"
        Dim DEVICE_PIN As String = "NO_DTAA"
        Dim DEVICE_BT As String = "NO_DTAA"
        Dim DEVICE_PAIR1 As String = "NO_DATA"
        Dim DEVICE_PAIR2 As String = "NO_DATA"
        If reader.Read() Then
            DEVICE_PAIR2 = reader("DEVICE_PAIR2").ToString()
            DEVICE_PAIR1 = reader("DEVICE_PAIR1").ToString()
        Else
            MsgBox("เครื่องยิงไม่มีข้อมูลในระบบ")
        End If
        reader.Close()
        Dim get_bt As String = "select * from DEVICE_MASTER  where  DEVICE_NAME = '" & DEVICE_PAIR2 & "' and  DEVICE_STATUS = '1'"
        ' MsgBox(get_bt)
        Dim command2 As SqlCommand = New SqlCommand(get_bt, myConn)
        reader = command2.ExecuteReader()
        If reader.Read() Then
            DEVICE_BT = reader("DEVICE_BT").ToString()
            pin_printer = reader("DEVICE_PIN").ToString()
        Else
            MsgBox("เครื่องยิงไม่มีข้อมูลในระบบ")
        End If
        reader.Close()
        scan_terminal_id = DEVICE_PAIR1
        pin_printer = DEVICE_PIN
        number_printter_bt = DEVICE_BT
        ' MsgBox("number_printter_bt" & number_printter_bt)
    End Sub
    Private Sub setting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles setting.Click
        Try
            Timer1.Enabled = True
            loader()
            Application.DoEvents()
            Dim setting As setting = New setting()
            setting.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("error next page setting")
        End Try
    End Sub
    Private Sub PictureBox2_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Try
            Module1.MENU_ID = "6"
            Timer1.Enabled = True
            loader()
            Application.DoEvents()
            Dim select_pick_add As select_pick_add = New select_pick_add()
            select_pick_add.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("error next page PICK_ADD")
        End Try
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        Try
            Module1.MENU_ID = "4"
            Timer1.Enabled = True
            loader()
            Application.DoEvents()
            Dim return_part As return_part = New return_part()
            return_part.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("error next page return_part")
        End Try
    End Sub
    Private Sub Label11_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.ParentChanged

    End Sub

    Private Sub PictureBox18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        Label9.Visible = False
        PictureBox1.Visible = False
        PictureBox21.Visible = False
        Label13.Visible = False
        Label12.Visible = False
        Label8.Visible = True
        PictureBox2.Visible = True
        PictureBox18.Visible = False
        setting.Visible = False
        PictureBox9.Visible = True
        Label11.Visible = True
        PictureBox19.Visible = False
        Label12.Visible = True
        PictureBox3.Visible = True
        Label6.Visible = True
        PictureBox20.Visible = True
        PictureBox22.Visible = False
        PictureBox19.Visible = True
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        PictureBox21.Visible = True
        Label13.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        PictureBox1.Visible = True
        PictureBox2.Visible = True
        PictureBox3.Visible = True
        PictureBox4.Visible = True
        PictureBox18.Visible = True
        Label9.Visible = True
        Label8.Visible = True
        setting.Visible = True
        setting.Visible = True
        Label7.Visible = True
        PictureBox9.Visible = False
        Label11.Visible = False
        PictureBox19.Visible = False
        Label12.Visible = False
        PictureBox3.Visible = False
        Label6.Visible = False
        PictureBox20.Visible = False
        PictureBox22.Visible = True
    End Sub

    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        Try
            Module1.MENU_ID = "3"
            Timer1.Enabled = True
            loader()
            Application.DoEvents()
            Dim cut_ng As cut_ng = New cut_ng()
            cut_ng.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("error next page cut_ng")
        End Try
    End Sub

    Private Sub PictureBox21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox21.Click
        Module1.MENU_ID = "2"
        Timer1.Enabled = True
        loader()
        Application.DoEvents()
        Dim Select_plan_fg As Select_plan_fg = New Select_plan_fg()
        Select_plan_fg.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox22.Click
        Rework_FG.Show()
        Me.Hide()

    End Sub
End Class