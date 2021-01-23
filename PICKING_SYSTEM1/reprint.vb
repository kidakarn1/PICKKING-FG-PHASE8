Imports System.Linq
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports Bt.CommLib
Imports Bt
Imports System.Data.SqlClient

Public Class reprint
    Public myConn As SqlConnection
    Public myConn_Resive As SqlConnection
    Public REMAIN_ID As String = "NO_DATA"
    Public ID_table_detail As String = "NO_DATA"
    Public check_process As String = "NO_OK"
    Public remain_qty1 As Double = 0
    Dim g_index As Integer = 0
    Dim id_cut_stock_FW As String = "no_data"
    Dim path As String
    Dim a As Integer = 0
    Dim count_arr_fw As Integer = 0
    Dim count_fw_final As Integer = 0
    Dim frith As Integer = 0
    Public leng_scan_qty As Integer = 0
    Dim imagefile As String
    Public PD5 As scan_location
    Dim reader As SqlDataReader
    Dim dat As String = String.Empty
    Dim path1 As String
    Dim htlogfile As String
    Dim pclogfile As String = "logfile.csv"
    Dim data_final As String = "NOOOOO"
    Dim data_final_loop As String = "NOOOOO"
    Public Line As Select_Line
    Dim CodeType As String = "QR"
    Public c_check As String = "no_process"
    Dim m As String = "no-data"
    Dim status_alert_image As String = "NO_STATUS"
    Dim status_image As String = "NO_DATA"
    Dim g_update As Integer = 0 '
    Dim brak_loop As Integer = 0 '
    Dim j As Integer = 0
    Dim count_update_fw As Integer = 0
    Dim count_scan As Integer = 0
    Dim fa_use_total As Integer = 0
    Public length As Integer = 0
    Public Len_length_QR As Integer = 0
    Public check_scan As Integer = 0
    Public check_count__data As Integer = 0
    Public Const INTERNET_DEFAULT_FTP_PORT As Int32 = 21
    Public Const INTERNET_OPEN_TYPE_PRECONFIG As Int32 = 0
    Public Const INTERNET_OPEN_TYPE_DIRECT As Int32 = 1
    Public Const INTERNET_OPEN_TYPE_PROXY As Int32 = 3
    Public Const INTERNET_INVALID_PORT_NUMBER As Int32 = 0
    Public Const INTERNET_SERVICE_FTP As Int32 = 1
    Public Const INTERNET_SERVICE_GOPHER As Int32 = 2
    Public Const INTERNET_SERVICE_HTTP As Int32 = 3
    Public Const FTP_TRANSFER_TYPE_BINARY As Int64 = &H2
    Public Const FTP_TRANSFER_TYPE_ASCII As Int64 = &H1
    Public Const INTERNET_FLAG_NO_CACHE_WRITE As Int64 = &H4000000
    Public Const INTERNET_FLAG_RELOAD As Int64 = &H80000000UI
    Public Const INTERNET_FLAG_KEEP_CONNECTION As Int64 = &H400000
    Public Const INTERNET_FLAG_MULTIPART As Int64 = &H200000
    Public Const INTERNET_FLAG_PASSIVE As Int64 = &H8000000
    Public Const FILE_ATTRIBUTE_READONLY As Int64 = &H1
    Public Const FILE_ATTRIBUTE_HIDDEN As Int64 = &H2
    Public Const FILE_ATTRIBUTE_SYSTEM As Int64 = &H4
    Public Const FILE_ATTRIBUTE_DIRECTORY As Int64 = &H10
    Public Const FILE_ATTRIBUTE_ARCHIVE As Int64 = &H20
    Public Const FILE_ATTRIBUTE_NORMAL As Int64 = &H80
    Public Const FILE_ATTRIBUTE_TEMPORARY As Int64 = &H100
    Public Const FILE_ATTRIBUTE_COMPRESSED As Int64 = &H800
    Public Const FILE_ATTRIBUTE_OFFLINE As Int64 = &H1000
    ' Constant to use with coredll.dll
    Public Const WAIT_OBJECT_0 As Int32 = &H0
    ' The constant that is used with printing data
    Public Const STX As [Byte] = &H2
    Public Const ETX As [Byte] = &H3
    Public Const DLE As [Byte] = &H10
    Public Const SYN As [Byte] = &H16
    Public Const ENQ As [Byte] = &H5
    Public Const ACK As [Byte] = &H6
    Public Const NAK As [Byte] = &H15
    Public Const ESC As [Byte] = &H1B
    Public Const LF As [Byte] = &HA

    Private Sub TextBox1_Text_key_down(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub
    Private Sub Button5_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub user_id_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles user_id.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                password.Focus()
        End Select
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        Dim strCommand12345678 As String = "select count(su.su_id) as c_id  , su.sug_id as per from sys_users su  , sys_user_groups sug  where su.emp_id = '" & user_id.Text & "' and su.sys_pass = '" & password.Text & "' and su.sug_id  = sug.sug_id and su.enable = '1' GROUP BY su.sug_id"
        ' MsgBox(strCommand12345678)
        Dim cmd As SqlCommand = New SqlCommand(strCommand12345678, myConn)
        reader = cmd.ExecuteReader()
        If reader.Read() Then
            If reader("c_id").ToString() = "1" And reader("per").ToString() <> "3" Then
                Module1.user_reprint = user_id.Text()
                Panel5.Visible = True
                Panel4.Visible = False
                TextBox1.Focus()
            Else
                MsgBox("คุณไม่มีสิทธิ์")
            End If
        Else
            MsgBox("คุณไม่มีสิทธิ์")
        End If
        reader.Close()
    End Sub
    Public Function insert_log(ByVal bef_qty As String, ByVal status As String, ByVal lot As String, ByVal seq As String, ByVal item_cd As String)
        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyy-MM-dd HH:mm:ss"
        Dim date_now = time.ToString(format)
        Dim str As String = "insert into sys_logs_reprint (reprint_by , reprint_date , reprint_aft , reprint_bef , qr_read , status , reprint_lot , reprint_seq , reprint_item_cd) values('" & user_id.Text & "' , '" & date_now & "' , '" & TextBox2.Text & "' , '" & bef_qty & "' , '" & TextBox1.Text & "' , '" & status & "' , '" & lot & "' , '" & seq & "' , '" & item_cd & "')"
        'MsgBox("str = " & str)
        Dim command2 As SqlCommand = New SqlCommand(str, myConn)
        reader = command2.ExecuteReader()
        reader.Close()
        Return 0
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim L_data = Len(TextBox1.Text)
        If L_data = "62" Then
            Dim re_qty As String = TextBox2.Text
            Dim re_qty_number As Double = 0.0
            re_qty_number = CDbl(Val(re_qty))
            Dim number_re_qty As Integer = Len(re_qty)
            Dim new_qr As String = TextBox1.Text.Substring(0, 50)
            Dim qty_old As String = TextBox1.Text.Substring(51, 8)
            Dim SEQ As String = TextBox1.Text.Substring(59)
            Dim charArray_re_qty() As Char = re_qty.ToCharArray
            Dim charArray_old_qty() As Char = qty_old.ToCharArray
            Dim key As Integer = 0
            Dim RESULT_QTY As String = ""
            Dim regit As Integer = 9 - number_re_qty
            For i As Integer = 0 To 8 Step +1
                Dim data_rigit As String = ""
                If i = regit Then
                    data_rigit = charArray_re_qty(key)
                    regit = regit + 1
                    key = key + 1
                Else
                    data_rigit = "0"
                End If
                RESULT_QTY &= data_rigit
            Next
            Dim new_qr_re_print As String = new_qr & RESULT_QTY & SEQ
            Dim sp_qr = new_qr_re_print.Split("               ")
            Dim data_arr As String = sp_qr(0)
            Dim part_no_detail As String = data_arr.Substring(12)
            Dim str As String = "select * from sup_work_plan_supply_dev where ITEM_CD = '" & part_no_detail & "'"
            'MsgBox(str)
            Dim cmd As SqlCommand = New SqlCommand(str, myConn)
            reader = cmd.ExecuteReader()
            Dim part_name_detail As String = "NOATA"
            Dim Model_detail As String = "NOATA"
            Dim loc_detail As String = "NOATA"
            If reader.Read() Then
                part_name_detail = reader("ITEM_NAME").ToString()
                Model_detail = reader("MODEL").ToString()
                loc_detail = reader("LOCATION_PART").ToString()
            Else
                MsgBox("NODATA")
            End If
            reader.Close()
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim time_detail As DateTime = DateTime.Now
            Dim format_time_detail As String = "HH:mm:ss"
            Dim now_time_detail = time_detail.ToString(format_time_detail)
            Dim date_detail As DateTime = DateTime.Now
            Dim format_date_detail As String = "dd-MM-yyyy"
            Dim now_date_detail = date_detail.ToString(format_date_detail)
            Dim reprint_qty As String = RESULT_QTY
            Dim user_detail As String = Module1.user_reprint
            Dim qr_detail_reprint As String = new_qr_re_print
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            'MsgBox(main.number_printter_bt)
            stInfoSet1.addr = main.number_printter_bt
            Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet.addr = main.number_printter_bt
            Dim pin As StringBuilder = New StringBuilder("0000")
            Dim pin1 As StringBuilder = New StringBuilder("0000")
            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            stInfoSet1.addr = main.number_printter_bt
            M_reprint = "WEB_POST"
            Dim pinlen As UInt32 = CType(pin.Length, UInt32)
            Dim n_old As Double = 0.0
            n_old = CDbl(Val(TextBox1.Text.Substring(51, 8)))
            Dim old_qty As String = n_old
            Dim check = check_reprint("62", old_qty, TextBox1.Text, re_qty_number)
            If check = "SUCCESS" Then
loop_check_open_bt_wp:
                If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
                    Panel1.Visible = False
                    alert_open_printer.Visible = False
                    Dim PO As String = TextBox1.Text.Substring(2, 10)
                    Dim seq_text As String = TextBox1.Text.Substring(59, 3)
                    Dim textbox_split = TextBox1.Text.Split(" ")
                    Dim item_cd As String = textbox_split(0)
                    Dim data_item_cd As String = item_cd.Substring(12)
                    insert_log(old_qty, "1", PO, seq_text, data_item_cd)
                    Dim status = Api_update_qty(part_no_detail, PO, seq_text, TextBox2.Text)
                    Update_webpost(part_no_detail, PO, seq_text, TextBox2.Text)
                    update_remain_sup_pick_detail(part_no_detail, PO, seq_text, TextBox2.Text)
                    If status = "SUCCESS" Then
                        Bluetooth_Reprint(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, Model_detail, re_qty_number, loc_detail, user_detail, now_date_detail, now_time_detail, new_qr_re_print, SEQ)
                    End If
                Else
                    'MsgBox("กรุณาเปิดเครื่องปริ้น")
                    Panel1.Visible = True
                    alert_open_printer.Visible = True
                    GoTo loop_check_open_bt_wp
                End If
            ElseIf check = "Run_out" Then
                Panel1.Visible = True
                image_loss.Visible = True
                alert()
                cerrent.Focus()
            ElseIf check = "FAILL" Then
                Panel1.Visible = True
                alert_qty.Visible = True
                status_image = "Not_enough_qty"
                alert()
                cerrent.Focus()
                ' MsgBox("QTY ใน stock ไม่เพียงพอต่อการ reprint ")
            ElseIf check = "NO_DATA" Then
                Panel1.Visible = True
                alert_no_qr_code.Visible = True
                status_image = "No_qr_code"
                alert()
                cerrent.Focus()
            End If
            ' MsgBox(new_qr_re_print & " ===>length => " & Len(new_qr_re_print))
        ElseIf L_data = "103" Then
            Dim re_qty As String = TextBox2.Text
            Dim re_qty_number As Double = 0.0
            re_qty_number = CDbl(Val(re_qty))
            Dim number_re_qty As Integer = Len(re_qty)
            Dim new_qr As String = TextBox1.Text.Substring(0, 52)
            Dim qty_old As String = TextBox1.Text.Substring(51, 8)
            Dim qty_old2 As String = TextBox1.Text.Substring(58)
            Dim charArray_re_qty() As Char = re_qty.ToCharArray
            Dim charArray_old_qty() As Char = qty_old.ToCharArray
            Dim key As Integer = 0
            Dim RESULT_QTY As String = ""
            Dim regit As Integer = 7 - number_re_qty
            Dim n As Integer = 0
            For i As Integer = 1 To 6 Step +1
                Dim data_rigit As String = ""
                If i = regit Then
                    data_rigit = charArray_re_qty(key)
                    regit = regit + 1
                    key = key + 1
                Else
                    data_rigit = " "
                    n = n + 1
                End If
                RESULT_QTY &= data_rigit
            Next
            ' MsgBox("====>" & n & "<========" & RESULT_QTY)
            Dim new_qr_re_print As String = new_qr & RESULT_QTY & qty_old2
            ' MsgBox(new_qr_re_print & " ===>length => " & Len(new_qr_re_print))
            '  MsgBox(TextBox1.Text.Substring(52, 6))
            Dim sp_qr = new_qr_re_print.Split(" ")
            Dim data_arr As String = sp_qr(0)
            Dim part_no_detail As String = data_arr.Substring(19)

            Dim check_null As Integer = 0
            Dim char_array() As Char = TextBox1.Text.ToCharArray
            For i As Integer = 17 To 55 Step +1
                Dim rigit As String = char_array(i)
                If rigit = " " Then
                    check_null = check_null + 1
                End If
            Next
            Dim cut_rigit = TextBox1.Text.Split(" ")
            Dim b As String = cut_rigit(check_null)
            '  MsgBox("check_null = " & check_null)
            ' MsgBox("cut_rigit = " & cut_rigit(check_null))
            ' MsgBox("b = " & b)
            Dim SEQ As String = TextBox1.Text.Substring(16, 3) 'SEQ FA'
            Dim str As String = "select * from sup_work_plan_supply_dev where ITEM_CD = '" & part_no_detail & "'"
            '  MsgBox(str)
            Dim cmd As SqlCommand = New SqlCommand(str, myConn)
            reader = cmd.ExecuteReader()
            Dim part_name_detail As String = "NOATA"
            Dim Model_detail As String = "NOATA"
            Dim loc_detail As String = "NOATA"
            If reader.Read() Then
                part_name_detail = reader("ITEM_NAME").ToString()
                Model_detail = reader("MODEL").ToString()
                loc_detail = reader("LOCATION_PART").ToString()
            Else
                MsgBox("NODATA")
            End If
            reader.Close()
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)

            Dim time_detail As DateTime = DateTime.Now
            Dim format_time_detail As String = "HH:mm:ss"
            Dim now_time_detail = time_detail.ToString(format_time_detail)

            Dim date_detail As DateTime = DateTime.Now
            Dim format_date_detail As String = "dd-MM-yyyy"
            Dim now_date_detail = date_detail.ToString(format_date_detail)
            Dim reprint_qty As String = RESULT_QTY
            Dim user_detail As String = Module1.user_reprint
            Dim qr_detail_reprint As String = new_qr_re_print
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet.addr = main.number_printter_bt
            Dim pin As StringBuilder = New StringBuilder("0000")
            Dim pin1 As StringBuilder = New StringBuilder("0000")
            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            stInfoSet1.addr = "a066109719bd"
            M_reprint = "WEB_POST"
            Dim pinlen As UInt32 = CType(pin.Length, UInt32)
            M_reprint = "FW"
            Dim old_qty As String = TextBox1.Text.Substring(52, 6)
            Dim check = check_reprint("103", Trim(old_qty), TextBox1.Text, re_qty_number)
            'MsgBox(check)
            If check = "SUCCESS" Then
loop_check_open_bt:
                If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
                    Panel1.Visible = False
                    alert_open_printer.Visible = False
                    Dim old2 As String = TextBox1.Text.Substring(58)
                    Dim data = old2.Split(" ")
                    Dim lot_fa As String = data(0)
                    Dim full_text = TextBox1.Text.Split(" ")
                    Dim data_item_cd As String = full_text(0)
                    Dim item_cd = data_item_cd.Substring(19)
                    Dim plan_seq As String = TextBox1.Text.Substring(16, 3)
                    Dim lot_sep As String = TextBox1.Text.Substring(58, 4)
                    Dim tag_number As String = TextBox1.Text.Substring(100, 3)
                    Dim tag_seq As String = plan_seq + lot_sep + tag_number
                    insert_log(Trim(old_qty), "0", lot_fa, tag_seq, item_cd)
                    FW_UPDATE(part_no_detail, lot_sep)
                    update_remain_sup_pick_detail(part_no_detail, lot_fa, tag_seq, TextBox2.Text)
                    Bluetooth_Reprint(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, Model_detail, re_qty_number, loc_detail, user_detail, now_date_detail, now_time_detail, new_qr_re_print, SEQ)
                Else
                    ' MsgBox("connect faill")
                    ' MsgBox("กรุณาเปิดเครื่องปริ้น")
                    Panel1.Visible = True
                    alert_open_printer.Visible = True
                    GoTo loop_check_open_bt
                End If
            ElseIf check = "Run_out" Then
                Panel1.Visible = True
                image_loss.Visible = True
                alert()
                cerrent.Focus()
            ElseIf check = "FAILL" Then
                Panel1.Visible = True
                alert_qty.Visible = True
                status_image = "Not_enough_qty"
                alert()
                cerrent.Focus()
                ' MsgBox("QTY ใน stock ไม่เพียงพอต่อการ reprint ")
            ElseIf check = "NO_DATA" Then
                Panel1.Visible = True
                alert_no_qr_code.Visible = True
                status_image = "No_qr_code"
                alert()
                cerrent.Focus()
            End If
        Else
            ' MsgBox("QR CODE ไม่ถูกต้อง")
            Panel1.Visible = True
            alert_no_qr_code.Visible = True
            status_image = "No_qr_code"
            alert()
            cerrent.Focus()
        End If
    End Sub
    Public Function check_reprint(ByVal l_size As String, ByVal old_qty As String, ByVal textbox As String, ByVal new_qty As String )

        If l_size = "62" Then
            Dim PO As String = textbox.Substring(2, 10)
            Dim textbox_split = textbox.Split(" ")
            Dim item_cd As String = textbox_split(0)
            Dim data_item_cd As String = item_cd.Substring(12)
            Dim sql As String = "select item_cd , qty , id ,PUCH_ODR_CD ,com_flg from sup_frith_in_out where item_cd ='" & data_item_cd & "' and  PUCH_ODR_CD ='" & PO & "'  "
            Dim qty_double As Double = 0.0
            Dim new_qty_double As Double = 0.0
            Dim cmd As SqlCommand = New SqlCommand(sql, myConn)
            reader = cmd.ExecuteReader()
            Dim com_flg As String = "NO_DATA"
            If reader.Read() Then
                Dim QTY_STOCK = reader("qty").ToString()
                qty_double = CDbl(Val(reader("qty").ToString))
                new_qty_double = CDbl(Val(new_qty))
                com_flg = reader("com_flg").ToString
                reader.Close()

                If com_flg = "1" Then
                    Return "Run_out"
                End If

                If qty_double >= new_qty_double Then
                    Return "SUCCESS"
                Else
                    Return "FAILL"
                End If
            Else
                reader.Close()
                Return "NO_DATA"
            End If

        ElseIf l_size = "103" Then
            Dim qty_double As Double = 0.0
            Dim new_qty_double As Double = 0.0
            Dim old2 As String = textbox.Substring(58)
            Dim data = old2.Split(" ")
            Dim lot_fa As String = data(0)
            Dim full_text = textbox.Split(" ")
            Dim data_item_cd As String = full_text(0)
            Dim item_cd = data_item_cd.Substring(19)
            Dim sql As String = "select fa_item_cd , fa_lot , fa_total  , fa_com_flg from sup_frith_in_out_fa where fa_item_cd ='" & item_cd & "' and  fa_lot ='" & lot_fa & "'"
            Dim cmd As SqlCommand = New SqlCommand(sql, myConn)
            reader = cmd.ExecuteReader()
            Dim fa_com_flg = "NO_DATA"
            If reader.Read() Then
                qty_double = CDbl(Val(reader("fa_total").ToString))
                new_qty_double = CDbl(Val(new_qty))
                Dim QTY_STOCK = reader("fa_total").ToString()
                fa_com_flg = reader("fa_com_flg").ToString()
                reader.Close()
                ' MsgBox("qty_double = " & qty_double)
                'MsgBox("new_qty_double = " & new_qty_double)
                If fa_com_flg = "1" Then
                    Return "Run_out"
                End If
                If qty_double >= new_qty_double Then
                    Return "SUCCESS"
                Else
                    Return "FAILL"
                End If
            Else
                reader.Close()
                Return "NO_DATA"
            End If
        End If
        Return 0
    End Function
    Private Function Bluetooth_Connect_MB200i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32) As [Boolean]
        Dim bRet As [Boolean] = False
        Dim ret As Int32 = 0
        Dim disp As [String] = ""


        Try
            ret = Bluetooth.btBluetoothOpen()
            If ret <> LibDef.BT_OK Then
                'disp = "btBluetoothOpen error ret[" & ret & "]"
                'MessageBox.Show(disp, "Error")
                'Return bRet
            End If

            ret = Bluetooth.btBluetoothPairing(stInfoSet, pinlen, pin)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothPairing error ret[" & ret & "]"
               ' MessageBox.Show(disp, "Error")
                Return bRet
            End If

            ret = Bluetooth.btBluetoothSPPConnect(stInfoSet, 30000)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPConnect error ret[" & ret & "]"
                ' MessageBox.Show(disp, "Error")
                Return bRet
            End If

            bRet = True
            Return bRet
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            Return bRet
        Finally
        End Try
    End Function
    Private Sub Bluetooth_Reprint(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32, ByVal part_number As String, ByVal part_name As String, ByVal part_model As String, ByVal reprint_qty As String, ByVal loc_detail As String, ByVal user_detail As String, ByVal date_detail As String, ByVal time_detail As String, ByVal qr_detail_remain As String, ByVal seq As String)
        Dim ret As Int32 = 0
        Dim disp As [String] = ""

        Dim sbBuf As New StringBuilder("")
        Dim ssizeGet As UInt32 = 0
        Dim bBuf As [Byte]() = New Byte() {}


        Dim rsizeGet As UInt32 = 0
        Dim bBufGet As [Byte]() = New [Byte](4094) {}

        Try
            bBuf = New [Byte](4094) {}
            Dim bBufWork As [Byte]() = New [Byte]() {}
            Dim bESC As [Byte]() = New [Byte](0) {ESC}
            Dim bSTX As [Byte]() = New [Byte](0) {STX}
            Dim bETX As [Byte]() = New [Byte](0) {ETX}
            Dim bLF As [Byte]() = New [Byte](0) {LF}
            Dim b00 As [Byte]() = New [Byte](0) {&H0}
            Dim b30 As [Byte]() = New [Byte](0) {&H30}
            Dim len As Int32 = 0


            ' Receipt mode
            bSTX.CopyTo(bBuf, len)
            len = len + bSTX.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("A")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("PG33A1010112800384+000+000+00+00+00005000") '// Label
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H40")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Part No : " & part_number)

            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V140")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H20")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Reprint Tag")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length





            '// Human readable 2
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H90")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Part Name : " & part_name)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H140")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Model : " & part_model)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length


            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V550")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H190")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Reprint QTY : " & reprint_qty)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V550")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H240")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Location : " & loc_detail)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V380")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H680")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & time_detail)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V550")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H290")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)

            If M_reprint = "WEB_POST" Then
                Dim PO = qr_detail_remain.Substring(2, 10)

                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "PO No: " & PO)
            ElseIf M_reprint = "FW" Then
                Dim LOT = qr_detail_remain.Substring(58, 4)
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "LOT : " & LOT)
            End If
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length



            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V550")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H340")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)


            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "SEQ : " & seq)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length




            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H405")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)


            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Revised By : " & user_detail)

            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            '-'----------------'----------------'

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V470")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H405")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)


            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "DATE : " & date_detail & " Time : " & time_detail)

            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            If CodeType = "C128" Then
                ''// barcode
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H220")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("BG02060>G" & qr_detail_remain) '// code 128
                'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("BD101060" & data) '// code 39
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length


            Else
                '// QR code
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V735")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H220")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("2D30,M,04,0,0") '// qr setting
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("DS2," & qr_detail_remain) '// qr data
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length
            End If






            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("Q0001")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("Z")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bETX.CopyTo(bBuf, len)
            len = len + bETX.Length

            If SppSend(bBuf, ssizeGet) = False Then
                GoTo L_END1
            End If

            ' Response wait
            Dim printflg As [Boolean] = False
            While True
                Dim recvFlg As [Boolean] = False
                For i As Int32 = 0 To 9
                    ' Receive data
                    bBufGet = New [Byte](0) {}
                    If SppRecv(bBufGet, rsizeGet) = False Then
                        Continue For
                    End If
                    recvFlg = True
                    Exit For
                Next
                If recvFlg = False Then
                    Exit While
                End If

                If bBufGet(0) = ACK Then
                    ' SBPL transmission
                    bBuf = New [Byte]() {ENQ}
                    If SppSend(bBuf, ssizeGet) = False Then
                        GoTo L_END1
                    End If
                ElseIf bBufGet(0) = NAK Then
                    GoTo L_END1
                ElseIf bBufGet(0) = STX Then
                    ' SPBL reception
                    bBufGet = New [Byte](4094) {}
                    If SppRecv(bBufGet, rsizeGet) = False Then
                        GoTo L_END1
                    End If
                    If bBufGet(9) <> ETX Then
                        GoTo L_END1
                    End If
                    If bBufGet(2) = &H47 OrElse bBufGet(2) = &H48 OrElse bBufGet(2) = &H53 OrElse bBufGet(2) = &H54 Then
                        ' Printing is being performed, so wait for a short amount of time.
                        Thread.Sleep(200)
                        bBuf = New [Byte]() {ENQ}
                        If SppSend(bBuf, ssizeGet) = False Then
                            GoTo L_END1
                        End If
                        Continue While
                    ElseIf (bBufGet(2) <> &H0) AndAlso (bBufGet(2) <> &H1) AndAlso (bBufGet(2) <> &H41) AndAlso (bBufGet(2) <> &H42) AndAlso (bBufGet(2) <> &H4E) AndAlso (bBufGet(2) <> &H4D) Then
                        Exit While
                    End If
                    ' Print success
                    printflg = True
                    Exit While
                End If
            End While
            If printflg = True Then

                disp = "Printing Successfully."
                MessageBox.Show(disp, "Printing complete")

                '// Append scanlog file
                Dim sw As New System.IO.StreamWriter(htlogfile, True, System.Text.Encoding.GetEncoding("Shift_JIS"))

                Dim dtNow As DateTime = DateTime.Now
                sw.Write(DateTime.Now.ToString("dd/MM/yyyy,HH:mm:ss,"))
                sw.Write("TEST" & vbCrLf)
                sw.Close()

            End If

            Return
L_END1:
            ret = Bluetooth.btBluetoothSPPDisconnect()
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPDisconnect error ret[" & ret & "]"
                MessageBox.Show(disp, "Error")
                GoTo L_END2
            End If
L_END2:
            'ButtonF2.Enabled = True
            ret = Bluetooth.btBluetoothClose()
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothClose error ret[" & ret & "]"
                MessageBox.Show(disp, "Error")
                Return
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            MsgBox("Program error please contact PCS department")
        End Try

        remain_qty1 = 0

    End Sub
    Private Function SppRecv(ByRef bBuf As [Byte](), ByRef rsize As UInt32) As [Boolean]
        Dim bRet As [Boolean] = False
        Dim ret As Int32 = 0
        Dim disp As [String] = ""

        Dim dsizeSet As UInt32 = 0
        Dim rsizeGet As UInt32 = 0
        Dim pBufGet As IntPtr
        Dim bBufGet As [Byte]() = New [Byte]() {}

        Try
            Thread.Sleep(1000)
            Dim buflen As Int32 = bBuf.Length
            bBufGet = New [Byte](buflen - 1) {}
            pBufGet = Marshal.AllocCoTaskMem(Marshal.SizeOf(bBufGet))
            dsizeSet = CType(buflen, UInt32)
            ret = Bluetooth.btBluetoothSPPRecv(pBufGet, dsizeSet, rsizeGet)
            Marshal.Copy(pBufGet, bBufGet, 0, CType(rsizeGet, Int32))
            Marshal.FreeCoTaskMem(pBufGet)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPRecv error ret[" & ret & "]"
                MessageBox.Show(disp, "Error")
                Return bRet
            End If

            bBuf = bBufGet
            rsize = rsizeGet
            bRet = True
            Return bRet
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            Return bRet
        Finally
        End Try
    End Function
    Private Function SppSend(ByVal bBuf As [Byte](), ByRef ssize As UInt32) As [Boolean]
        Dim bRet As [Boolean] = False
        Dim ret As Int32 = 0
        Dim disp As [String] = ""

        Dim dsizeSet As UInt32 = 0
        Dim ssizeGet As UInt32 = 0
        Dim pBufSet As IntPtr

        Try
            dsizeSet = CType(bBuf.Length, UInt32)
            pBufSet = Marshal.AllocCoTaskMem(CType(dsizeSet, Int32))
            Marshal.Copy(bBuf, 0, pBufSet, CType(dsizeSet, Int32))
            ret = Bluetooth.btBluetoothSPPSend(pBufSet, dsizeSet, ssizeGet)
            Marshal.FreeCoTaskMem(pBufSet)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPSend error ret[" & ret & "]"
                MessageBox.Show(disp, "Error")
                Return bRet
            End If

            ssize = ssizeGet
            bRet = True
            Return bRet
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            Return bRet
        Finally
        End Try
    End Function

    Private Sub reprint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim connect_db = New connect()
            myConn = connect_db.conn()
        Finally
            Panel1.Visible = False
            alert_open_printer.Visible = False
            alert_qty.Visible = False
            Panel5.Hide()
            ' TextBox2.Enabled = False
            main.loader()
            Label1.Text = "QTY BEFORE : 0"
            Label2.Text = "QTY AFTER : 0"
            main.Panel2.Visible = False
            ' main.Panel2.Visible = False
            'main.loader()
        End Try
    End Sub

    Private Sub PictureBox6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Show()
        Me.Hide()
    End Sub

    Private Sub user_id_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles user_id.TextChanged, user_id.KeyDown

    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Label1.Text = "QTY BEFORE : 0"
        TextBox1.Text = ""
        TextBox2.Text = ""
        Label2.Text = "QTY AFTER : 0"
        TextBox1.Focus()
    End Sub

    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged

    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                ' MsgBox("001")
                If Len(TextBox1.Text) = "62" Then
                    ' MsgBox("1")
                    Dim n_old As Double = 0.0
                    n_old = CDbl(Val(TextBox1.Text.Substring(51, 8)))
                    Dim qty_old As String = n_old
                    'check_sock("62", qty_old)
                    Label1.Text = "QTY BEFORE :" & qty_old
                ElseIf Len(TextBox1.Text) = "103" Then
                    ' MsgBox("2")
                    Dim qty_old As String = TextBox1.Text.Substring(52, 6)
                    'check_sock("103", qty_old)
                    Label1.Text = "QTY BEFORE :" & Trim(qty_old)
                Else
                    Label1.Text = "QTY BEFORE : 0"
                End If
                TextBox2.Focus()
        End Select
    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.ParentChanged

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Label2.Text = "QTY AFTER : " & TextBox2.Text
    End Sub

    Private Sub Panel5_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel5.GotFocus

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub
    Public Function Api_update_qty(ByVal item_cd As String, ByVal PO As String, ByVal SEQ As String, ByVal qty As Integer)
        Dim result As String = "NO_DATA"
        Try
            'MsgBox("http://192.168.161.102/exp_api3party/Api_cut_stock_web_post/get_data_service?ITEM_CD=" & F_item_cd & "&PO=" & PO & "&SEQ=" & updated_seq & " &QTY=" & used_qty)
            result = Api.update_data("http://192.168.161.102/exp_api3party/Api_cut_stock_web_post/update_qty?ITEM_CD=" & item_cd & "&PO=" & PO & "&SEQ=" & SEQ & "&QTY=" & qty)
            'MsgBox("status = " & result)
            If result <> "NO_DATA" Then
                ' MsgBox("FALL UPDATE CUT_STOCK P_NICE = " & result)
            End If
        Catch ex As Exception
            MsgBox("FAILL cut_stock_FASYSTEM" & vbNewLine & ex.Message, "FAILL")
        End Try
        Return result
    End Function
    Public Function Update_webpost(ByVal item_cd As String, ByVal PO As String, ByVal SEQ As String, ByVal qty As Integer)
        Dim result As String = "NO_DATA"
        Try
            M_CHECK_TYPE = "WEB_POST"
            Dim result_api As String = Api.update_data("http://192.168.161.102/exp_api3party/API_SET_QTY_REPRINT_WEBPOST/get_data_service?ITEM_CD=" & item_cd & "&PUCH_ODR_CD=" & PO)
            'MsgBox("status = " & result)
            If result_api <> "SUCCESS" Then
                MsgBox("FALL UPDATE REPRINT  P_NICE = " & result)
            End If
        Catch ex As Exception
            MsgBox("FAILL " & vbNewLine & ex.Message, "FAILL")
        End Try
        Return result
    End Function
    Public Sub FW_UPDATE(ByVal item_cd As String, ByVal lot As String)
        Dim qty_databse As Double = 0.0
        Dim qty_databse_STring As String = ""
        M_CHECK_TYPE = "FW"
        Dim get_data As String = "select sum(fa_total) - sum(fa_use) as total_qty  ,  fa_total , fa_id , fa_use , fa_re from sup_frith_in_out_fa where fa_item_cd = '" & item_cd & "' and fa_lot = '" & lot & "' group by  fa_total , fa_id , fa_use , fa_re"
        Dim command2 As SqlCommand = New SqlCommand(get_data, myConn)
        reader = command2.ExecuteReader()
        Dim get_total As Double = 0.0
        Dim fa_total As Double = 0.0
        Dim com_flg As Integer = 0
        Dim fa_id As String = "NO_DATA"
        Dim update_qty As Double = 0.0
        Dim fa_re As String = "NO_DATA"
        Dim default_qty As String = "NO_DATA"
        Dim qty_update As Double = 0.0
        If reader.Read Then
            fa_id = reader("fa_id").ToString()
            fa_total = CDbl(Val(reader("fa_total").ToString()))
            update_qty = fa_total + (CDbl(Val(TextBox2.Text)))
            fa_re = (CDbl(Val(TextBox2.Text)))
            default_qty = fa_total - CDbl(Val(reader("fa_re").ToString()))
            qty_update = default_qty + (CDbl(Val(TextBox2.Text)))
            If update_qty > CDbl(Val(reader("fa_use").ToString())) Then
                com_flg = 0
            Else
                com_flg = 1
            End If
        End If
        reader.Close()
        qty_databse = update_qty
        qty_databse_STring = update_qty
        Dim update_stock_fa As String = "update sup_frith_in_out_fa set fa_total = '" & qty_update & "' , fa_com_flg = '" & com_flg & "'  , fa_re = '" & fa_re & "' where fa_id = '" & fa_id & "'"
        Dim cmd_stock_fa As SqlCommand = New SqlCommand(update_stock_fa, myConn)
        reader = cmd_stock_fa.ExecuteReader()
        reader.Close()
    End Sub

    Private Sub PictureBox4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click

    End Sub

    Private Sub cerrent_down(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cerrent.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                Panel1.Visible = False
                alert_open_printer.Visible = False
                alert_qty.Visible = False
                image_loss.Visible = False
        End Select
    End Sub
    Public Sub alert()
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
    End Sub
    Public Sub update_remain_sup_pick_detail(ByVal item_cd As String, ByVal PO As String, ByVal SEQ As String, ByVal qty As Integer)
        Dim update_qty As String = "update sup_scan_pick_detail set tag_remain_qty = '" & qty & "' where com_flg = '0' and  scan_lot = '" & PO & "' and item_cd = '" & item_cd & "' and tag_seq  = '" & SEQ & "'"
        Dim command_update As SqlCommand = New SqlCommand(update_qty, myConn)
        reader = command_update.ExecuteReader()
        reader.Close()
    End Sub
    Private Sub Panel3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel3.GotFocus

    End Sub

    Private Sub Panel4_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel4.GotFocus

    End Sub

    Private Sub Label7_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.ParentChanged

    End Sub
End Class