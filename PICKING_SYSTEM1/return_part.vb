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
Imports System.Linq
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Threading
Imports System.Net.Sockets
Imports Bt.CommLib
Imports Bt

Public Class return_part
    Public myConn As SqlConnection
    Dim reader As SqlDataReader
    Dim PO_LOT As String = "NO_DATA"
    Public SEQ As String = "NO_DATA"
    Dim CodeType As String = "QR"
    Public tag_read As String = "NO_DATA"
    Public model As String = "NO_DATA"
    Public M_CHECK_TYPE As String = "NO_DATA"
    Dim htlogfile As String
    Public remain_qty1 As Double = 0
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
    Private Sub select_pick_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
            Dim connect_db = New connect()
            myConn = connect_db.conn()
        Finally
            alert_open_printer.Visible = False
            TextBox1.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            Panel4.Visible = False
            alert_no_tag_supply.Visible = False
            alert_not_enogth_return.Visible = False
            TextBox2.Focus()
        End Try

    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Panel2.Visible = False
        Me.Close()
        main.Show()
    End Sub

    Private Sub scan_supply(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                If Len(TextBox2.Text) <> "62" And Len(TextBox2.Text) <> "103" Then
                    Dim data As String = "NODATA"
                    data = TextBox2.Text.Substring(57)
                    Dim data_split = data.Split(" ")
                    Dim data_supply = TextBox2.Text.Split()
                    Dim id As String = data_supply(21)
                    Dim sql As String = "select * from sup_pick_log , sup_scan_pick_detail  where sup_pick_log.id = '" & id & "' and sup_pick_log.REF_ID = sup_scan_pick_detail.id"
                    Dim command As SqlCommand = New SqlCommand(sql, myConn)
                    reader = command.ExecuteReader()
                    'MsgBox("SQL = " & sql)
                    Dim part_no As String = "NO_DATA"
                    Dim part_name As String = "NO_DATA"
                    Dim location As String = "NO_DATA"
                    Dim wi As String = "NO_DATA"
                    Dim ref_id As String = "NODATA"

                    If reader.Read() Then
                        part_no = reader("item_cd").ToString()
                        PO_LOT = reader("scan_lot").ToString()
                        wi = reader("wi").ToString()
                        ref_id = reader("ref_id").ToString()
                        SEQ = reader("tag_seq").ToString()
                        Module1.M_SEQ_PRINT = SEQ.Substring(0, 3)
                        tag_read = reader("tag_readed").ToString()
                    End If
                    reader.Close()

                    Dim get_data_show As String = "select * from sup_work_plan_supply_dev where ITEM_CD = '" & part_no & "' and WI = '" & wi & "'"
                    Dim command2 As SqlCommand = New SqlCommand(get_data_show, myConn)
                    reader = command2.ExecuteReader()
                    If reader.Read() Then
                        part_name = reader("ITEM_NAME").ToString()
                        location = reader("LOCATION_PART").ToString()
                        model = reader("MODEL").ToString()
                    End If
                    reader.Close()
                    TextBox6.Text = PO_LOT
                    TextBox1.Text = part_no
                    TextBox4.Text = part_name
                    TextBox5.Text = location
                    TextBox3.Focus()
                Else
                    'MsgBox("NO TAG SUPPLY")
                    Panel4.Visible = True
                    alert_no_tag_supply.Visible = True
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
                    cerrent.Focus()
                End If
        End Select
    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.ParentChanged

    End Sub

    Private Sub Label3_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.ParentChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox2.Focus()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim qr_detail_remain As String = "NO_DATA"
        Dim qty_databse As Double = 0.0
        Dim qty_databse_STring As String = "NO___"
        Dim status As Integer = 0
        Dim qr_supply = TextBox2.Text.Split(" ")
        Dim QTY_to_first = qr_supply(19) 'QTY ที่เอามาครั้งแรก จาก supply'

        MsgBox("QTY_to_first  = " & QTY_to_first)
        MsgBox("TextBox3.Text  = " & TextBox3.Text)
        If CDbl(Val(TextBox3.Text)) <= CDbl(Val(QTY_to_first)) Then
            If PO_LOT.Length = "4" And PO_LOT <> "NO_DATA" Then
                M_CHECK_TYPE = "FW"
                Dim get_data As String = "select sum(fa_total) - sum(fa_use) as total_qty  ,  fa_total , fa_id , fa_use  from sup_frith_in_out_fa where fa_item_cd = '" & TextBox1.Text & "' and fa_lot = '" & TextBox6.Text & "' group by  fa_total , fa_id , fa_use"
                Dim command2 As SqlCommand = New SqlCommand(get_data, myConn)
                reader = command2.ExecuteReader()
                Dim get_total As Double = 0.0
                Dim fa_total As Double = 0.0
                Dim com_flg As Integer = 0
                Dim fa_id As String = "NO_DATA"
                Dim update_qty As Double = 0.0
                If reader.Read Then
                    fa_id = reader("fa_id").ToString()
                    fa_total = CDbl(Val(reader("fa_total").ToString()))
                    update_qty = fa_total + (CDbl(Val(TextBox3.Text)))
                    If update_qty > CDbl(Val(reader("fa_use").ToString())) Then
                        com_flg = 0
                    Else
                        com_flg = 1
                    End If
                End If
                reader.Close()
                qty_databse = update_qty
                qty_databse_STring = update_qty
                Dim update_stock_fa As String = "update sup_frith_in_out_fa set fa_total = '" & update_qty & "' , fa_com_flg = '" & com_flg & "' where fa_id = '" & fa_id & "'"
                Dim cmd_stock_fa As SqlCommand = New SqlCommand(update_stock_fa, myConn)
                reader = cmd_stock_fa.ExecuteReader()
                reader.Close()
                MsgBox("UPDATE OK")
                status = 1
            ElseIf PO_LOT.Length = "10" And PO_LOT <> "NO_DATA" Then
                M_CHECK_TYPE = "WEB_POST"
                Dim qty As Double = 0.0
                Dim id As String = "NO_DATA"
                Dim get_data As String = "select * from sup_frith_in_out where item_cd = '" & TextBox1.Text & "' and PUCH_ODR_CD = '" & TextBox6.Text & "'"
                Dim command2 As SqlCommand = New SqlCommand(get_data, myConn)
                reader = command2.ExecuteReader()
                If reader.Read Then
                    qty = CDbl(Val(reader("qty").ToString()))
                    id = reader("id").ToString()
                End If
                reader.Close()
                Dim total As Double = qty + (CDbl(Val(TextBox3.Text)))
                Dim com_flg = 0
                If total <> 0 And total < 0 Then
                    com_flg = 1
                ElseIf total > 0 Then
                    com_flg = 0
                End If
                Dim update_stock_wp As String = "update sup_frith_in_out set qty = '" & total & "' , com_flg = '" & com_flg & "' where id = '" & id & "'"
                Dim cmd_stock_wp As SqlCommand = New SqlCommand(update_stock_wp, myConn)
                reader = cmd_stock_wp.ExecuteReader()
                reader.Close()
                MsgBox("UPDATE OK")
                Dim get_qty = get_qty_database(TextBox1.Text, TextBox6.Text, Module1.M_SEQ_PRINT)
                'MsgBox("get_qty = " & get_qty)
                qty_databse = 0.0
                If get_qty <> "NO_DATA" Then
                    qty_databse = CDbl(Val(get_qty)) + CDbl(Val(TextBox3.Text))
                    qty_databse_STring = qty_databse
                    status = 1
                Else
                    MsgBox("NO DATA SEQ AND ITEM  AND PO")
                    status = 0
                End If
                status = 1
            End If
            Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet.addr = main.number_printter_bt
            Dim pin As StringBuilder = New StringBuilder("0000")

            Dim pinlen As UInt32 = CType(pin.Length, UInt32)
loop_check_printer:
            If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
                Panel4.Visible = False
                alert_open_printer.Visible = False
                Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
                stInfoSet1.addr = main.number_printter_bt
                Dim pin1 As StringBuilder = New StringBuilder("0000")
                Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
                Dim time As DateTime = DateTime.Now
                Dim format As String = "yyyy-MM-dd HH:mm:ss"
                Dim date_now = time.ToString(format)
                Dim time_detail As DateTime = DateTime.Now
                Dim format_time_detail As String = "HH:mm:ss"
                Dim now_time_detail = time_detail.ToString(format_time_detail)
                Dim date_detail As DateTime = DateTime.Now
                Dim format_date_detail As String = "dd-MM-yyyy"
                Dim now_date_detail = date_detail.ToString(format_date_detail)
                Dim RESULT_QTY As String = ""
                If PO_LOT.Length = "4" Then
                    'MsgBox("qty lot = " & qty_databse_STring)
                    Dim number_re_qty As Integer = Len(qty_databse_STring)
                    ' MsgBox("number_re_qty =  " & number_re_qty)
                    Dim regit As Integer = 7 - number_re_qty
                    ' MsgBox("regit = " & regit)
                    Dim charArray_re_qty() As Char = qty_databse_STring.ToCharArray
                    Dim key As Integer = 0
                    Dim n As Integer = 0
                    For i As Integer = 1 To 6 Step +1
                        Dim data_rigit As String = ""
                        ' MsgBox(i & "=" & regit)
                        If i = regit Then
                            'MsgBox("data = ===" & charArray_re_qty(key))
                            data_rigit = charArray_re_qty(key)
                            regit = regit + 1
                            key = key + 1
                        Else
                            data_rigit = " "
                            n = n + 1
                        End If
                        RESULT_QTY &= data_rigit
                    Next
                    'MsgBox("RESULT_QTY = " & RESULT_QTY)
                    qr_detail_remain = tag_read.Substring(0, 52) & RESULT_QTY & tag_read.Substring(58)
                ElseIf PO_LOT.Length = "10" Then
                    Dim key As Integer = 0
                    RESULT_QTY = ""
                    Dim charArray_re_qty() As Char = qty_databse_STring.ToCharArray
                    Dim regit As Integer = 9 - Len(qty_databse_STring)
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
                    qr_detail_remain = tag_read.Substring(0, 50) & RESULT_QTY & tag_read.Substring(59)
                End If

                If status = 1 Then
                    Api_update_qty(TextBox1.Text, qr_detail_remain.Substring(2, 10), Module1.M_SEQ_PRINT, qty_databse)
                    F_save_logs_return(TextBox1.Text)
                    Bluetooth_Print_MB300i(stInfoSet, pin, pinlen1, TextBox1.Text, TextBox4.Text, model, qty_databse, TextBox5.Text, Module1.user_id, now_date_detail, now_time_detail, qr_detail_remain, M_CHECK_TYPE)
                End If
            Else
                Panel4.Visible = True
                alert_open_printer.Visible = True
                GoTo loop_check_printer
            End If
        Else
            Panel4.Visible = True
            alert_not_enogth_return.Visible = True
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
            cerrent.Focus()
        End If
    End Sub
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
                ' MsgBox("กรุณา เปิดเครื่องปริ้นด้วย")
                'disp = "btBluetoothPairing error ret[" & ret & "]"
                'MessageBox.Show(disp, "Error")
                Return bRet
            End If

            ret = Bluetooth.btBluetoothSPPConnect(stInfoSet, 30000)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPConnect error ret[" & ret & "]"
                MessageBox.Show(disp, "Error")
                Return bRet
            End If
            alert_open_printer.Visible = False 'ปิด alert Printet'
            alert_open_printer.Enabled = False
            bRet = True
            Return bRet
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            Return bRet
        Finally
        End Try
    End Function
    Private Sub Bluetooth_Print_MB300i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32, ByVal part_number As String, ByVal part_name As String, ByVal part_model As String, ByVal remain_qty As String, ByVal loc_detail As String, ByVal user_detail As String, ByVal date_detail As String, ByVal time_detail As String, ByVal qr_detail_remain As String, ByVal M_CHECK_TYPE As String)
        Dim ret As Int32 = 0
        Dim disp As [String] = ""

        Dim sbBuf As New StringBuilder("")
        Dim ssizeGet As UInt32 = 0
        Dim bBuf As [Byte]() = New Byte() {}


        Dim rsizeGet As UInt32 = 0
        Dim bBufGet As [Byte]() = New [Byte](4094) {}

        Try


            'MsgBox("wtf")

            ' Data transmission
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
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("PG33A1010112800384+000+000+00+00+00005000")
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("PG33A1010112800384+000+000+00+00+00005100") '// Journal
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            '// Title
            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("v0020")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("h0010")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("p02")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("l0202")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            'bESC.CopyTo(bBuf, len)

            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("k9b")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("label print")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            '// Human readable
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Remain Tag")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Remain QTY. : " & remain_qty)
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

            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Location : " & loc_detail)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length





            '  bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V180")
            ' bBufWork.CopyTo(bBuf, len)
            ' len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H210")
            ' bBufWork.CopyTo(bBuf, len)
            ' len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            ' bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

            '  bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            '  bBufWork.CopyTo(bBuf, len)
            ' len = len + bBufWork.Length
            ' bESC.CopyTo(bBuf, len)

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & time_detail)
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length



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

            If M_CHECK_TYPE = "WEB_POST" Then
                Dim PO = qr_detail_remain.Substring(2, 10)
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "PO No: " & PO)
            ElseIf M_CHECK_TYPE = "FW" Then
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "SEQ : " & Module1.M_SEQ_PRINT)
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Revised By : " & Module1.A_USER_ID)

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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "DATE : " & date_detail & " TIME : " & time_detail)

            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length


            '  bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            '  bBufWork.CopyTo(bBuf, len)
            ' len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            '  len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V300")
            ' bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            '  bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H405")
            '  bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            ' bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            '  bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length
            '   bESC.CopyTo(bBuf, len)

            '  bESC.CopyTo(bBuf, len)
            '  len = len + bESC.Length
            '   bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "< *** Remain Tag *** >")
            '   bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

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
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H200")
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
                '''''''''''''''''''''''''''''''
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V20")
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V710")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H360")
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
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & " OO ")
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

    Private Sub alert_open_printer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel3.GotFocus

    End Sub

    Private Sub Panel4_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Public Function get_qty_database(ByVal item_cd As String, ByVal PO As String, ByVal SEQ As String)
        Dim result As String = "NO_DATA"
        Try
            'MsgBox("http://192.168.161.102/exp_api3party/Api_cut_stock_web_post/get_data_service?ITEM_CD=" & F_item_cd & "&PO=" & PO & "&SEQ=" & updated_seq & " &QTY=" & used_qty)
            result = Api.update_data("http://192.168.161.102/exp_api3party/Api_cut_stock_web_post/get_data_qty_from_seq?ITEM_CD=" & item_cd & "&PO=" & PO & "&SEQ=" & SEQ)
            'MsgBox("status = " & result)
            If result <> "NO_DATA" Then
                ' MsgBox("FALL UPDATE CUT_STOCK P_NICE = " & result)
            End If
        Catch ex As Exception
            MsgBox("FAILL cut_stock_FASYSTEM" & vbNewLine & ex.Message, "FAILL")
        End Try
        Return result
    End Function
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

    Private Sub Panel2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel2.GotFocus

    End Sub

    Private Sub Label12_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.ParentChanged

    End Sub

    Private Sub cerrent_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cerrent.KeyDown
         Select e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                Panel4.Visible = False
                alert_no_tag_supply.Visible = False
                alert_not_enogth_return.Visible = False
                TextBox2.Text = ""
                TextBox2.Focus()
        End Select
    End Sub
    Public Sub F_save_logs_return(ByVal item_cd As String)
        Try
            Dim supply_tag As String = TextBox2.Text
            Dim arr_data = supply_tag.Split(" ")
            Dim ID_pick_log As String = arr_data(21)
            Dim WI = arr_data(2)
            Dim QTY_RETURN As String = TextBox3.Text
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)

            Dim time_detail As DateTime = DateTime.Now
            Dim format_time_detail As String = "HH:mm:ss"
            Dim now_time_detail = time_detail.ToString(format_time_detail)

            Dim date_detail As DateTime = DateTime.Now
            Dim format_date_detail As String = "dd-MM-yyyy"
            Dim now_date_detail = date_detail.ToString(format_date_detail)
            Dim REF_ID As String = "NO_DTA"
            Dim get_data As String = "select REF_ID from sup_pick_log where ID='" & ID_pick_log & "'"
            MsgBox(get_data)
            Dim command2 As SqlCommand = New SqlCommand(get_data, myConn)
            reader = command2.ExecuteReader()
            MsgBox("0")
            If reader.Read() Then
                REF_ID = reader("REF_ID").ToString()
            End If
            reader.Close()
            MsgBox("1")
            If REF_ID <> "NO_DTA" Then
                MsgBox("2")
                Dim str As String = "insert into sys_logs_return ( REF_ID , WI_NO , RETURN_QTY , CREATED_DATE ,CREATED_BY , UPDATED_DATE , UPDATED_BY , enable , STATUS) values('" & REF_ID & "' , '" & WI & "' , '" & QTY_RETURN & "' , '" & now_date_detail & "' , '" & Module1.A_USER_ID & "','" & now_date_detail & "' , '" & Module1.A_USER_ID & "' , '1' , '2')"
                Dim cmd_insert As SqlCommand = New SqlCommand(str, myConn)
                MsgBox(str)
                MsgBox("3")
                reader = cmd_insert.ExecuteReader()
                MsgBox("4")
                reader.Close()
            Else
                MsgBox("REF_ID = NO_DATA")
            End If
        Catch ex As Exception
            MsgBox("ERROR F_save_logs_return" & vbNewLine & ex.Message, "FAILL")
        End Try
      
    End Sub
End Class