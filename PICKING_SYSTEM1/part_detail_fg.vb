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

Public Class part_detail_fg
    Inherits Form
    Public myConn As SqlConnection
    Public myConn_fa As SqlConnection
    Public myConn_Resive As SqlConnection
    Public REMAIN_ID As String = "NO_DATA"
    Public ID_table_detail As String = "NO_DATA"
    Public check_process As String = "NO_OK"
    Public c_num As Integer = 0
    Dim g_index As Integer = 0
    Dim id_cut_stock_FW As String = "no_data"
    Dim path As String
    Public fa_qty_total_check As Integer = 0
    Dim a As Integer = 0
    Dim count_arr_fw As Integer = 0
    Dim count_fw_final As Integer = 0
    Dim frith As Integer = 0
    Public leng_scan_qty As Integer = 0
    Dim imagefile As String
    Public PD_ADD_PART As select_pick_add
    Dim reader As SqlDataReader
    Dim dat As String = String.Empty
    Dim path1 As String
    Dim htlogfile As String
    Dim pclogfile As String = "logfile.csv"
    Dim data_final As String = "NOOOOO"
    Dim data_final_loop As String = "NOOOOO"
    Public totall_qty_scan As Double = 0.0
    Public Line As Select_Line
    Dim CodeType As String = "QR"
    Public c_check As String = "no_process"
    Dim m As String = "no-data"
    Dim status_alert_image As String = "NO_STATUS"
    Dim g_update As Integer = 0 '
    Dim brak_loop As Integer = 0 '
    Public check_po_lot As String = "NODATA"
    Public id_pick_log_supply As String = "NODATA"
    Dim j As Integer = 0
    Dim count_update_fw As Integer = 0
    Dim count_scan As Integer = 0
    Dim fa_use_total As Integer = 0
    Public length As Integer = 0
    Public Len_length_QR As Integer = 0
    Public check_scan As Integer = 0
    Public check_count__data As Integer = 0
    Public QTY_INSERT_LOT_PO As Double = 0.0
    Dim arr_remain_qty As ArrayList = New ArrayList()
    Dim arr_up_id As ArrayList = New ArrayList()
    Dim F_wi As ArrayList = New ArrayList()
    Dim F_item_cd As ArrayList = New ArrayList()
    Dim F_scan_qty As ArrayList = New ArrayList()
    Dim F_scan_lot As ArrayList = New ArrayList()
    Dim F_tag_typ As ArrayList = New ArrayList()
    Dim F_tag_readed As ArrayList = New ArrayList()
    Dim F_Line_cd As ArrayList = New ArrayList()
    Dim F_scan_emp As ArrayList = New ArrayList()
    Dim F_term_cd As ArrayList = New ArrayList()
    Dim F_updated_date As ArrayList = New ArrayList()
    Dim F_updated_by As ArrayList = New ArrayList()
    Dim F_updated_seq As ArrayList = New ArrayList()
    Dim F_com_flg As ArrayList = New ArrayList()
    Dim F_tag_remain_qty As ArrayList = New ArrayList()
    Dim F_Create_Date As ArrayList = New ArrayList()
    Dim F_Create_By As ArrayList = New ArrayList()
    Dim F_delivery_date As ArrayList = New ArrayList()
    Dim F_control_box As ArrayList = New ArrayList()
    Dim check_data As ArrayList = New ArrayList()
    Dim arr_pick_log As ArrayList = New ArrayList()
    '--------------------------------------------------------------
    ' Constant definitions
    '--------------------------------------------------------------
    ' Constant to use with wininet
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

    Private Sub part_detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            Dim en As Int32 = path.LastIndexOf("\")
            path = path.Substring(0, en)
            path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            Dim connect_db = New connect()
            myConn = connect_db.conn()
            myconn_fa = connect_db.conn_fa()
        Finally

            ' Dim path = Me.GetType().Assembly.GetModules()(0).FullyQualifiedName
            'MsgBox(path)
            Panel4.Visible = False
            scan_qty.Visible = False
            lb_code_user.Text = main.show_code_id_user()
            lb_code_pd.Text = "FINISH GOOD" 'PD5.lb_code_pd.Text
            Part_No.Text = Module1.FG_PART_CD
            Part_name.Text = "Part Name:" & Module1.FG_PART_NAME.Substring(12) ' PD5.Part_Name.Text
            show_qty.Text = "QTY : " & Module1.FG_QTY.Substring(6)
            location.Text = "Location : " & Module1.FG_LOCATIONS.Substring(10)
            show_number_supply.Text = 0
            show_number_remain.Text = 0
            want_to_tag.Text = show_qty.Text.Substring(6)
            lot_no.Hide()
            Button2.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            PictureBox3.Visible = False
            alert_detail.Visible = False
            alert_pa.Visible = False
            alert_tag_remain.Visible = False
            alert_pickdetail_number.Visible = False
            alert_pickdetail_ok.Visible = False
            'show_pln.Visible = False
            ' show_pln.Visible = False
            text_box_success.Visible = False
            alert_success.Visible = False
            alert_success_remain.Visible = False
            alert_loop.Visible = False
            alert_right_fa.Visible = False
            Panel5.Visible = False
            alert_reprint.Visible = False
            alert_open_printer.Visible = False
            alert_no_tranfer_data.Visible = False
            Panel7.Visible = False
            'ชั่วคราว'
            ' Panel6.Visible = False
            btn_detail_part.Visible = True
            'get_data_tetail() 'ปิดโชว์ FIFO
        End Try
    End Sub
    Public Sub New()
        InitializeComponent()
        If Api.DownloadImage("http://192.168.161.102/picking_system/uploads/pic/" & Module1.FG_PART_CD.Substring(16) & ".jpg") IsNot Nothing Then
            show_img.Image = Api.DownloadImage("http://192.168.161.102/picking_system/uploads/pic/" & Module1.FG_PART_CD.Substring(16) & ".jpg")

        End If
        'show_img_part.Image = 
    End Sub
    Private Sub lb_code_user_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub lb_code_pd_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub lb_code_user_ParentChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb_code_user.ParentChanged

    End Sub

    Private Sub Part_No_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Part_No.ParentChanged

    End Sub

    Private Sub Label1_ParentChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged

    End Sub

    Private Sub Part_name_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Part_name.ParentChanged

    End Sub

    Private Sub PictureBox2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub location_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles location.ParentChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        set_default_data()
        QTY_INSERT_LOT_PO = 0.0
        Module1.G_show_data_supply = 0.0
        'MsgBox("BACK")
        Module1.arr_check_lot_scan = New ArrayList()
        Module1.arr_check_PO_scan = New ArrayList()
        Module1.arr_check_QTY_scan = New ArrayList()
        ' MsgBox("DATA NULL")
        Module1.show_data_supply = 0.0
        Module1.show_data_remain = 0.0
        Module1.check_pick_detail = 0
        Module1.total_qty = 0
        Module1.total_database = 0
        delete_data_check_qr_part() 'ยังไม่ได้สร้าง function'
        Select_plan_fg.Show()
        Me.Close()
    End Sub

    Private Sub scan_part_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles text_tmp.TextChanged

    End Sub
    Public Sub get_data_image()
        Dim temp_part As String = Part_No.Text.Substring(16)
        'MsgBox("temp_part = " & temp_part)
        Try
            'MsgBox("1")
            imagefile = path & "\img\" & temp_part & ".jpg"
            If Not show_img.Image Is Nothing Then
                show_img.Image.Dispose()
                show_img.Image = Nothing
            End If

            show_img.Image = New Bitmap(imagefile)
            'show_img.Focus()
        Catch ex As Exception
            'MsgBox("2")
            imagefile = "\img\t.jpg"
            If Not show_img.Image Is Nothing Then
                show_img.Image.Dispose()
                show_img.Image = Nothing
            End If

            'show_img.Image = New Bitmap(imagefile)
            'show_img.Focus()
        End Try

        'text_tmp.Focus()
    End Sub
    Public Sub check_qr_part()
        Dim ps = Part_No.Text.Substring(16)
        Dim scan_p = text_tmp.Text.Substring(19)
        'If ps = scan_p Then

        '  End If

    End Sub
    Private Sub show_img_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles show_img.Click

    End Sub

    Private Sub lb_code_line_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub

    Dim scan_qty_total As Integer = 0
    Dim comp_flg As String = "0"
    Dim firstscan As String = "0"

    Dim sup_seq_scan_no As Integer = 0
    Dim sup_list As New ArrayList

    Dim supp_total_qty As Integer = 0
    Dim supp_tag_qty As Integer = 0
    Dim supp_seq As String = 0
    Dim supplier_cd As String = 0
    Dim order_number As String = ""

    Dim fa_total_qty As Integer = 0
    Dim fa_tag_qty As Integer = 0
    Dim fa_lot_seq As String = 0
    Dim fa_tag_seq As String = 0
    Dim fa_line_cd As String = 0
    Dim fa_lot_no As String = ""
    Dim fa_act_date As String = ""
    Dim fa_qty As Integer = 0
    Dim fa_plant_cd As String = ""
    Dim fa_seq As String = 0

    Dim fa_shift_seq As String = ""

    Dim item_cd_scan As String = ""
    Public remain_qty1 As Double = 0

    Dim scan_qty_arrlist As New ArrayList
    Dim scan_lot_arrlist As New ArrayList
    Dim scan_read_arrlist As New ArrayList
    Dim scan_seq_arrlist As New ArrayList

    Private Sub scan_qty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles scan_qty.KeyDown
comeback:
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.F3
                'Module1.check_page = "part_detail_fg"
                'Dim setting As setting = New setting()
                'setting.Show()
                'Me.Hide()
            Case System.Windows.Forms.Keys.F4
                If Panel6.Visible = True Then
                    Panel6.Visible = False
                Else
                    Panel6.Visible = True
                End If
            Case System.Windows.Forms.Keys.Enter
                Dim testString As String = scan_qty.Text
                Module1.scan_qr_part_detail = scan_qty.Text
                ' Returns 11.
                Dim testLen As Integer = Len(testString)

                length = testLen
                leng_scan_qty = length
                'MsgBox(testLen)วว
                Dim req_qty As Double = 0.0
                req_qty = Val(Module1.FG_QTY.Substring(6))
                Dim number_remain As Double = 0.0
                number_remain = CDbl(Val(show_number_remain.Text))
                Dim ps = Part_No.Text.Substring(16)
                Dim QTY_show As Integer = show_qty.Text.Substring(6)
                'MsgBox("number_remain = " & number_remain)
                If comp_flg = "0" Then
                    ' MsgBox(testLen)
                    If testLen = 62 Then
                        supplier_cd = scan_qty.Text.Substring(37, 6)
                        If check_count__data <> 1 Then '1 คือ ห้ามใส่ค่าลงไป supp_total_qty'
                            supp_total_qty = scan_qty.Text.Substring(43, 8)
                        End If
                        supp_tag_qty = scan_qty.Text.Substring(51, 8)
                        supp_seq = scan_qty.Text.Substring(59, 3)
                        fa_qty = scan_qty.Text.Substring(52, 6)
                        Dim SearchForThis As String = " "
                        Dim FirstCharacter As Integer = testString.IndexOf(SearchForThis)

                        item_cd_scan = scan_qty.Text.Substring(12, FirstCharacter - 12)
                        order_number = scan_qty.Text.Substring(2, 10)
                        If ps = item_cd_scan Then
                            'เคสเหลือจาก Tag 
                            text_tmp.Text = supp_tag_qty
                            If Ck_dup(ListBox, order_number & supp_seq) = True Then
                                If bool_check_scan = "ever" Then
                                    scan_qty.Text = ""
                                    text_tmp.Text = scan_qty_total
                                    Re_scan()
                                ElseIf bool_check_scan = "HAVE_TAG_REMAIN" Then
                                    ' MsgBox("มี Tag Remain เหลืออยู่ กรุณา สแกน Tag Remain ก่อน", 16, "Alert")
                                    'alert_detail.Visible = False
                                    Panel7.Visible = True
                                    alert_tag_remain.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()

                                ElseIf bool_check_scan = "Plase_scna_detail" Then
                                    ' MsgBox("กรุณา scan ตาม Detail", 16, "Alert")
go_pelase_detail_webpost:
                                    Panel7.Visible = True
                                    alert_detail.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "HAVE_Reprint" Then
                                    'MsgBox("TAG นี้ไม่สามารถ สแกนได้แล้ว เนื่องจาก reprint ไปแล้ว")
                                    Panel7.Visible = True
                                    alert_reprint.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "scan_ok_pickdetail" Then
go_scan_ok_pickdetail:
                                    Panel7.Visible = True
                                    alert_pickdetail_ok.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "pick_detail_number" Then
go_pick_detail_number:
                                    Panel7.Visible = True
                                    alert_pickdetail_number.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                End If
                            Else
                                If show_number_supply.Text > req_qty And firstscan = "0" And number_remain > 0 Then
                                    'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                    check_scan = 2
                                    check_text_box_qr_code()
                                    text_tmp.Text = supp_tag_qty
                                    remain_qty1 = supp_tag_qty - req_qty
                                    Module1.tag_remain_qty = supp_tag_qty - req_qty
                                    'Button3.Visible = True
                                    'Button2.Visible = True
                                    Dim summa As Integer = supp_tag_qty - remain_qty1
                                    scan_qty_arrlist.Add(summa)
                                    scan_lot_arrlist.Add(order_number)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(order_number & supp_seq)
                                    scan_qty.Visible = False
                                    comp_flg = "1"
                                    Panel7.Visible = True
                                    alert_success_remain.Visible = True
                                    status_alert_image = "success_remain"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                    'เคสเท่ากับ Tag
                                ElseIf req_qty = show_number_supply.Text And number_remain = 0 Then
                                    'MsgBox("คุณสแกนครบแล้ว", 16, "Alert")
                                    check_text_box_qr_code()
                                    Button2.Visible = True
                                    scan_qty_arrlist.Add(supp_tag_qty)
                                    scan_lot_arrlist.Add(order_number)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(order_number & supp_seq)
                                    comp_flg = "1"
                                    scan_qty.Visible = False
                                    check_scan = 2
                                    Panel7.Visible = True
                                    alert_success.Visible = True
                                    status_alert_image = "success"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                    'เคสยิงสะสม
                                ElseIf show_number_supply.Text = req_qty And number_remain > 0 Then
                                    'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                    check_scan = 2
                                    check_text_box_qr_code()
                                    text_tmp.Text = supp_tag_qty
                                    remain_qty1 = supp_tag_qty - req_qty
                                    Module1.tag_remain_qty = supp_tag_qty - req_qty
                                    'Button3.Visible = True
                                    Button2.Visible = True
                                    Dim summa As Integer = supp_tag_qty - remain_qty1
                                    scan_qty_arrlist.Add(summa)
                                    scan_lot_arrlist.Add(order_number)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(order_number & supp_seq)
                                    scan_qty.Visible = False
                                    comp_flg = "1"
                                    Panel7.Visible = True
                                    alert_success_remain.Visible = True
                                    status_alert_image = "success_remain"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                Else
                                    sup_seq_scan_no = sup_seq_scan_no + 1
                                    sup_list.Add(supp_seq)
                                    Dim num As Integer = sup_seq_scan_no

                                    'MsgBox(sup_seq_scan_no)

                                    If Module1.check_count = 1 Or Module1.check_count2 = 1 Then
                                        If Module1.bool_check_scan = "ever" Then
                                            'MsgBox("Scan ซ้ำ! มีการสแกนแล้วเมื่อสักครู่", 16, "Alert")
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
                                            '  Bt.SysLib.Device.btBuzzer(1, stBuz)
                                            '  Bt.SysLib.Device.btVibrator(1, stVib)
                                            ' Bt.SysLib.Device.btLED(1, stLed)
                                            text_tmp.Text = scan_qty_total
                                            Panel7.Visible = True
                                            alert_loop.Visible = True
                                            status_alert_image = "loop"
                                            text_box_success.Focus()
                                            GoTo exit_keydown
                                        End If
                                    Else
                                        'MsgBox("ready check")
                                        check_po_lot = "pick_ok"
                                        Dim qty_scan_wp = scan_qty.Text.Substring(51, 8)
                                        totall_qty_scan += CDbl(Val(qty_scan_wp))
                                        Dim Code_suppier As String = scan_qty.Text.Substring(37, 5)
                                        Dim order_number_check As String = scan_qty.Text.Substring(2, 10)
                                        Dim check_po As Integer = check_scan_detail_PO(order_number_check, Code_suppier)
                                        If check_po = 0 Then 'check ว่า scan ถูกใน  pickdetail มั้ย'
                                            bool_check_scan = "Plase_scna_detail"
                                            GoTo go_pelase_detail_webpost
                                        ElseIf check_po = 2 Then
                                            bool_check_scan = "pick_detail_number"
                                            GoTo go_pick_detail_number
                                        ElseIf check_po = 3 Then
                                            bool_check_scan = "scan_ok_pickdetail"
                                            GoTo go_scan_ok_pickdetail
                                        ElseIf check_po = 5 Then
                                            bool_check_scan = "ever"
                                            GoTo alert_ever
                                        ElseIf check_po = 1 Then
                                            inset_check_qr_part()
                                        End If
                                        number_remain = CDbl(Val(show_number_remain.Text)) 'เอาค่า remain แปลง'
                                        'MsgBox("number_remain = " & number_remain)
                                        ListBox.Items.Add(order_number & supp_seq)
                                        scan_qty_total = supp_tag_qty + scan_qty_total
                                        Module1.SCAN_QTY_TOTAL = scan_qty_total
                                        text_tmp.Text = scan_qty_total
                                        '  MsgBox("ยอดที่คุณสแกน : " & supp_tag_qty, 16, "Alert")
                                        Module1.QTY_OF_SCAN = supp_tag_qty
                                        firstscan = "1"
                                        check_scan = 1
                                        'Button2.Visible = True 'สำหรับเอา stock แค่ครึ่งเดียว'asd

                                        If show_number_supply.Text > req_qty And number_remain > 0 Then
                                            'MsgBox("")
                                            'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                            check_scan = 2
                                            'Button3.Visible = True
                                            Button2.Visible = True
                                            remain_qty1 = scan_qty_total - req_qty
                                            Dim summa As Integer = supp_tag_qty - remain_qty1
                                            scan_qty_arrlist.Add(summa)
                                            scan_lot_arrlist.Add(order_number)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(order_number & supp_seq)
                                            comp_flg = "1"
                                            scan_qty.Visible = False
                                            check_text_box_qr_code()
                                            Panel7.Visible = True
                                            alert_success_remain.Visible = True
                                            status_alert_image = "success_remain"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                        ElseIf req_qty = show_number_supply.Text And number_remain = 0 Then
                                            'MsgBox("คุณสแกนครบแล้ว", 16, "Alert")
                                            Button2.Visible = True
                                            scan_qty_arrlist.Add(supp_tag_qty)
                                            scan_lot_arrlist.Add(order_number)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(order_number & supp_seq)
                                            comp_flg = "1"
                                            check_scan = 2
                                            scan_qty.Visible = False
                                            Panel7.Visible = True
                                            alert_success.Visible = True
                                            status_alert_image = "success"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                        ElseIf show_number_supply.Text = req_qty And number_remain > 0 Then
                                            'MsgBox("")
                                            'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                            check_scan = 2
                                            'Button3.Visible = True
                                            Button2.Visible = True
                                            remain_qty1 = scan_qty_total - req_qty
                                            Dim summa As Integer = supp_tag_qty - remain_qty1
                                            scan_qty_arrlist.Add(summa)
                                            scan_lot_arrlist.Add(order_number)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(order_number & supp_seq)
                                            comp_flg = "1"
                                            scan_qty.Visible = False
                                            check_text_box_qr_code()
                                            Panel7.Visible = True
                                            alert_success_remain.Visible = True
                                            status_alert_image = "success_remain"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                        Else
                                            scan_qty_arrlist.Add(supp_tag_qty)
                                            scan_lot_arrlist.Add(order_number)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(order_number & supp_seq)
                                        End If

                                    End If
                                    scan_qty.Text = ""
                                    scan_qty.Focus()

                                End If
                            End If
                        Else
                            'MsgBox("Part incorrect")
                            status_alert_image = "Part_incorrect"
                            Panel7.Visible = True
                            alert_pa.Visible = True
                            text_box_success.Focus()
                        End If

                    ElseIf testLen = 76 Then
                        status_alert_image = "alert_right_fa"
                        Panel7.Visible = True
                        alert_right_fa.Visible = True
                        text_box_success.Focus()
                        'MsgBox("Please scan FA tag on the top right")
                    ElseIf testLen = 103 Then
                        Try
                            Dim SearchForThis As String = " "
                            Dim FirstCharacter As Integer = testString.IndexOf(SearchForThis)
                            item_cd_scan = scan_qty.Text.Substring(19, FirstCharacter - 19)

                            fa_line_cd = scan_qty.Text.Substring(2, 6)
                            fa_act_date = scan_qty.Text.Substring(8, 8)
                            fa_lot_seq = scan_qty.Text.Substring(16, 3)
                            fa_qty = scan_qty.Text.Substring(52, 6)
                            ' If check_count__data <> 1 Then '1 คือ ห้ามใส่ค่าลงไป supp_total_qty'
                            'scan_qty_total = fa_qty
                            ' End If
                            fa_lot_no = scan_qty.Text.Substring(58, 4)
                            fa_shift_seq = scan_qty.Text.Substring(95, 3)
                            fa_plant_cd = scan_qty.Text.Substring(98, 2)
                            fa_tag_seq = scan_qty.Text.Substring(100, 3)


                        Catch ex As Exception
                            Dim MyValue As Integer
                            MyValue = Int((9 * Rnd()) + 1) & Int((9 * Rnd()) + 1) & Int((9 * Rnd()) + 1)
                            fa_line_cd = scan_qty.Text.Substring(2, 6)
                            fa_act_date = Trim(scan_qty.Text.Substring(44, 8))
                            fa_lot_seq = MyValue
                            fa_qty = scan_qty.Text.Substring(52, 6)
                            fa_lot_no = scan_qty.Text.Substring(58, 4)
                            fa_plant_cd = "51"
                            fa_tag_seq = MyValue
                            item_cd_scan = Trim(scan_qty.Text.Substring(18, 25))
                        End Try
                        If ps <> item_cd_scan Then
                            Try
                                Dim str_item = scan_qty.Text.Split(" ")
                                Dim arr_data = str_item(18)
                                If arr_data = "" Then
                                    arr_data = str_item(29)
                                End If
                                Dim res_data = arr_data.Substring(6)
                                '' MsgBox(ps)
                                '  MsgBox(res_data)
                                If ps = res_data Or ps.Substring(0, 11) = res_data Then
                                    item_cd_scan = res_data
                                    GoTo next_station
                                End If
                            Catch ex As Exception
                                GoTo check_E
                            End Try
check_E:
                            Try
                                Dim str_item = scan_qty.Text.Split(" ")
                                Dim arr_data = str_item(18)
                                If arr_data = "" Then
                                    arr_data = str_item(29)
                                End If
                                Dim res_data = arr_data.Substring(6, 12)
                                'MsgBox("CHECK_E===" & ps)
                                ' MsgBox("CHECK_E" & res_data)
                                If ps = res_data Then
                                    item_cd_scan = res_data
                                    GoTo next_station
                                End If
                            Catch ex As Exception
                                GoTo check_loop
                            End Try
                        End If
check_loop:
                        If ps = item_cd_scan Then 'note'
next_station:
                            ' Button2.Visible = True 'สำหรับเอา stock แค่ครึ่งเดียว'
                            text_tmp.Text = fa_qty
                            'MsgBox(req_qty.Text)
                            'MsgBox(text_tmp.Text)

                            'เคสเหลือจาก Tag
                            If Ck_dup(ListBox, order_number & supp_seq) = True Then
alert_ever:
                                If bool_check_scan = "ever" Then
                                    text_tmp.Text = ""
                                    Re_scan_fa()
                                    GoTo exit_keydown
                                ElseIf bool_check_scan = "HAVE_TAG_REMAIN" Then
                                    '  MsgBox("มี Tag Remain เหลืออยู่ กรุณา สแกน Tag Remain ก่อน", 16, "Alert")
                                    Panel7.Visible = True
                                    alert_tag_remain.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "Plase_scna_detail" Then
                                    ' MsgBox("กรุณา scan ตาม Detail", 16, "Alert")
go_pelase_detail:
                                    Panel7.Visible = True
                                    alert_detail.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "HAVE_Reprint" Then
                                    Panel7.Visible = True
                                    alert_reprint.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "scan_ok_pickdetail" Then
go_scan_ok_pickdetail_fw:
                                    Panel7.Visible = True
                                    alert_pickdetail_ok.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "pick_detail_number" Then
go_pick_detail_number_fw:
                                    Panel7.Visible = True
                                    alert_pickdetail_number.Visible = True
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                ElseIf bool_check_scan = "NO_data_tranfer" Then

go_No_data_tranfer:
                                    Panel7.Visible = True
                                    alert_no_tranfer_data.Visible = True
                                    check_qr.Visible = True
                                    check_qr.Focus()
                                    GoTo exit_keydown
                                End If
                            Else
                                If show_number_supply.Text > req_qty And firstscan = "0" And number_remain > 0 Then
                                    ' MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                    text_tmp.Text = fa_qty
                                    remain_qty1 = fa_tag_qty - req_qty
                                    'Button2.Visible = True
                                    'MsgBox(remain_qty1)
                                    'Button4.Visible = True
                                    Button2.Visible = True
                                    Dim summa As Integer = fa_tag_qty - remain_qty1
                                    check_scan = 2
                                    scan_qty_arrlist.Add(summa)
                                    scan_lot_arrlist.Add(fa_lot_no)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                    comp_flg = "1"
                                    firstscan = "1"
                                    check_text_box_qr_code()
                                    scan_qty.Visible = False
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                    Panel7.Visible = True
                                    alert_success_remain.Visible = True
                                    status_alert_image = "success_remain"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                    'เคสเท่ากับ Tag
                                ElseIf req_qty = show_number_supply.Text And number_remain = 0 Then
                                    '  MsgBox("คุณสแกนครบแล้ว", 16, "Alert")
                                    Button2.Visible = True
                                    check_scan = 2
                                    scan_qty_arrlist.Add(fa_qty)
                                    scan_lot_arrlist.Add(fa_lot_no)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                    comp_flg = "1"
                                    firstscan = "1"
                                    check_text_box_qr_code()
                                    scan_qty.Visible = False
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                    Panel7.Visible = True
                                    alert_success.Visible = True
                                    status_alert_image = "success"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                    'เคสยิงสะสม
                                ElseIf show_number_supply.Text = req_qty And firstscan = "0" And number_remain > 0 Then
                                    ' MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                    text_tmp.Text = fa_qty
                                    remain_qty1 = fa_tag_qty - req_qty
                                    'Button2.Visible = True
                                    'MsgBox(remain_qty1)
                                    'Button4.Visible = True
                                    Button2.Visible = True
                                    Dim summa As Integer = fa_tag_qty - remain_qty1
                                    check_scan = 2
                                    scan_qty_arrlist.Add(summa)
                                    scan_lot_arrlist.Add(fa_lot_no)
                                    scan_read_arrlist.Add(scan_qty.Text)
                                    scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                    comp_flg = "1"
                                    firstscan = "1"
                                    check_text_box_qr_code()
                                    scan_qty.Visible = False
                                    text_box_success.Visible = True
                                    text_box_success.Focus()
                                    Panel7.Visible = True
                                    alert_success_remain.Visible = True
                                    status_alert_image = "success_remain"
                                    text_box_success.Focus()
                                    GoTo exit_scan
                                Else
                                    'MsgBox("test")
                                    'fa_tag_seq = fa_tag_seq + 1
                                    fa_seq = fa_seq + 1
                                    sup_list.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                    Dim num As Integer = fa_seq
                                    'MsgBox(fa_tag_seq)
                                    'MsgBox(fa_seq)

                                    If Module1.check_count = 1 Or Module1.check_count2 = 1 Then 'มี part แล้ว'
                                        Re_scan_fa()
                                        GoTo exit_keydown
                                    Else
                                        check_po_lot = "pick_ok"
                                        Dim QTY_FW = scan_qty.Text.Substring(52, 6)
                                        totall_qty_scan += CDbl(Val(QTY_FW))

                                        If check_FA_TAG_FG() = False Then
                                            bool_check_scan = "No_data_tranfer"
                                            GoTo go_No_data_tranfer
                                        ElseIf check_FA_TAG_FG() = 1 Then
                                            bool_check_scan = "ever"
                                            GoTo alert_ever
                                        End If

                                        Dim check_po As Integer = check_scan_detail_PO("NO_DATA", "NO_DATA")
                                        If check_po = 0 Then 'check ว่า scan ถูกใน  pickdetail มั้ย'
                                            bool_check_scan = "Plase_scna_detail"
                                            GoTo go_pelase_detail
                                        ElseIf check_po = 2 Then
                                            bool_check_scan = "pick_detail_number"
                                            GoTo go_pick_detail_number_fw
                                        ElseIf check_po = 3 Then
                                            bool_check_scan = "scan_ok_pickdetail"
                                            GoTo go_scan_ok_pickdetail_fw
                                        ElseIf check_po = 5 Then
                                            bool_check_scan = "ever"
                                            GoTo alert_ever
                                        ElseIf check_po = 1 Then
                                            Module1.G_show_data_supply = fa_qty + scan_qty_total
                                            inset_check_qr_part()
                                        End If
                                        number_remain = CDbl(Val(show_number_remain.Text))
                                        'เคสยิงสะสม
                                        '  MsgBox("fa_qty = " & fa_qty)
                                        ' MsgBox("scan_qty_total = " & scan_qty_total)
                                        ListBox.Items.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                        scan_qty_total = fa_qty + scan_qty_total

                                        text_tmp.Text = scan_qty_total
                                        '  MsgBox("ยอดที่คุณสแกน : " & fa_qty, 16, "Alert")
                                        check_scan = 1
                                        If show_number_supply.Text > req_qty And number_remain > 0 Then
                                            'MsgBox(fa_qty)
                                            'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                            ' Button4.Visible = True
                                            Button2.Visible = True
                                            remain_qty1 = scan_qty_total - req_qty
                                            Dim summa As Integer = fa_qty - remain_qty1
                                            check_scan = 2
                                            'remain_qty1 = scan_qty_total - req_qty.Text

                                            'Dim summa As Integer = fa_tag_qty - remain_qty1

                                            scan_qty_arrlist.Add(summa)
                                            scan_lot_arrlist.Add(fa_lot_no)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                            comp_flg = "1"
                                            firstscan = "1"
                                            check_text_box_qr_code()
                                            scan_qty.Visible = False
                                            text_box_success.Visible = True
                                            text_box_success.Focus()
                                            Panel7.Visible = True
                                            alert_success_remain.Visible = True
                                            status_alert_image = "success_remain"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                            'MsgBox(remain_qty1)
                                        ElseIf req_qty = show_number_supply.Text And number_remain = 0 Then
                                            ' MsgBox("คุณสแกนครบแล้ว", 16, "Alert")
                                            Button2.Visible = True
                                            check_scan = 2
                                            scan_qty_arrlist.Add(fa_qty)
                                            scan_lot_arrlist.Add(fa_lot_no)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                            comp_flg = "1"
                                            firstscan = "1"
                                            check_text_box_qr_code()
                                            scan_qty.Visible = False
                                            text_box_success.Visible = True
                                            text_box_success.Focus()
                                            Panel7.Visible = True
                                            alert_success.Visible = True
                                            status_alert_image = "success"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                        ElseIf show_number_supply.Text = req_qty And number_remain > 0 Then
                                            'MsgBox(fa_qty)
                                            'MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                            ' Button4.Visible = True
                                            Button2.Visible = True
                                            remain_qty1 = scan_qty_total - req_qty
                                            Dim summa As Integer = fa_qty - remain_qty1
                                            check_scan = 2
                                            scan_qty_arrlist.Add(summa)
                                            scan_lot_arrlist.Add(fa_lot_no)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                            comp_flg = "1"
                                            firstscan = "1"
                                            check_text_box_qr_code()
                                            scan_qty.Visible = False
                                            text_box_success.Visible = True
                                            text_box_success.Focus()
                                            Panel7.Visible = True
                                            alert_success_remain.Visible = True
                                            status_alert_image = "success_remain"
                                            text_box_success.Focus()
                                            GoTo exit_scan
                                        Else
                                            Button2.Visible = True
                                            scan_qty_arrlist.Add(fa_qty)
                                            scan_lot_arrlist.Add(fa_lot_no)
                                            scan_read_arrlist.Add(scan_qty.Text)
                                            scan_seq_arrlist.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                            firstscan = "1"
                                            check_text_box_qr_code()

                                        End If

                                    End If

                                    scan_qty.Text = ""
                                    scan_qty.Focus()
exit_scan:
                                End If
                            End If
                        Else
                            'MsgBox("Part incorrect")
                            status_alert_image = "Part_incorrect"
                            Panel7.Visible = True
                            alert_pa.Visible = True
                            text_box_success.Focus()
                        End If
                        'Dim item_tempp As String = scan_qty.Text.Substring(95, 5)

                    Else
                        MsgBox("Please scan tag part")

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

                        scan_qty.Text = ""
                        scan_qty.Focus()
                    End If

                End If
            Case System.Windows.Forms.Keys.Down

            Case System.Windows.Forms.Keys.F1
                Panel4.Visible = True

                user_id.Focus()
exit_keydown:
        End Select



    End Sub

    Private Sub show_qty_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles show_qty.ParentChanged

    End Sub
    Public Function Ck_dup(ByVal Lis As ListBox, ByVal Str As String)
        Dim Len_length As Integer = Len(scan_qty.Text)
        Dim tag_number As String = ""
        Dim plan_seq As String = ""
        Dim lot_sep As String = ""
        Dim tag_seq As String = ""
        Dim scan As String = ""
        scan = scan_qty.Text
        Dim num As Integer
        num = 0
        Dim count As Integer = 0
        Dim check_com_flg As String = "no data"
        Dim id As String = "no data"
        Dim qty As Integer = 0
        Dim order_number As String = ""
        Dim Code_suppier As String = "nodata"
        Dim qty_scan As Integer = 0
        Dim seq_check_reprint As String = "NODATA"
        If Len_length = 103 Then 'Fa '
            plan_seq = scan_qty.Text.Substring(16, 3)
            lot_sep = scan_qty.Text.Substring(58, 4)
            tag_number = scan_qty.Text.Substring(100, 3)
            tag_seq = scan_qty.Text.Substring(87, 16) 'plan_seq + lot_sep + tag_number
            Dim check_arr = tag_seq.Split(" ")
            Dim i As Integer = 0
            For Each value As String In check_arr
                'MsgBox(i)
                If check_arr(i) <> "" Then
                    tag_seq = check_arr(i)
                    GoTo out
                End If
                i = i + 1
            Next
out:
            'order_number = scan_qty.Text.Substring(2 , 10)
            order_number = scan_qty.Text.Substring(58, 4)
            qty_scan = scan_qty.Text.Substring(52, 6)
            seq_check_reprint = plan_seq
        ElseIf Len_length = 62 Then 'web post'
            qty_scan = scan_qty.Text.Substring(51, 8)
            Code_suppier = scan_qty.Text.Substring(37, 5)
            order_number = scan_qty.Text.Substring(2, 10)
            tag_seq = scan_qty.Text.Substring(59, 3)
            seq_check_reprint = tag_seq
        End If
        'Dim strCommand As String = "select count (id)as c from sup_scan_pick_detail where tag_readed = '" & scan & "'"
        Dim check_re = check_reprint(Module1.past_numer, order_number, tag_seq, qty_scan)
        If check_re = "HAVE_Reprint" Then 'check reprint'
            bool_check_scan = "HAVE_Reprint"
            text_tmp.Text = scan_qty_total
            scan_qty.Text = ""
            Return True
        Else
            bool_check_scan = "NO_Reprint"
        End If
        Dim strCommand3 As String = "SELECT COUNT(id) as c, com_flg  as com_flg , id as i  , scan_qty as qty FROM sup_scan_pick_detail  where item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' and scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and line_cd = '" & scan_qty.Text.Substring(2, 6) & "' and scan_qty >= '" & qty_scan & "' group by com_flg , id , scan_qty"
        ' MsgBox(strCommand3)
        Dim command3 As SqlCommand = New SqlCommand(strCommand3, myConn)
        reader = command3.ExecuteReader()
        Do While reader.Read = True
            count = reader("c").ToString()
            check_com_flg = reader("com_flg").ToString()
            id = reader("i").ToString()
            qty = reader("qty").ToString()
        Loop
        reader.Close()
        'MsgBox("c = " & count)
        '  If check_lot_scan_web_post() = True Then 'CHECK LOT ว่า ถูกต้องหรือไม่'
        Dim status As Integer = 0
        status = check_remain_in_detail_test(order_number, tag_seq, qty_scan)
        If status = 0 Then 'ตวจสอบว่า remain ว่ามีมั้ย ใน item_cd'
LOOP_INSERT:

            Dim check_detail_po As Integer = check_scan_detail_PO(order_number, Code_suppier)
            If check_detail_po = 0 Then 'check ว่า scan ถูกใน  pickdetail มั้ย'
                bool_check_scan = "Plase_scna_detail"
                Return True
            ElseIf check_detail_po = 2 Then
                bool_check_scan = "pick_detail_number"
                Return True
            ElseIf check_detail_po = 3 Then
                bool_check_scan = "scan_ok_pickdetail"
                Return True
            ElseIf check_detail_po = 4 Then
                bool_check_scan = "NO_data_tranfer"
                Return True
            End If
            If (check_com_flg = "0" Or check_com_flg = "9") Or check_com_flg = "8" Then
                Module1.check_count = 0
                If RE_check_qr() = 0 Then 'ถ้า 0 คือ insert ได้ 1 insert ไม่ได้'
                    'inset_check_qr_part()
                    bool_check_scan = "no_ever"
                    Module1.check_count = 0
                    ' Return False
                Else
                    bool_check_scan = "ever"
                    Module1.check_count = 1
                    'Return True
                End If
                'update_qty_sup_scan_pick_detail(id, qty)
                Return 0
            End If
            If count = 0 Then
                If check_qr_part_in_table() = True Then 'True หมายถึง ไม่มีข้อมูล ใน ตาราง check qr  '
                    bool_check_scan = "no_ever"
                    Module1.check_count = 0
                    ' inset_check_qr_part()
                    Return False
                Else
                    bool_check_scan = "ever"
                    Module1.check_count = 1
                    Return True
                End If
            Else
                bool_check_scan = "ever"
                Return True
            End If
        ElseIf check_remain_in_detail_test(order_number, tag_seq, qty_scan) = 2 Then 'ตวจสอบว่า remain ว่ามีมั้ย ใน item_cd'
            bool_check_scan = "HAVE_TAG_REMAIN"
            Return True
        ElseIf check_remain_in_detail_test(order_number, tag_seq, qty_scan) = 1 Then
            GoTo LOOP_INSERT
        End If
        Return True
        Return 0
    End Function
    Public Function check_reprint(ByVal item_cd As String, ByVal lot As String, ByVal seq As String, ByVal qty As String)
        Dim str_check_reprint As String = "select count(id) as c_id from sys_logs_reprint where reprint_lot = '" & lot & "' and reprint_seq = '" & seq & "' and reprint_bef = '" & qty & "' "
        Dim command3 As SqlCommand = New SqlCommand(str_check_reprint, myConn)
        reader = command3.ExecuteReader()
        Dim status As String = "NODATA"
        If reader.Read() Then
            If reader("c_id").ToString() >= "1" Then
                status = "HAVE_Reprint"
                reader.Close()
                Return status
            Else
                status = "NO_Reprint"
            End If
            reader.Close()
        Else
            reader.Close()
            status = "NO_Reprint"
        End If
        Return status
    End Function

    Public Function check_reprint_stock(ByVal l_size As String, ByVal old_qty As String, ByVal textbox As String, ByVal new_qty As String)
        If l_size = "62" Then
            Dim PO As String = textbox.Substring(2, 10)
            Dim textbox_split = textbox.Split(" ")
            Dim item_cd As String = textbox_split(0)
            Dim data_item_cd As String = item_cd.Substring(12)
            Dim sql As String = "select item_cd , qty , id ,PUCH_ODR_CD from sup_frith_in_out where item_cd ='" & data_item_cd & "' and  PUCH_ODR_CD ='" & PO & "'"
            Dim qty_double As Double = 0.0
            Dim new_qty_double As Double = 0.0
            Dim cmd As SqlCommand = New SqlCommand(sql, myConn)
            reader = cmd.ExecuteReader()
            If reader.Read() Then
                Dim QTY_STOCK = reader("qty").ToString()
                qty_double = CDbl(Val(reader("qty").ToString))
                new_qty_double = CDbl(Val(new_qty))
                reader.Close()
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
            Dim sql As String = "select fa_item_cd , fa_lot , fa_total from sup_frith_in_out_fa where fa_item_cd ='" & item_cd & "' and  fa_lot ='" & lot_fa & "'"
            Dim cmd As SqlCommand = New SqlCommand(sql, myConn)
            reader = cmd.ExecuteReader()
            If reader.Read() Then
                qty_double = CDbl(Val(reader("fa_total").ToString))
                new_qty_double = CDbl(Val(new_qty))
                Dim QTY_STOCK = reader("fa_total").ToString()
                reader.Close()
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
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'scan แบบ FA'
        Dim total_qty = text_tmp.Text - Module1.check_QTY
        Button4.Visible = True
        hidden_text_qr_code()


        Dim sel_where1 As String = Module1.wi
        Dim sel_where2 As String = Module1.past_numer
        Dim emp_cd As String = Module1.user_id
        Dim term_id As String = main.scan_terminal_id


        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyy-MM-dd HH:mm:ss"
        Dim date_now = time.ToString(format)

        Dim time_detail As DateTime = DateTime.Now
        Dim format_time_detail As String = "HH:mm:ss"
        Dim now_time_detail = time_detail.ToString(format_time_detail)

        Dim date_detail As DateTime = DateTime.Now
        Dim format_date_detail As String = "dd-MM-yyyy"
        Dim now_date_detail = date_detail.ToString(format_date_detail)
        'MsgBox(date_now)

        Dim sel_itemSpa As String = "                        "
        Dim ps = Part_No.Text.Substring(16)
        Dim part_no_detail As String = ps
        Dim get_name = Select_Line.get_Part_Name()
        Dim name_part = get_name.Substring(11)
        Dim part_name_detail As String = name_part
        Dim Model_detail As String = "  -  "
        Dim qty_detail As Integer = CDbl(Val(Module1.A_PAST_QTY)) 'req_qty.Text
        Dim remain_qty_detail As Double = remain_qty1
        Dim line_detail As String = Module1.line
        Dim loc_detail As String = location.Text.Substring(10)
        Dim user_detail As String = Module1.user_id

        Dim wi_code As String = Module1.wi

        Dim itemStrqr As String = item_cd_scan

        Dim strCount As Integer = Len(item_cd_scan)

        Dim numCountTemp As Integer = 25 - strCount

        For index As Integer = 1 To numCountTemp
            itemStrqr = itemStrqr & " "
        Next

        Dim itemNStrqr As String = part_name_detail
        Dim strNCount As Integer = Len(part_name_detail)

        Dim numNCountTemp As Integer = 25 - strNCount

        For indexN As Integer = 1 To numNCountTemp
            itemNStrqr = itemNStrqr & " "
        Next

        'MsgBox(supplier_cd)
        'MsgBox(Len(itemStrqr))

        Dim remainStr As String = supp_total_qty
        Dim total_len1 As Integer = Len(remainStr)
        Dim total_num As Integer = 6 - total_len1

        Dim testStrr As String = ""

        For index1 As Integer = 1 To total_num
            '    remainStr = total_len1 & remainStr
            testStrr = " " & testStrr
        Next

        remainStr = testStrr & remainStr

        Dim remainqtyStr As String = total_qty

        Dim total_len2 As Integer = Len(remainqtyStr)
        Dim remain_num As Integer = 6 - total_len2

        For index2 As Integer = 1 To remain_num
            remainqtyStr = " " & remainqtyStr
        Next

        '        supp_seq = scan_qty.Text.Substring(59, 3)
        Dim qr_detail_remain As String = "GD" & order_number & itemStrqr & supplier_cd & remainStr & remainqtyStr & supp_seq
        Dim date_qr_supply = now_date_detail.Split("-")
        Dim date_sup = date_qr_supply(0) & date_qr_supply(1) & date_qr_supply(2)

        Dim time_qr_supply = now_time_detail.Split(":")
        Dim time_sup = time_qr_supply(0) + time_qr_supply(1) + time_qr_supply(2)


        Dim qrdetailSupply As String = Module1.line & " " & wi_code & " " & itemStrqr & " " & Module1.check_QTY & " " & date_sup & " " & time_sup


        Dim numarrlist As Integer = scan_qty_arrlist.Count

        Try
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If
            Dim scan = scan_qty.Text
            Dim count As Integer = 0
            Dim strCommand1 As String = "select * from check_qr_part where S_number = '" & main.scan_terminal_id & "'"
            'MsgBox("strCommand1 = " & strCommand1)
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()

            'MsgBox(reader.Item(1).GetType)
            Do While reader.Read()
                F_wi.Add(reader.Item(1))
                F_item_cd.Add(reader.Item(2))
                F_scan_qty.Add(reader.Item(3))
                F_scan_lot.Add(reader.Item(4))
                F_tag_typ.Add(reader.Item(5))
                F_tag_readed.Add(reader.Item(6))
                F_scan_emp.Add(reader.Item(7))
                F_term_cd.Add(reader.Item(8))
                F_updated_date.Add(reader.Item(9))
                F_updated_by.Add(reader.Item(10))
                F_updated_seq.Add(reader.Item(11))
                F_com_flg.Add(reader.Item(13))
                F_tag_remain_qty.Add(reader.Item(14))
                F_Create_Date.Add(reader.Item(15))
                F_Create_By.Add(reader.Item(16))
                count += 1
                count_arr_fw = count_arr_fw + 1
            Loop
            reader.Close()
            Dim array_id() As Object = F_wi.ToArray()
            Dim array_item_cd() As Object = F_item_cd.ToArray()
            Dim num As Integer = 0
            For Each key In F_wi
                Dim wi As String = key
                Dim item_cd As String = F_item_cd(num)
                Dim scan_qty As String = F_scan_qty(num)
                Dim scan_lot As String = F_scan_lot(num)
                Dim tag_typ As String = F_tag_typ(num)
                Dim tag_readed As String = F_tag_readed(num)
                Dim scan_emp As String = F_scan_emp(num)
                Dim term_cd As String = F_term_cd(num)
                Dim updated_date As String = F_updated_date(num)
                Dim updated_by As String = F_updated_by(num)
                Dim updated_seq As String = F_updated_seq(num)
                Dim com_flg_table As String = F_com_flg(num)
                Dim tag_remain_qty As String = F_tag_remain_qty(num)
                Dim Create_date As String = F_Create_Date(num)
                Dim Create_By As String = F_Create_By(num)
                num += 1

                'MsgBox("data retuen  = " & item_cd)
                sup_scan_pick_detail(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, "", "", "")
            Next
            delete_data_check_qr_part()
        Catch ex As Exception
            MsgBox("Can not insert in to database detail <btn4>")
        End Try


        ' MsgBox("scan_qr_part_detail = " & Module1.scan_qr_part_detail)
        Dim firstStrscan As String = Module1.scan_qr_part_detail.Substring(0, 52)
        Dim secondStrscan As String = Module1.scan_qr_part_detail.Substring(58)
        Dim total_len3 As Integer = Len(remainqtyStr)
        Dim remain_num3 As Integer = 6 - total_len3

        For index3 As Integer = 1 To remain_num3
            remainqtyStr = " " & remainqtyStr
        Next
        'MsgBox(remainqtyStr)
        qr_detail_remain = firstStrscan & remainqtyStr & secondStrscan
        'MsgBox("qr_detail_remain=====>" & qr_detail_remain)
        'MsgBox("firstStrscan = " & firstStrscan)
        'MsgBox("remainqtyStr = " & remainqtyStr)
        'MsgBox("secondStrscan = " & secondStrscan)
        Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
        stInfoSet.addr = main.number_printter_bt
        Dim pin As StringBuilder = New StringBuilder("0000")

        Dim pinlen As UInt32 = CType(pin.Length, UInt32)
        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")

            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            Bluetooth_Print_MB200i(stInfoSet, pin, pinlen1, Module1.FG_PART_CD.Substring(16), Module1.A_ITEM_NAME, Module1.A_WI, qty_detail, Module1.A_LINE, Module1.A_USER_ID, now_date_detail, now_time_detail, qrdetailSupply)
        End If
remain_seq_FW:
        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            'ButtonF2.Enabled = False
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")

            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)

            Bluetooth_Print_MB300i(stInfoSet, pin, pinlen1, part_no_detail, Module1.past_name, Model_detail, total_qty, loc_detail, user_detail, now_date_detail, now_time_detail, qr_detail_remain)
        End If


        Try
            Dim total_pick_qty As Double = 0.0
            Dim str_plus As String = "SELECT PICK_QTY , qty FROM sup_work_plan_supply_dev WHERE line_cd  = '" & Module1.M_LINE_CD & "' AND item_cd = '" & Module1.FG_PART_CD.Substring(16) & "'AND wi  = '" & Module1.M_WI_STOP_SCAN & "' "
            Dim cmd_plus As SqlCommand = New SqlCommand(str_plus, myConn)
            reader = cmd_plus.ExecuteReader()
            Dim total_pig_qty As Double = 0.0
            Do While reader.Read()
                total_pig_qty = CDbl(Val(reader("PICK_QTY").ToString)) + CDbl(Val(Module1.check_QTY))
            Loop
            reader.Close()
            Dim strCommand As String = "UPDATE sup_work_plan_supply_dev SET update_date = '" & date_now & "' , pick_flg = '1' ,  update_by = '" & emp_cd & "' , term_cd = '" & term_id & "' , PICK_QTY = '" & total_pig_qty & "' WHERE wi  = '" & sel_where1 & "' AND item_cd = '" & sel_where2 & "'"
            'MsgBox(strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("Can not update into database")

        End Try



        'Select_Line.Line_list_view.Items.Clear()

        Try
            Dim x As ListViewItem
            Dim strCommand1 As String = "SELECT item_cd, wi, qty  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & Module1.line & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE = '" & date_now_database & "' ORDER BY wi ASC"
            'MsgBox(strCommand1)
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()
            Dim num As Integer
            num = 0

            Do While reader.Read()
                x = New ListViewItem(reader("item_cd").ToString)
                x.SubItems.Add(reader("wi").ToString)
                x.SubItems.Add(reader("qty").ToString)
                Select_Line.Line_list_view.Items.Add(x)
            Loop

            reader.Close()
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status BTN$")
        Finally
            'MsgBox("OK")
        End Try

        scan_qty_arrlist.Clear()
        scan_lot_arrlist.Clear()
        scan_read_arrlist.Clear()
        scan_seq_arrlist.Clear()


        scan_location.Location.Text = ""
        text_tmp.Text = String.Empty
        ListBox.Items.Clear()
        scan_qty.Text = String.Empty

        remain_qty.Text = ""
        remain_qty_detail = 0
        remain_qty1 = 0

        scan_qty_total = 0
        comp_flg = "0"
        firstStrscan = "0"

        scan_location.text_box_location.Focus()
        Button3.Visible = False

        Module1.total_qty = 0
        Module1.total_database = 0
        set_default_data()
        'MsgBox("End the process")

        check_process = "OK"
        set_image()
        PictureBox3.Visible = True
        text_box_success.Visible = True
        text_box_success.Focus()
        Dim ret As Int32 = 0
        ret = Bluetooth.btBluetoothSPPDisconnect()
        ret = Bluetooth.btBluetoothClose()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button2.Visible = False
        hidden_text_qr_code()
        Module1.M_QTY_STOP_SCAN = text_tmp.Text
        Dim total_qty = text_tmp.Text - Module1.check_QTY
        'Button2.Enabled = False
        Button2.Visible = False
        btn_detail_part.Visible = False
        Button1.Visible = False
        Dim sel_where1 As String = Module1.wi
        Dim sel_where2 As String = Module1.past_numer
        Dim emp_cd As String = Module1.user_id
        Dim term_id As String = main.scan_terminal_id

        'Dim part_name_detail As String = item_name.Text


        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyy-MM-dd HH:mm:ss"
        Dim date_now = time.ToString(format)

        Dim time_detail As DateTime = DateTime.Now
        Dim format_time_detail As String = "HH:mm:ss"
        Dim now_time_detail = time_detail.ToString(format_time_detail)

        Dim date_detail As DateTime = DateTime.Now
        Dim format_date_detail As String = "dd-MM-yyyy"
        Dim now_date_detail = date_detail.ToString(format_date_detail)
        'MsgBox(date_now)

        Dim part_no_detail As String = Module1.past_numer
        Dim part_name_detail As String = Module1.past_name
        Dim Model_detail As String = "  -  "
        Dim qty_detail As Integer = Module1.check_QTY
        Dim line_detail As String = Module1.line
        Dim user_detail As String = Module1.user_id


        Dim wi_code As String = Module1.wi


        Dim itemStrqr As String = item_cd_scan
        Dim strCount As Integer = Len(item_cd_scan)

        Dim numCountTemp As Integer = 25 - strCount

        For index As Integer = 1 To numCountTemp
            itemStrqr = itemStrqr & " "
        Next

        Dim itemNStrqr As String = part_name_detail
        Dim strNCount As Integer = Len(part_name_detail)

        Dim numNCountTemp As Integer = 25 - strNCount

        For indexN As Integer = 1 To numNCountTemp
            itemNStrqr = itemNStrqr & " "
        Next


        Dim sel_itemSpa As String = "                        "


        Try
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If
            Dim scan = scan_qty.Text
            Dim count As Integer = 0
            Dim strCommand1 As String = "select * from check_qr_part where S_number = '" & main.scan_terminal_id & "'"
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()

            'MsgBox(reader.Item(1).GetType)
            Do While reader.Read()
                F_wi.Add(reader.Item(1))
                F_item_cd.Add(reader.Item(2))
                F_scan_qty.Add(reader.Item(3))
                F_scan_lot.Add(reader.Item(4))
                F_tag_typ.Add(reader.Item(5))
                F_tag_readed.Add(reader.Item(6))
                F_scan_emp.Add(reader.Item(7))
                F_term_cd.Add(reader.Item(8))
                F_updated_date.Add(reader.Item(9))
                F_updated_by.Add(reader.Item(10))
                F_updated_seq.Add(reader.Item(11))
                F_com_flg.Add(reader.Item(13))
                F_tag_remain_qty.Add(reader.Item(14))
                F_Create_Date.Add(reader.Item(15))
                F_Create_By.Add(reader.Item(16))
                F_Line_cd.Add(reader.Item(18))
                F_delivery_date.Add(reader.Item(19))
                F_control_box.Add(reader.Item(21))
                count += 1
                count_arr_fw = count_arr_fw + 1
            Loop
            reader.Close()
            Dim array_id() As Object = F_wi.ToArray()
            Dim array_item_cd() As Object = F_item_cd.ToArray()
            Dim num As Integer = 0
            For Each key In F_wi
                Dim wi As String = key
                Dim item_cd As String = F_item_cd(num)
                Dim scan_qty As String = F_scan_qty(num)
                Dim scan_lot As String = F_scan_lot(num)
                Dim tag_typ As String = F_tag_typ(num)
                Dim tag_readed As String = F_tag_readed(num)
                Dim scan_emp As String = F_scan_emp(num)
                Dim term_cd As String = F_term_cd(num)
                Dim updated_date As String = F_updated_date(num)
                Dim updated_by As String = F_updated_by(num)
                Dim updated_seq As String = F_updated_seq(num)
                Dim com_flg_table As String = F_com_flg(num)
                Dim tag_remain_qty As String = F_tag_remain_qty(num)
                Dim Create_date As String = F_Create_Date(num)
                Dim Create_By As String = F_Create_By(num)
                Dim Line_cd As String = F_Line_cd(num)
                Dim delivery_date As String = F_delivery_date(num)
                Dim box_control As String = F_control_box(num)
                num += 1
                'MsgBox("data retuen  = " & item_cd)
                sup_scan_pick_detail(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, Line_cd, delivery_date, box_control)
            Next
            delete_data_check_qr_part()
        Catch ex As Exception
            MsgBox("Can not insert in to database detail <btn4>")
        End Try

        Dim date_qr_supply = now_date_detail.Split("-")
        Dim date_sup = date_qr_supply(0) & date_qr_supply(1) & date_qr_supply(2)

        Dim time_qr_supply = now_time_detail.Split(":")
        Dim time_sup = time_qr_supply(0) + time_qr_supply(1) + time_qr_supply(2)


        Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
        stInfoSet.addr = main.number_printter_bt
        Dim pin As StringBuilder = New StringBuilder("0000")

        Dim pinlen As UInt32 = CType(pin.Length, UInt32)
loop_check_open_printer:
        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            'ButtonF2.Enabled = False
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")

            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            Dim c As Integer = 0
            Dim text_temp_del_remain As Double = 0.0
            text_temp_del_remain = CDbl(Val(text_tmp.Text))
            For Each key In F_wi
                If F_com_flg(c) = "0" Then
                    text_temp_del_remain -= CDbl(Val(F_tag_remain_qty(c)))
                End If
                c += 1
            Next
            If check_scan = 1 Then
                qty_detail = text_temp_del_remain 'text_tmp.Text
            End If
            If text_temp_del_remain < Module1.check_QTY Then
                Module1.check_QTY = text_temp_del_remain
            End If

            Dim qrdetailSupply As String = "SUP-FG " & Module1.SLIP_CD & " " & Module1.FG_CUS_ORDER_ID & " " & itemStrqr & " " & Module1.G_show_data_supply & " " & Module1.A_USER_ID & " " & id_pick_log_supply & " " & date_sup & " " & time_sup

            Dim qr_detail_remain As String = "nodata"
            Dim index As Integer = 0
           ' MsgBox("----<>>>>>")
            For Each key2 In arr_pick_log
                qrdetailSupply &= " " + arr_pick_log(index)
                'MsgBox("---->" & arr_pick_log(index))
                index += 1
            Next
            Bluetooth_Print_MB200i(stInfoSet, pin, pinlen1, Module1.FG_PART_CD.Substring(16), Module1.FG_PART_NAME.Substring(12), Module1.FG_CUS_ORDER_ID, qty_detail, Module1.FG_LINE, Module1.A_USER_ID, now_date_detail, now_time_detail, qrdetailSupply)
            Dim num As Integer = 0
            For Each key In F_wi
                ''''''''''''''''''''''''''''''''''''
                If F_com_flg(num) = "0" Then
                    Dim arr_wi As String = key
                    Dim arr_item_cd As String = F_item_cd(num)
                    Dim arr_scan_qty As String = F_scan_qty(num)
                    Dim arr_scan_lot As String = F_scan_lot(num)
                    Dim arr_tag_typ As String = F_tag_typ(num)
                    Dim arr_tag_readed As String = F_tag_readed(num)
                    Dim arr_scan_emp As String = F_scan_emp(num)
                    Dim arr_term_cd As String = F_term_cd(num)
                    Dim arr_updated_date As String = F_updated_date(num)
                    Dim arr_updated_by As String = F_updated_by(num)
                    Dim arr_updated_seq As String = F_updated_seq(num)
                    Dim arr_com_flg_table As String = F_com_flg(num)
                    Dim arr_tag_remain_qty As String = F_tag_remain_qty(num)
                    Dim select_data As String = "select top 1 * from sup_work_plan_supply_dev where ITEM_CD = '" & arr_item_cd & "' and WI ='" & arr_wi & "'"
                    Dim command_get_detail As SqlCommand = New SqlCommand(select_data, myConn)
                    reader = command_get_detail.ExecuteReader()
                    Dim data_PART_NAME As String = "NO_DATA"
                    Dim data_MODEL As String = "NO_DATA"
                    Dim data_location As String = "NO_DATA"
                    If reader.Read() Then
                        data_PART_NAME = reader("ITEM_NAME").ToString
                        data_MODEL = reader("MODEL").ToString
                        data_location = reader("LOCATION_PART").ToString
                        reader.Close()
                    Else
                        reader.Close()
                    End If
                    Try
                        If Len(arr_tag_readed) = "103" Or Len(arr_tag_readed) = 103 Then
                            Dim number_re_qty As Integer = Len(arr_tag_remain_qty)
                            ' MsgBox("02")
                            Dim charArray_re_qty() As Char = arr_tag_remain_qty.ToCharArray
                            ' MsgBox("03")
                            Dim regit As Integer = 7 - number_re_qty
                            'MsgBox("04")
                            Dim n As Integer = 0
                            'MsgBox("05")
                            Dim RESULT_QTY As String = ""
                            'MsgBox("06")
                            Dim keys As Integer = 0
                            For i As Integer = 1 To 6 Step +1
                                Dim data_rigit As String = ""
                                If i = regit Then
                                    data_rigit = charArray_re_qty(keys)
                                    regit = regit + 1
                                    keys = keys + 1
                                Else
                                    data_rigit = " "
                                    n = n + 1
                                End If
                                RESULT_QTY &= data_rigit
                            Next
                            ' MsgBox("07")
                            qr_detail_remain = arr_tag_readed.Substring(0, 52) & RESULT_QTY & arr_tag_readed.Substring(58)
                        ElseIf Len(arr_tag_readed) = "62" Or Len(arr_tag_readed) = 62 Then
                            Dim keys As Integer = 0
                            Dim number_re_qty As Integer = Len(arr_tag_remain_qty)
                            Dim regit As Integer = 9 - number_re_qty
                            Dim charArray_re_qty() As Char = arr_tag_remain_qty.ToCharArray
                            Dim RESULT_QTY As String = ""
                            For i As Integer = 0 To 8 Step +1
                                Dim data_rigit As String = ""
                                If i = regit Then
                                    data_rigit = charArray_re_qty(keys)
                                    regit = regit + 1
                                    keys = keys + 1
                                Else
                                    data_rigit = "0"
                                End If
                                RESULT_QTY &= data_rigit
                            Next
                            qr_detail_remain = arr_tag_readed.Substring(0, 50) & RESULT_QTY & arr_tag_readed.Substring(59, 3)
                        End If
loop_check_open_bt:
                        'MsgBox(qr_detail_remain)
                        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
                            Module1.M_SEQ_PRINT = arr_updated_seq.Substring(0, 3) 'plan seq'
                            Panel7.Visible = False
                            alert_open_printer.Visible = False
                            Dim data_date_time = arr_updated_date.Split(" ")
                            Dim data_date = data_date_time(0)
                            Dim data_time = data_date_time(1)
                            Dim data_date_format = date_detail.ToString(data_date)
                            Dim data_time_format = date_detail.ToString(data_time)
                            'ชั่วคราว FG ''Bluetooth_Print_MB300i(stInfoSet, pin, pinlen1, arr_item_cd, data_PART_NAME, data_MODEL, arr_tag_remain_qty, data_location, arr_updated_by, data_date_format, data_time_format, qr_detail_remain)
                        Else
                            ' MsgBox("connect faill")
                            ' MsgBox("กรุณาเปิดเครื่องปริ้น")
                            Panel7.Visible = True
                            alert_open_printer.Visible = True
                            GoTo loop_check_open_bt
                        End If
                    Catch ex As Exception
                        MsgBox("ERROR Fail" & vbNewLine & ex.Message, 16, "Status ")
                    End Try
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                num += 1
            Next
        Else 'check ถ้าไม่เปิดเครื่องปริ้นให้แสดง POP UP'
            Panel7.Visible = True
            alert_open_printer.Visible = True
            GoTo loop_check_open_printer
        End If
        'MsgBox("------------------<>")
        Try
            Dim strCommand As String = ""
            Dim str_plus As String = "select * from EXP_ORDER_FG where ORDER_FG_ID = '" & Module1.FG_ORDER_ID & "' "
            ' MsgBox("0")
            Dim cmd_plus As SqlCommand = New SqlCommand(str_plus, myConn_fa)
            ' MsgBox("00")
            reader = cmd_plus.ExecuteReader()
            'MsgBox("1")
            Dim total_pig_qty As Double = 0.0
            Dim remain_qty As Double = 0.0
            Do While reader.Read()
                ' MsgBox("===>" & reader("PICK_QTY").ToString)
                ' Dim pick_system As Integer = Integer.Parse(reader("PICK_QTY").ToString)
                ' MsgBox("2")
                '  MsgBox("3")
                ' Dim pick_scan As Integer = Integer.Parse(text_tmp.Text)
                If check_scan = 1 Then
                    total_pig_qty = CDbl(Val(reader("USE_QTY").ToString)) + CDbl(Val(show_number_supply.Text))
                ElseIf check_scan = 2 Then
                    If CDbl(Val(reader("REMAIN_SHIP_QTY").ToString)) = CDbl(Val(show_number_supply.Text)) Then
                        total_pig_qty = CDbl(Val(show_number_supply.Text))
                    Else
                        total_pig_qty = CDbl(Val(reader("USE_QTY").ToString)) + CDbl(Val(show_number_supply.Text))
                    End If
                End If
                remain_qty = CDbl(Val(reader("REMAIN_SHIP_QTY").ToString)) - CDbl(Val(show_number_supply.Text))
            Loop
            reader.Close()

            If check_scan = 1 Then
                If remain_qty <= 0 Then 'check ว่า ถ้า REMAIN  = 0 ให้ไปเป็น FLG 1 '
                    check_scan = 2
                    GoTo next_flg
                End If
                strCommand = "EXEC dbo.update_data_remain_qty @use_qty = '" & total_pig_qty & "' , @remain_qty = '" & remain_qty & "' , @com_flg = '0' , @update_date = '" & date_now & "' , @update_by = '" & Module1.A_USER_ID & "', @order_id = '" & Module1.FG_ORDER_ID & "' "
            ElseIf check_scan = 2 Then
next_flg:
                strCommand = "EXEC dbo.update_data_remain_qty @use_qty = '" & total_pig_qty & "' , @remain_qty = '" & remain_qty & "' , @com_flg = '1' , @update_date = '" & date_now & "' , @update_by = '" & Module1.A_USER_ID & "', @order_id = '" & Module1.FG_ORDER_ID & "' "
            End If
            'MsgBox(strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn_fa)
            reader = command.ExecuteReader()
            reader.Close()


        Catch ex As Exception
            MsgBox("Can not update into database")
        End Try

        'Select_Line.Line_list_view.Items.Clear()

        Try
            Dim x As ListViewItem
            Dim strCommand1 As String = "SELECT item_cd, wi, qty  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & Module1.line & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE  = '" & date_now_database & "' ORDER BY wi ASC"
            'MsgBox("====>>>" & strCommand1)
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()
            Dim num As Integer
            num = 0
            Do While reader.Read()
                x = New ListViewItem(reader("item_cd").ToString)
                x.SubItems.Add(reader("wi").ToString)
                x.SubItems.Add(reader("qty").ToString)
                Select_Line.Line_list_view.Items.Add(x)
            Loop

            reader.Close()
            'selLine.scan_pick.line_cd.Text = selLine.ComboBox1.SelectedItem.ToString()
            'selLine.part_detail.line_cd.Text = selLine.ComboBox1.SelectedItem.ToString()
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status")
        Finally
            'MsgBox("OK")


        End Try

        Dim numarrlist As Integer = scan_qty_arrlist.Count

        Try
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If

            For i = 0 To numarrlist - 1

                'Dim strCommand2 As String = "INSERT INTO sup_scan_pick_detail (wi,item_cd,scan_qty,scan_lot,tag_typ,tag_readed,scan_emp,term_cd,updated_date,updated_by,tag_seq,com_flg,tag_remain_qty) VALUES ('" & Module1.wi & "','" & Module1.past_numer & "','" & scan_qty_arrlist(i) & "','" & scan_lot_arrlist(i) & "','1','" & scan_read_arrlist(i) & "','" & emp_cd & "','" & term_id & "','" & date_now & "','" & emp_cd & "','" & scan_seq_arrlist(i) & "','" & com_flg & "','" & total_qty & "')"
                'MsgBox(strCommand2)
                'Dim strCommand3 As String = "UPDATE sup_scan_pick_detail SET update_date = '" & date_now & "' , pick_flg = '1' , update_by = '" & emp_cd & "' , term_cd = '" & term_id & "' , pick_qty = '" & req_qty.Text & "'  WHERE wi  = '" & sel_where1 & "' AND item_cd = '" & sel_where2 & "'"
                'MsgBox(strCommand)
                ' Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
                ' reader = command2.ExecuteReader()
                ' reader.Close()

            Next
            delete_data_check_qr_part()
        Catch ex As Exception

        End Try
        scan_qty_arrlist.Clear()
        scan_lot_arrlist.Clear()
        scan_read_arrlist.Clear()
        scan_seq_arrlist.Clear()
        scan_location.text_box_location.Text = ""
        'scan_pick.temp_loc.Text = String.Empty
        scan_location.text_box_location.Focus()
        text_tmp.Text = String.Empty
        scan_qty.Text = String.Empty

        scan_qty_total = 0

        ListBox.Items.Clear()
        comp_flg = "0"
        firstscan = "0"

        Button2.Visible = False

        set_default_data()
        'MsgBox("End the process")
        Panel7.Visible = False
        check_process = "OK"
        set_image()
        PictureBox3.Visible = True
        text_box_success.Visible = True
        text_box_success.Focus()
        Dim ret As Int32 = 0
        ret = Bluetooth.btBluetoothSPPDisconnect()
        ret = Bluetooth.btBluetoothClose()

    End Sub
    Public Sub set_image()
        alert_pa.Visible = False
        alert_success.Visible = False
        alert_loop.Visible = False
        alert_success_remain.Visible = False
        alert_right_fa.Visible = False
        alert_reprint.Visible = False
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click 'web post'
        hidden_text_qr_code()
        set_image()
        Dim total_qty = text_tmp.Text - Module1.check_QTY
        Button3.Visible = False
        Dim req_qty = Module1.A_PAST_QTY
        Dim sel_where1 As String = Select_Line.get_wi()
        Dim ps = Part_No.Text.Substring(16)
        Dim get_name = Select_Line.get_Part_Name()
        Dim name_part = get_name.Substring(11)
        Dim code_id_user = lb_code_user.Text.Substring(6)
        Dim sel_where2 As String = ps
        Dim emp_cd As String = code_id_user
        Dim term_id As String = main.scan_terminal_id


        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyy-MM-dd HH:mm:ss"
        Dim date_now = time.ToString(format)

        Dim time_detail As DateTime = DateTime.Now
        Dim format_time_detail As String = "HH:mm:ss"
        Dim now_time_detail = time_detail.ToString(format_time_detail)

        Dim date_detail As DateTime = DateTime.Now
        Dim format_date_detail As String = "dd-MM-yyyy"
        Dim now_date_detail = date_detail.ToString(format_date_detail)
        'MsgBox(date_now)

        Dim sel_itemSpa As String = "                        "

        Dim part_no_detail As String = ps

        Dim part_name_detail As String = name_part
        Dim Model_detail As String = "  -  "
        Dim qty_detail As Integer = req_qty
        Dim remain_qty_detail As Double = remain_qty1
        Dim line_detail As String = "" 'lb_code_line.Text.Substring(6) 'waring'
        Dim loc_detail As String = location.Text
        Dim user_detail As String = code_id_user
        'part_name_detail = PD5.Part_Name.Text.Substring(11)

        Dim itemStrqr As String = item_cd_scan
        Dim strCount As Integer = Len(item_cd_scan)

        Dim numCountTemp As Integer = 25 - strCount

        For index As Integer = 1 To numCountTemp
            itemStrqr = itemStrqr & " "
        Next

        Dim itemNStrqr As String = part_name_detail
        Dim strNCount As Integer = Len(part_name_detail)

        Dim numNCountTemp As Integer = 25 - strNCount

        For indexN As Integer = 1 To numNCountTemp
            itemNStrqr = itemNStrqr & " "
        Next
        'MsgBox(supplier_cd)
        'MsgBox(Len(itemStrqr))

        Dim remainStr As String = supp_total_qty
        Dim total_len1 As Integer = Len(remainStr)
        Dim total_num As Integer = 8 - total_len1

        Dim testStrr As String = ""

        For index1 As Integer = 1 To total_num
            '    remainStr = total_len1 & remainStr
            testStrr = "0" & testStrr
        Next

        remainStr = testStrr & remainStr

        Dim remainqtyStr As String = remain_qty_detail

        Dim total_len2 As Integer = Len(remainqtyStr)
        Dim remain_num As Integer = 8 - total_len2

        For index2 As Integer = 1 To remain_num
            remainqtyStr = "0" & remainqtyStr
        Next


        Dim wi_code As String = PD_ADD_PART.ComboBox3.Text



        Dim date_qr_supply = now_date_detail.Split("-")
        Dim date_sup = date_qr_supply(0) & date_qr_supply(1) & date_qr_supply(2)

        Dim time_qr_supply = now_time_detail.Split(":")
        Dim time_sup = time_qr_supply(0) + time_qr_supply(1) + time_qr_supply(2)
        Dim qr_detail_remain As String = "GD" & order_number & itemStrqr & supplier_cd & remainStr & remainqtyStr & supp_seq 'qr remain'

        Dim qrdetailSupply As String = line_detail & " " & wi_code & " " & itemStrqr & " " & qty_detail & " " & date_sup & " " & time_sup

        'Dim sel_itemSpa As String = "                        "

        Try
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If
            Dim scan = scan_qty.Text
            Dim count As Integer = 0
            Dim strCommand1 As String = "select * from check_qr_part where S_number = '" & main.scan_terminal_id & "'"
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()

            'MsgBox(reader.Item(1).GetType)
            Do While reader.Read()
                F_wi.Add(reader.Item(1))
                F_item_cd.Add(reader.Item(2))
                F_scan_qty.Add(reader.Item(3))
                F_scan_lot.Add(reader.Item(4))
                F_tag_typ.Add(reader.Item(5))
                F_tag_readed.Add(reader.Item(6))
                F_scan_emp.Add(reader.Item(7))
                F_term_cd.Add(reader.Item(8))
                F_updated_date.Add(reader.Item(9))
                F_updated_by.Add(reader.Item(10))
                F_updated_seq.Add(reader.Item(11))
                F_com_flg.Add(reader.Item(13))
                F_tag_remain_qty.Add(reader.Item(14))
                F_Create_Date.Add(reader.Item(15))
                F_Create_By.Add(reader.Item(16))
                count += 1
                count_arr_fw = count_arr_fw + 1
            Loop
            reader.Close()
            Dim array_id() As Object = F_wi.ToArray()
            Dim array_item_cd() As Object = F_item_cd.ToArray()
            Dim num As Integer = 0
            For Each key In F_wi
                Dim wi As String = key
                Dim item_cd As String = F_item_cd(num)
                Dim scan_qty As String = F_scan_qty(num)
                Dim scan_lot As String = F_scan_lot(num)
                Dim tag_typ As String = F_tag_typ(num)
                Dim tag_readed As String = F_tag_readed(num)
                Dim scan_emp As String = F_scan_emp(num)
                Dim term_cd As String = F_term_cd(num)
                Dim updated_date As String = F_updated_date(num)
                Dim updated_by As String = F_updated_by(num)
                Dim updated_seq As String = F_updated_seq(num)
                Dim com_flg_table As String = F_com_flg(num)
                Dim tag_remain_qty As String = F_tag_remain_qty(num)
                Dim Create_date As String = F_Create_Date(num)
                Dim Create_By As String = F_Create_By(num)
                num += 1

                'MsgBox("data retuen  = " & item_cd)
                sup_scan_pick_detail(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, "", "", "")
            Next
            delete_data_check_qr_part()
        Catch ex As Exception
            MsgBox("Can not insert in to database detail <btn4>")
        End Try

        Dim stInfoSet As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
        stInfoSet.addr = main.number_printter_bt
        Dim pin As StringBuilder = New StringBuilder("0000")

        Dim pinlen As UInt32 = CType(pin.Length, UInt32)
        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            'ButtonF2.Enabled = False
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")
            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            Bluetooth_Print_MB200i(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, wi_code, qty_detail, line_detail, user_detail, now_date_detail, now_time_detail, qrdetailSupply)
        End If

        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            'ButtonF2.Enabled = False
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")

            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            loc_detail = location.Text.Substring(10)
            Bluetooth_Print_MB300i(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, Model_detail, remain_qty_detail, loc_detail, user_detail, now_date_detail, now_time_detail, qr_detail_remain)

        End If

        Try

            'Dim strCommand As String = "UPDATE sup_work_plan_supply_dev SET update_date = '" & date_now & "' , pick_flg = '1' , update_by = '" & emp_cd & "' , term_cd = '" & term_id & "' , pick_qty = '" & req_qty & "'  WHERE wi  = '" & Module1.wi & "' AND item_cd = '" & sel_where2 & "'"
            Dim str_plus As String = "SELECT PICK_QTY , qty FROM sup_work_plan_supply_dev WHERE line_cd  = '" & Module1.M_LINE_CD & "' AND item_cd = '" & Module1.FG_PART_CD.Substring(16) & "'AND wi  = '" & Module1.M_WI_STOP_SCAN & "' "
            Dim cmd_plus As SqlCommand = New SqlCommand(str_plus, myConn)
            reader = cmd_plus.ExecuteReader()
            Dim total_pig_qty As Double = 0.0
            Do While reader.Read()
                total_pig_qty = CDbl(Val(reader("PICK_QTY").ToString)) + CDbl(Val(Module1.check_QTY))
            Loop
            reader.Close()
            Dim strCommand As String = "UPDATE sup_work_plan_supply_dev SET pick_flg = '1'  , PICK_QTY = '" & total_pig_qty & "'  WHERE wi  = '" & Module1.wi & "' AND item_cd = '" & sel_where2 & "'"
            'MsgBox("strCommand = " & strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("Can not update into database")
        End Try


        Select_Line.Line_list_view.Items.Clear()

        Try
            Dim line_id = Select_Line.given_code_line()
            line_id = " " 'lb_code_line.Text.Substring(10)
            Dim x As ListViewItem
            Dim strCommand1 As String = "SELECT item_cd, wi, qty  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & Module1.line & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE = '" & date_now_database & "' ORDER BY wi ASC"
            Dim command1 As SqlCommand = New SqlCommand(strCommand1, myConn)
            reader = command1.ExecuteReader()
            Dim num As Integer
            num = 0
            'MsgBox("BTN 3 FA TAG")
            Do While reader.Read()
                x = New ListViewItem(reader("item_cd").ToString)
                x.SubItems.Add(reader("wi").ToString)
                x.SubItems.Add(reader("qty").ToString)
                Select_Line.Line_list_view.Items.Add(x)
            Loop

            reader.Close()

            reader.Close()
            'selLine.scan_pick.line_cd.Text = selLine.ComboBox1.SelectedItem.ToString()
            'selLine.part_detail.line_cd.Text = selLine.ComboBox1.SelectedItem.ToString()
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status omg")
        Finally
            'MsgBox("OK")


        End Try

        Dim numarrlist As Integer = scan_qty_arrlist.Count

        Try
            Dim wi As String = Select_Line.get_wi()
            Dim past = Part_No.Text.Substring(16)
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If

            For i = 0 To numarrlist - 1
            Next
        Catch ex As Exception
            MsgBox("Can not insert in to database detail <btn3>")
        End Try
        scan_qty_arrlist.Clear()
        scan_lot_arrlist.Clear()
        scan_read_arrlist.Clear()
        scan_seq_arrlist.Clear()
        scan_location.text_box_location.Text = ""
        text_tmp.Text = String.Empty
        ListBox.Items.Clear()
        scan_qty.Text = String.Empty
        remain_qty.Text = ""
        remain_qty_detail = 0
        remain_qty1 = 0
        scan_qty_total = 0
        comp_flg = "0"
        firstscan = "0"
        scan_location.text_box_location.Focus()
        Button3.Visible = False

        'Dim page As Page_projects = New Page_projects()
        'Dim Line As Select_Line = New Select_Line()

        'Select_Line.part = Me
        set_default_data()
        'MsgBox("End the process")
        Panel7.Visible = False
        check_process = "OK"
        set_image()
        PictureBox3.Visible = True
        text_box_success.Visible = True
        text_box_success.Focus()

        Dim ret As Int32 = 0
        ret = Bluetooth.btBluetoothSPPDisconnect()
        ret = Bluetooth.btBluetoothClose()

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
    Private Sub Bluetooth_Print_MB200i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32, ByVal part_number As String, ByVal part_name As String, ByVal wi_code As String, ByVal tag_qty As String, ByVal line_detail As String, ByVal user_detail As String, ByVal date_detail As String, ByVal time_detail As String, ByVal qrdetailSupply As String)
        Dim ret As Int32 = 0
        Dim disp As [String] = ""
        Dim sbBuf As New StringBuilder("")
        Dim ssizeGet As UInt32 = 0
        Dim bBuf As [Byte]() = New Byte() {}
        Dim rsizeGet As UInt32 = 0
        Dim bBufGet As [Byte]() = New [Byte](4094) {}

        Try
            ' Data transmission
            bBuf = New [Byte](4094) {}
            Dim bBufWork As [Byte]() = New [Byte]() {}
            Dim bBufWork_l1 As [Byte]() = New [Byte]() {}
            Dim bBufWork_l2 As [Byte]() = New [Byte]() {}
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Part No : " & Module1.FG_PART_CD.Substring(16))
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length


            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V170")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H23")
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

            ' bBufWork_l1 = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "---")
            '--------------------------------------------------------------------------------'
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "QGate --> FG")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length + bBufWork_l1.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H110")
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
            Dim result_qty As String = "NO_DATA"
            '     If tag_qty = "0" Or tag_qty = 0 Then
            'result_qty = Module1.FG_QTY.Substring(6)
            'Else
            'result_qty = tag_qty
            ' End If
            result_qty = Module1.G_show_data_supply
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Pick : " & result_qty & " pcs." & "  " & Module1.FG_MODEL)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length


            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V440")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H180")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0102")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            '       bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Name : " & "TESTTTTTTTTTTTTTTTTT")
            Dim del = Module1.delivery_date.Split()
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "DEL : " & del(0) & " | " & del(1))
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            '---------------------------------------------------------------------------------------------------------'
            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            '  len = len + bESC.Length
            '  bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            '  bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H190")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            '    bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            '   bBufWork.CopyTo(bBuf, len)
            '  len = len + bBufWork.Length

            ' bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            ' bBufWork.CopyTo(bBuf, len)
            ' len = len + bBufWork.Length
            ' bESC.CopyTo(bBuf, len)

            ' bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Instruction : XXXXXX")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length





            'bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            '---------------------------------------------------------------------------------------------------------------------'
            '           bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            '          bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            '            bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length

            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H250")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            '  bESC.CopyTo(bBuf, len)
            ' len = len + bESC.Length
            '    bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            '   bBufWork.CopyTo(bBuf, len)
            '     len = len + bBufWork.Length

            '       bESC.CopyTo(bBuf, len)
            '      len = len + bESC.Length
            '        bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            '         bBufWork.CopyTo(bBuf, len)
            '           len = len + bBufWork.Length
            '          bESC.CopyTo(bBuf, len)
            '
            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'Dim result_qty As Integer = 0
            'If tag_qty = "0" Or tag_qty = 0 Then
            '    result_qty = Module1.FG_QTY.Substring(6)
            ' Else
            '      result_qty = tag_qty
            '   End If

            '    bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Supply QTY. : " & result_qty)
            '     bBufWork.CopyTo(bBuf, len)
            '      len = len + bBufWork.Length


            '    bESC.CopyTo(bBuf, len)
            '   len = len + bESC.Length
            '  bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            ' bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H300")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            'bESC.CopyTo(bBuf, len)

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Trip : " & Module1.Trip)
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length



            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H350")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0202")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            'bESC.CopyTo(bBuf, len)

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Line : " & line_detail)
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length


            '--------------------------------------------------------------------------------------'
            '         bESC.CopyTo(bBuf, len)
            '        len = len + bESC.Length
            '       bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            '           bBufWork.CopyTo(bBuf, len)
            '          len = len + bBufWork.Length

            '            bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H350")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length

            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0101")
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            'bESC.CopyTo(bBuf, len)
            'bESC.CopyTo(bBuf, len)
            'len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Delivery Date : " & Module1.delivery_date)
            'bBufWork.CopyTo(bBuf, len)
            'len = len + bBufWork.Length
            '-------------------------------------------------------------------------------------------------------------'
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V440")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H260")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("P00")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0299")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            '   bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Model : " & "HONDA")

            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Pick Date : " & date_detail & " " & "| " & time_detail)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            '---------------------------------------------------------------------------------------'
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V440")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("L0299")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            bESC.CopyTo(bBuf, len)
            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Pick By : " & user_detail)
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            '----------------------------------------------------------------------------------------'
            If CodeType = "C128" Then
                ''// barcode
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V0071")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H0010")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("BG02060>G" & qrdetailSupply) '// code 128
                'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("BD101060" & data) '// code 39
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length
            Else
                '// QR code
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H170")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("2D30,M,4,0,0") '// qr setting
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("DS2," & qrdetailSupply) '// qr data
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length
                ''''''''''''''''''''''''''''''''''''''''''
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("%1")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V210")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H400")
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
                ' Dim status_peper As String = "NO_STATUS"
                'If check_scan = 1 Then
                'status_peper = "INCOMPLETE"
                'ElseIf check_scan = 2 Then
                'status_peper = "COMPLETED"
                'End If
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                ' bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Status : " & status_peper)
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K2B" & "Use At Phase " & Module1.FG_PHASE)
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
        Finally
        End Try
    End Sub
    Private Sub Bluetooth_Print_MB300i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32, ByVal part_number As String, ByVal part_name As String, ByVal part_model As String, ByVal remain_qty As String, ByVal loc_detail As String, ByVal user_detail As String, ByVal date_detail As String, ByVal time_detail As String, ByVal qr_detail_remain As String)
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Model : " & Module1.A_MODEL)
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
    Private Sub Bluetooth_Print_MB400i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32, ByVal part_number As String, ByVal part_name As String, ByVal part_model As String, ByVal remain_qty As String, ByVal loc_detail As String, ByVal user_detail As String, ByVal date_detail As String, ByVal time_detail As String, ByVal qr_detail_remain As String)
        Dim ret As Int32 = 0
        Dim disp As [String] = ""

        Dim sbBuf As New StringBuilder("")
        Dim ssizeGet As UInt32 = 0
        Dim bBuf As [Byte]() = New Byte() {}


        Dim rsizeGet As UInt32 = 0
        Dim bBufGet As [Byte]() = New [Byte](4094) {}

        Try


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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H250")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V180")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H210")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H300")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Revised By : " & user_detail)
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V700")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H350")
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Date : " & date_detail)
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
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V300")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length

            bESC.CopyTo(bBuf, len)
            len = len + bESC.Length
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H10")
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
            'bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & data)
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "< *** Remain Tag *** >")
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length








            If CodeType = "C128" Then
                ''// barcode
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V0071")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H0010")
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
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("V200")
                bBufWork.CopyTo(bBuf, len)
                len = len + bBufWork.Length

                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
                bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("H240")
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
        Finally
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
    Public Sub Get_img()

    End Sub

    Private Sub scan_qty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  System.Windows.Forms.Application.DoEvents()

        '        If Len(scan_qty.Text) >= 1 Then



        '       End If
    End Sub
    Public Function check_qr_part_in_table()
        Dim boll = False
        Try
            Dim Len_length As Integer = Len(scan_qty.Text)
            Dim scan As String = ""
            Dim plan_seq As String = ""
            Dim lot_sep As String = ""
            Dim tag_number As String = ""
            Dim tag_seq As String = ""
            scan = scan_qty.Text
            Dim order_number = scan_qty.Text.Substring(2, 10)

            Dim check_com_flg As String = ""
            Dim num As Integer
            num = 0
            Dim count As Integer = 0
            Dim id As String = "no data"
            Dim qty As String = "no data"
            If Len_length = 103 Then 'Fa '
                plan_seq = scan_qty.Text.Substring(16, 3)
                lot_sep = scan_qty.Text.Substring(58, 4)
                tag_number = scan_qty.Text.Substring(100, 3)
                tag_seq = scan_qty.Text.Substring(87, 16) 'plan_seq + lot_sep + tag_number
                Dim check_arr = tag_seq.Split(" ")
                Dim i As Integer = 0
                For Each value As String In check_arr
                    'MsgBox(i)
                    If check_arr(i) <> "" Then
                        tag_seq = check_arr(i)
                        GoTo out
                    End If
                    i = i + 1
                Next
out:
                order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
            ElseIf Len_length = 62 Then 'web post'
                order_number = scan_qty.Text.Substring(2, 10)
                tag_seq = scan_qty.Text.Substring(59, 3)
            End If
            Dim strCommand As String = "SELECT COUNT(id) as c, com_flg  as com_flg , id as i  , scan_qty as qty FROM check_qr_part  where item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' and scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and line_cd = '" & scan_qty.Text.Substring(2, 6) & "'group by com_flg , id , scan_qty"
            'MsgBox("strCommand = " & strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            'MsgBox("strCommand == >" & strCommand)
            'Dim LI As New ListViewItem 'new obj ''
            Do While reader.Read = True
                count = reader("c").ToString()
                check_com_flg = reader("com_flg").ToString()
                id = reader("i").ToString()
                qty = reader("qty").ToString()
            Loop
            reader.Close()
            ' MsgBox("function check_qr_part = " & count)

            If check_com_flg = "0" Then
                'MsgBox("UPDATE")
                Module1.check_count = 0
                inset_check_qr_part()
                '  update_qty_check_qr_part(id, qty)
                Return 0
            End If

            If count = 0 Then
                Module1.check_count2 = 0
                boll = True
            Else
                text_tmp.Text = scan_qty_total
                ' MsgBox("value123 = " & scan_qty_total)
                check_count__data = scan_qty_total
                Module1.check_count2 = 1
                boll = False
            End If

        Catch ex As Exception
            MsgBox("SORRY FUNCTION check_qr_part ERROR!!!!")
        End Try
        Return boll
    End Function

    Public Sub inset_check_qr_part()
        Dim Len_length As Integer = Len(scan_qty.Text)
        Len_length_QR = Len_length
        Dim strCommand2 As String = " no data"
        Try
            Dim com_flg As Integer = 0
            Dim S_number As String = main.scan_terminal_id
            Dim order_number = scan_qty.Text.Substring(2, 10)
            Dim supp_seq = scan_qty.Text.Substring(59, 3)
            Dim tag_seq = order_number & supp_seq
            Dim scan_qr = scan_qty.Text()
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            ' Dim total_qty = text_tmp.Text - Module1.check_QTY
            Dim best_qty = show_qty.Text.Substring(6)
            Dim qty_s As Integer = 0
            Dim total_qty As Integer = 0
            Dim t As Integer = 0
            Dim plan_seq As String = ""
            Dim lot_sep As String = ""
            Dim tag_number As String = ""
            ' MsgBox("number scan = " & qty_s)
            If Len_length = 103 Then 'Fa '
                'MsgBox("OK0_1")
                qty_s = scan_qty.Text.Substring(55, 3)
                'MsgBox("OK0_2")
                plan_seq = scan_qty.Text.Substring(16, 3)
                'MsgBox("OK0_3")
                lot_sep = scan_qty.Text.Substring(58, 4)
                'MsgBox("OK0_4")
                tag_number = scan_qty.Text.Substring(100, 3)
                'MsgBox("OK0_5")
                'MsgBox("OK0_6")

                tag_seq = scan_qr.Substring(87, 16) 'plan_seq & lot_sep & tag_number
                Dim check_arr = tag_seq.Split(" ")
                Dim i As Integer = 0
                For Each value As String In check_arr
                    'MsgBox(i)
                    If check_arr(i) <> "" Then
                        tag_seq = check_arr(i)
                        GoTo out
                    End If
                    i = i + 1
                Next
out:
                order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
                'MsgBox("OK0_7")
            ElseIf Len_length = 62 Then 'web post'
                'MsgBox("OK0_8")
                qty_s = scan_qty.Text.Substring(51, 8)
                'MsgBox("OK0_9")
                'MsgBox("OK0_10")
                tag_seq = scan_qty.Text.Substring(59, 3)
                'MsgBox("OK0_11")
                'order_number = Module1.M_LOT 'LOT มาจาก page scan location'
                order_number = scan_qty.Text.Substring(2, 10)
            End If
            If Module1.total_database = 0 Then
                'MsgBox("OK0_12")
                t = best_qty - qty_s
                'MsgBox("OK0_13")
                Module1.total_qty = best_qty - qty_s
                'MsgBox("best_qty = " & best_qty)
                'MsgBox("qty_s = " & qty_s)
                'MsgBox("OK0_14")
                Module1.total_database = Module1.total_database + 1
                'MsgBox("OK0_15")
                'MsgBox("ค่าคงเหลือเริ่มต้น = " & Module1.total_qty)
                'MsgBox("OK0_16")
            Else
                'MsgBox("รอบมากกว่า1")
                Module1.total_qty = Module1.total_qty - qty_s
                t = Module1.total_qty
                'MsgBox("ค่าคงเหลือรอบปัจจุบัน = " & Module1.total_qty)
                Module1.total_database = Module1.total_database + 1
            End If
            'MsgBox("OK1")
            ' MsgBox(" Module1.total_qty = " & Module1.total_qty)
            If Module1.total_qty = 0 Then
                'MsgBox("if 1")
                t = 0
                com_flg = 1
                'MsgBox("if 2")
                Dim pase_number_qty_scan As Integer = 0
                'MsgBox("if 3")
                If Len(scan_qty.Text) = 62 Then
                    ' MsgBox("WEB POST")
                    pase_number_qty_scan = scan_qty.Text.Substring(51, 8)
                    'tag_seq = order_number + tag_seq
                    tag_seq = scan_qty.Text.Substring(59, 3)
                    text_tmp.Text = pase_number_qty_scan
                    ' order_number = Module1.M_LOT 'LOT มาจาก page scan location'
                    order_number = scan_qty.Text.Substring(2, 10)
                ElseIf Len(scan_qty.Text) = 103 Then
                    ' MsgBox("FA")
                    ' MsgBox("tag_seq = " & tag_seq)
                    order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
                End If
                'MsgBox("if 4")

                'MsgBox("เพิ่มมาใหม่ BTN2")
            End If 'เพิ่มมาใหม่ BTN2 '
            'MsgBox("order_number = " & order_number)
            If Module1.total_qty > 0 Then
                t = 0
                com_flg = 1
                'MsgBox("เพิ่มมาใหม่ BTN4")
            End If
            'MsgBox("OK3")
            Dim t_string As String = "no data"
            If t < 0 Then
                t_string = t
                t = t_string.Substring(1)
                ' MsgBox("condition t = " & t)
            End If
            'MsgBox("text_tmp = " & text_tmp.Text)
            '  MsgBox("QTY_INSERT_LOT_PO = " & QTY_INSERT_LOT_PO)
            ' MsgBox("t = " & t)
            If QTY_INSERT_LOT_PO > 0 Then
                com_flg = 0
                t = QTY_INSERT_LOT_PO
                QTY_INSERT_LOT_PO = 0 'set = 0'
            ElseIf QTY_INSERT_LOT_PO <= 0 Then
                com_flg = 1
                t = 0
                QTY_INSERT_LOT_PO = 0 'set = 0'
            End If

            Module1.show_data_remain = CDbl(Val(show_qty.Text.Substring(6))) - CDbl(Val(Module1.G_show_data_supply)) 'show RM'
            Module1.show_data_supply = Module1.G_show_data_supply 'text_tmp.Text
            set_show_remain() 'show RM'
            set_show_supply()
            Dim box_control As Integer = 0
            Dim status_box As Integer = 0
            'select box '
            Try
                Dim check_box = "SELECT CASE WHEN MAX (BOX_CONTROL)  IS NULL THEN 0 ELSE MAX (BOX_CONTROL) END AS box FROM sup_scan_pick_detail WHERE SLIP_CD = '" & Module1.SLIP_CD & "'"
                ' MsgBox("SUP_SCAN_PICKK_DETAIL" & check_box)
                Dim cmd_check_box As SqlCommand = New SqlCommand(check_box, myConn)
                reader = cmd_check_box.ExecuteReader()
                If reader.Read() Then
                    If reader("box").ToString() = "0" Then
                        'MsgBox("testtttt")
                        Do While reader.Read
                            ' MsgBox(reader("box").ToString())
                            'MsgBox("STOP NAJA")
                        Loop
                        status_box = 1
                        reader.Close()
                        GoTo check_box_part
                    Else
                        box_control = CDbl(Val(reader("box").ToString())) + 1
                        reader.Close()
                    End If

                End If
check_box_part:
                If status_box = 1 Then
                    'reader.Close()
                    Dim check_box_part = "SELECT CASE WHEN MAX (BOX_CONTROL) IS NULL THEN 0 ELSE MAX (BOX_CONTROL) END AS box FROM check_qr_part WHERE SLIP_CD = '" & Module1.SLIP_CD & "'"
                    'MsgBox("check_box = " & check_box)
                    Dim cmd_check_box_part As SqlCommand = New SqlCommand(check_box_part, myConn)
                    reader = cmd_check_box_part.ExecuteReader()
                    ' MsgBox("ttt")
                    If reader.Read() Then
                        ' MsgBox(reader("box").ToString())
                        If reader("box").ToString() = "0" Then
                            box_control = box_control + 1
                        Else
                            box_control = CDbl(Val(reader("box").ToString())) + 1
                        End If
                    End If
                End If
                reader.Close()
            Catch ex As Exception

            End Try
            'end box '
            strCommand2 = "INSERT INTO check_qr_part (wi,item_cd,scan_qty,scan_lot,tag_typ,tag_readed,scan_emp,term_cd,updated_date,updated_by,tag_seq,S_number , com_flg ,tag_remain_qty  , CREATE_DATE , CREATE_BY , MENU_ID , line_cd , Delivery_date , SLIP_CD , BOX_CONTROL) VALUES ('" & Module1.FG_CUS_ORDER_ID & "','" & Module1.FG_PART_CD.Substring(16) & "','" & text_tmp.Text & "','" & order_number & "','1','" & scan_qr & "','" & Module1.A_USER_ID & "','" & S_number & "','" & date_now & "','" & Module1.A_USER_ID & "','" & tag_seq & "','" & S_number & "','" & com_flg & "','" & t & "' , '" & date_now & "' , '" & Module1.A_USER_ID & "', '" & Module1.MENU_ID & "' , '" & scan_qty.Text.Substring(2, 6) & "' , '" & Module1.delivery_date & "' , '" & Module1.SLIP_CD & "', '" & box_control & "')"
            '           MsgBox("TEST OPEN")
            reader.Close()
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            '            MsgBox("TEST OPEN3")
            reader.Close()
            arr_check_qr_remain_lot.Add(order_number)
            arr_check_qr_remain_seq.Add(tag_seq)
            If Len_length = 62 Then
                count_scan = count_scan + 1
                check_lot_scan_web_post()
            Else
                ' MsgBox("NO FUNCTION FA")
                check_lot_scan_fw()
            End If
            'ordernumber คือ lot'

        Catch ex As Exception
            MsgBox("SORRY Insert ERROR" & vbNewLine & ex.Message)
            MsgBox(strCommand2)
        End Try
    End Sub
    Public Function sup_scan_pick_detail(ByVal count As String, ByVal F_wi As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal scan_lot As String, ByVal tag_typ As String, ByVal tag_readed As String, ByVal scan_emp As String, ByVal term_cd As String, ByVal updated_date As String, ByVal updated_by As String, ByVal updated_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal create_date As String, ByVal create_by As String, ByVal line_cd As String, ByVal delivery_date As String, ByVal box_control As String)
        ', ByVal F_item_cd As String, ByVal F_scan_qty As String, ByVal F_scan_lot As String, ByVal F_tag_typ As String, ByVal F_tag_readed As String, ByVal F_scan_emp As String, ByVal F_term_cd As String, ByVal F_updated_date As String, ByVal F_updated_by As String, ByVal F_updated_seq As String, ByVal com_flg As String, ByVal total_qty As String
        Dim Len_length As Integer = length
        Dim strCommand2 As String = "no data"
        Dim PO = scan_lot 'lot คือ PO'
        Try
            'MsgBox("ready insert OR update")
            'MsgBox("scan_lot = " & scan_lot & "\n updated_seq = " & updated_seq & "\n F_item_cd" & F_item_cd)
            If check_remain(scan_lot, updated_seq, F_item_cd) = 1 Then
                'MsgBox("check_remain = 1")
                If REMAIN_ID <> "NO_DATA" Then
                    ' MsgBox("check update_qty_sup_scan_pick_detail(" & REMAIN_ID & tag_remain_qty & ")")
                    update_qty_sup_scan_pick_detail(REMAIN_ID, tag_remain_qty, scan_qty, updated_date, updated_by)
                    MsgBox(REMAIN_ID)
                    If Len_length = 62 Then
                        WEB_POST_Cut_stock_frith_in_out(PO, F_item_cd, scan_qty, tag_readed, updated_seq, com_flg_table, tag_remain_qty)
                    ElseIf Len_length = 103 Then
                        brak_loop = 0
                        For Each key3 In Module1.arr_check_QTY_scan
                            brak_loop = brak_loop + 1
                        Next
                        FW_Cut_stock_frith_in_out(PO, F_item_cd, scan_qty, tag_readed, updated_seq, com_flg_table, tag_remain_qty, scan_lot)
                    End If
                End If
            Else
                strCommand2 = "INSERT INTO sup_scan_pick_detail (wi , item_cd , scan_qty ,scan_lot , tag_typ , tag_readed , scan_emp, term_cd , updated_date , updated_by , tag_seq  , com_flg , tag_remain_qty , CREATE_DATE , CREATE_BY , MENU_ID , line_cd , delivery_date , SLIP_CD , BOX_CONTROL ) VALUES ('" & F_wi & "' ,'" & F_item_cd & "','" & scan_qty & "' ,'" & scan_lot & "','" & tag_typ & "','" & tag_readed & "','" & scan_emp & "','" & term_cd & "','" & updated_date & "','" & updated_by & "','" & updated_seq & "','" & com_flg_table & "','" & tag_remain_qty & "' , '" & create_date & "' , '" & create_by & "', '" & Module1.MENU_ID & "' , '" & line_cd & "', '" & delivery_date & "' , '" & Module1.SLIP_CD & "', '" & box_control & "')"
                Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
                reader = command2.ExecuteReader()
                reader.Close()

                If Len_length = 62 Then
                    WEB_POST_Cut_stock_frith_in_out(PO, F_item_cd, scan_qty, tag_readed, updated_seq, com_flg_table, tag_remain_qty)
                ElseIf Len_length = 103 Then
                    brak_loop = 0
                    For Each key3 In Module1.arr_check_QTY_scan
                        brak_loop = brak_loop + 1
                    Next
                    FW_Cut_stock_frith_in_out(PO, F_item_cd, scan_qty, tag_readed, updated_seq, com_flg_table, tag_remain_qty, scan_lot)
                End If
            End If
            Dim used_qty As Double = 0.0
            used_qty = CDbl(Val(scan_qty)) - CDbl(Val(tag_remain_qty))
            insert_pick_log(REMAIN_ID, F_wi, used_qty, create_date, create_by, updated_date, updated_by, scan_lot, updated_seq, F_item_cd)
            If tag_readed.Length = "103" Then 'ตัด เฉพาะ FG'\
                cut_stock_FASYSTEM(used_qty, F_item_cd, updated_seq, scan_lot, tag_readed)
            End If
        Catch ex As Exception
            MsgBox("ERROR sup_scan_pick_detail Insert " & vbNewLine & ex.Message, 16, "ALERT")
            MsgBox("data sql  = " & strCommand2)
        End Try
        Return 0
    End Function

    Public Sub delete_data_check_qr_part()
        Try
            Dim S_number As String = main.scan_terminal_id
            Dim strCommand2 As String = "delete from check_qr_part where S_number = '" & S_number & "'"
            'MsgBox(strCommand2)
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("SORRY Delete ERROR Function delete_data_check_qr_part")
        End Try

    End Sub
    Public Sub Re_scan()
        'MsgBox("Scan ซ้ำ! มีการสแกนแล้วเมื่อสักครู่", 16, "Alert")
        Panel7.Visible = True
        alert_loop.Visible = True
        status_alert_image = "loop_re_scan"
        text_box_success.Focus()

    End Sub
    Public Sub Re_scan_fa()
        'MsgBox("Scan ซ้ำ!!!!! มีการสแกนแล้วเมื่อสักครู่ ครับผม ", 16, "Alert")
        'MsgBox("data = " & scan_qty_total)
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
        ' Bt.SysLib.Device.btBuzzer(1, stBuz)
        ' Bt.SysLib.Device.btVibrator(1, stVib)
        ' Bt.SysLib.Device.btLED(1, stLed)
        ' MsgBox("01")
        Panel7.Visible = True
        ' MsgBox("02")
        alert_loop.Visible = True
        'MsgBox("03")
        status_alert_image = "loop_re_scan_fa"
        ' MsgBox("04")
        text_box_success.Visible = True
        'MsgBox("05")
        text_box_success.Focus()
    End Sub
    Public Sub Re_scan_default()
        Dim Len_length As Integer = Len(scan_qty.Text)
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
        'text_tmp.Text = Module1.SCAN_QTY_TOTAL
        'MsgBox("check_count__data = " & check_count__data)
        If count_scan = 0 Then
            text_tmp.Text = 0
        Else
            Dim total As Integer = 0
            If Len_length = 62 Then
                total = scan_qty_total
            ElseIf Len_length = 103 Then
                total = scan_qty_total
            End If
            text_tmp.Text = total
        End If
        scan_qty.Text = ""
        scan_qty.Focus()
    End Sub
    Public Sub update_qty_sup_scan_pick_detail(ByVal id As String, ByVal qty_database As Integer, ByVal scan_qty_default As Double, ByVal up_date_date As String, ByVal up_date_by As String)
        Dim Len_length As Integer = Len(scan_qty.Text)
        'MsgBox("0")
        Dim S_number As String = main.scan_terminal_id
        'MsgBox("00")
        Dim qty_s As Integer = 0
        If Len_length = 103 Then 'Fa '
            'MsgBox("OK0_1")
            qty_s = scan_qty.Text.Substring(55, 3)
        ElseIf Len_length = 62 Then 'web post'
            'MsgBox("OK0_8")
            qty_s = scan_qty.Text.Substring(51, 8)
        End If
        'MsgBox("000")
        Dim sum_qty = qty_database - qty_s
        'MsgBox("0000")
        'MsgBox(qty_database & "-" & qty_s & "= " & sum_qty)
        'MsgBox("00000")
        Dim com_flg As String = "no data"
        'MsgBox("001")
        Try
            If sum_qty > 0 Then
                sum_qty = 0
                com_flg = 1
            End If
            'MsgBox("002")
            If Module1.M_check_remain = "have_data" Then
                Dim remain As Double = 0.0
                remain = CDbl(Val(qty_database))
                If remain <= 0 Then
                    com_flg = 1
                    sum_qty = 0
                Else
                    com_flg = 0
                    sum_qty = remain
                End If

            End If
            'MsgBox("003")
            reader.Close()
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)

            Dim str_update As String = "update sup_scan_pick_detail set com_flg  = '" & com_flg & "' ,tag_remain_qty = '" & sum_qty & "'  , updated_date = '" & up_date_date & "' , updated_by = '" & up_date_by & "' where id = '" & id & "'"
            MsgBox(str_update)
            'MsgBox("004")
            Dim command2 As SqlCommand = New SqlCommand(str_update, myConn)
            'MsgBox("005")
            reader = command2.ExecuteReader()
            'MsgBox("006")
            reader.Close()
            'MsgBox("007")
            ' MsgBox("UPDATE OK ===>" & str_update)
        Catch ex As Exception
            MsgBox("SORRY Update ERROR Function update_qty_sup_scan_pick_detail" & vbNewLine & ex.Message, 16, "Status")
        End Try
    End Sub
    Public Sub update_qty_check_qr_part(ByVal id As String, ByVal qty_database As Integer)
        Dim S_number As String = main.scan_terminal_id
        Dim qty_s As Integer = scan_qty.Text.Substring(55, 3)
        Dim sum_qty = qty_database - qty_s
        Dim com_flg As String = "no data"
        Try
            If sum_qty > 0 Then
                sum_qty = 0
                com_flg = 1
            End If
            Dim strCommand2 As String = "update check_qr_part set com_flg  = '" & com_flg & "' ,tag_remain_qty = '" & sum_qty & "' where id = '" & id & "'"
            'MsgBox(strCommand2)
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
            'MsgBox("UPDATE OK update_qty_check_qr_part  ===>" & strCommand2)
        Catch ex As Exception
            MsgBox("SORRY Update ERROR Function update_qty_check_qr_part" & vbNewLine & ex.Message, 16, "Status")
        End Try
    End Sub
    Public Function check_lot_scan_web_post()
        Dim SCAN = scan_qty.Text
        Dim PO As String = SCAN.Substring(2, 10)
        Dim Code_suppier As String = SCAN.Substring(37, 5)
        Dim QTY_SCAN As Integer = scan_qty.Text.Substring(51, 8)
        Dim sql_c2 As String = "SELECT DISTINCT c.ITEM_CD, c.PO, c.CODE_SUPPIER, c.LT FROM ( SELECT SUBSTRING (AB.LOT_RECEIVE, 0, 6) AS CODE_SUPPIER, AB.com_flg, AB.ITEM_CD AS ITEM_CD, AB.PUCH_ODR_CD AS PO, AB.LOT_RECEIVE AS LT FROM sup_frith_in_out AS AB ) c WHERE c.com_flg = 0 AND c.PO = '" & PO & "' AND c.ITEM_CD = '" & Module1.FG_PART_CD.Substring(16) & "' AND c.CODE_SUPPIER = '" & Code_suppier & "'"
        'MsgBox("sql___DATA" & sql_c2)
        Dim command_c2 As SqlCommand = New SqlCommand(sql_c2, myConn)
        reader = command_c2.ExecuteReader()
        Do While reader.Read = True
            Module1.arr_check_PO_scan.Add(reader("PO").ToString())
            Module1.arr_check_lot_scan.Add(reader("LT").ToString())
            Module1.arr_check_QTY_scan.Add(QTY_SCAN)
        Loop
        reader.Close()
        Module1.check_pick_detail = 1
        Return 0
    End Function
    Public Function check_lot_scan_fw()
        Try
            Dim SCAN = scan_qty.Text
            Dim lot As String = SCAN.Substring(58, 4)
            Dim QTY_SCAN As Integer = scan_qty.Text.Substring(52, 6)
            ' Dim sql_c2 As String = "SELECT f_fa.fa_id as id ,COUNT (f_fa.fa_id) AS c, SUM (f_fa.fa_total) AS QTY_OF_LOT, f_fa.fa_lot as LT , f_fa.fa_total FROM sup_work_plan_supply_dev sd LEFT JOIN sup_frith_in_out_fa f_fa ON sd.ITEM_CD = f_fa.fa_item_cd WHERE f_fa.fa_ITEM_CD = '" & Module1.A_PAST_NO & "' and f_fa.fa_lot ='" & lot & "' AND sd.PICK_FLG = '0' AND f_fa.fa_com_flg = '0' GROUP BY f_fa.fa_lot , f_fa.fa_total ,  f_fa.fa_id"
            'Dim command_c2 As SqlCommand = New SqlCommand(sql_c2, myConn)
            'reader = command_c2.ExecuteReader()

            ' If reader.Read Then
            Module1.arr_check_PO_scan.Add("I AM FG")
            Module1.arr_check_lot_scan.Add(lot)
            Module1.arr_check_QTY_scan.Add(CDbl(Val(QTY_SCAN)))
            ' End If
            ' reader.Close()
            Module1.check_pick_detail = 1
        Catch ex As Exception
            reader.Close()
            MsgBox("SELECT FAILL" & vbNewLine & ex.Message, 16, "ALERT")
        End Try

        Return 0
    End Function
    Public Sub check_text_box_qr_code()
        'scan_qty.Visible = False
    End Sub
    Public Sub hidden_text_qr_code()
        scan_qty.Visible = False
    End Sub
    Public Sub set_default_data()
        QTY_INSERT_LOT_PO = 0.0
        Module1.SCAN_QTY_TOTAL = 0.0
        Module1.show_data_supply = 0.0
        Module1.show_data_remain = 0.0
        Module1.G_show_data_supply = 0.0
        Module1.total_qty = 0
        Module1.total_database = 0
        Module1.check_pick_detail = 0
        Module1.M_CHECK_LOT_COUNT_FW = New ArrayList()
        Module1.arr_pick_detail_po = New ArrayList()
        Module1.arr_pick_detail_qty = New ArrayList()
        Module1.arr_pick_detail_lot = New ArrayList()
        Module1.arr_check_lot_scan = New ArrayList()
        Module1.arr_check_PO_scan = New ArrayList()
        Module1.arr_check_QTY_scan = New ArrayList()
    End Sub

    Private Sub show_detail_scan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        detail_scan.Show()
    End Sub

    Private Sub lot_no_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lot_no.ParentChanged

    End Sub
    Public Sub SetConnect()
        myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=FASYSTEM;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Module1.total_qty = 0
        Module1.total_database = 0
        delete_data_check_qr_part()
        PD_ADD_PART.Show()
        Me.Close()
    End Sub
    Public Sub Cut_stock_lot()

    End Sub

    Private Sub btn_detail_part_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_detail_part.Click
        'Dim page_PO_NO As PO_NO = New PO_NO()
        'page_PO_NO.Show()
        get_data_tetail()
    End Sub

    Public Function check_remain(ByVal scan_lot As String, ByVal updated_seq As String, ByVal F_item_cd As String)
        Try
            Dim count As Integer = check_count(scan_lot, updated_seq, F_item_cd)
            Dim status As Integer = 0
            If count <> "0" Then
                Dim strCommand123 As String = "SELECT COUNT(id) as c,  id as i   FROM sup_scan_pick_detail  where scan_lot = '" & scan_lot & "' and tag_seq = '" & updated_seq & "' AND com_flg = '0' and item_cd ='" & F_item_cd & "'  GROUP BY id"
                'MsgBox("strCommand1_bast == >" & strCommand123)
                ' MsgBox("000")
                Dim command123 As SqlCommand = New SqlCommand(strCommand123, myConn)
                'MsgBox("999")
                reader = command123.ExecuteReader() 'ติดบรรทัดนี้'
                'MsgBox("555")
                Dim id As Integer = 0
                'MsgBox("READY LOOP")
                Do While reader.Read = True
                    If reader("c").ToString() = "0" Then
                        REMAIN_ID = "NO_DATA"
                        status = 0
                    Else
                        Module1.M_check_remain = "have_data"
                        REMAIN_ID = reader("i").ToString()
                        status = 1
                    End If
                Loop
                reader.Close()
            Else
                REMAIN_ID = "NO_DATA"
                status = 0
            End If
            Return status
        Catch ex As Exception
            MsgBox("ERROR QUERY DATA NULL" & vbNewLine & ex.Message, 16, "Status")
        End Try

        Return 0
    End Function
    Public Function check_count(ByVal scan_lot As String, ByVal updated_seq As String, ByVal F_item_cd As String)
        Dim count As String = "0"
        Dim sql_check As String = "SELECT COUNT(id) as c FROM sup_scan_pick_detail  where scan_lot = '" & scan_lot & "' and tag_seq = '" & updated_seq & "' AND com_flg = '0' and item_cd ='" & F_item_cd & "'"
        Dim command2 As SqlCommand = New SqlCommand(sql_check, myConn)
        reader = command2.ExecuteReader()
        Do While reader.Read = True
            count = reader("c").ToString()
        Loop
        reader.Close()
        'MsgBox("count = " & count)
        'MsgBox("READY IF ")
        Return count
    End Function
    Public Function RE_check_qr()
        Dim scan As String = ""
        Dim plan_seq As String = ""
        Dim lot_sep As String = ""
        Dim tag_number As String = ""
        Dim tag_seq As String = ""
        scan = scan_qty.Text
        Dim order_number = scan_qty.Text.Substring(2, 10)
        Dim Len_length As Integer = Len(scan_qty.Text)
        Dim qty As Integer = 0
        If Len_length = 103 Then 'Fa '
            plan_seq = scan_qty.Text.Substring(16, 3)
            lot_sep = scan_qty.Text.Substring(58, 4)
            tag_number = scan_qty.Text.Substring(100, 3)
            tag_seq = scan_qty.Text.Substring(87, 16) 'plan_seq + lot_sep + tag_number
            Dim check_arr = tag_seq.Split(" ")
            Dim i As Integer = 0
            For Each value As String In check_arr
                'MsgBox(i)
                If check_arr(i) <> "" Then
                    tag_seq = check_arr(i)
                    GoTo out
                End If
                i = i + 1
            Next
out:
            order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
            qty = scan_qty.Text.Substring(52, 6)
        ElseIf Len_length = 62 Then 'web post'
            order_number = scan_qty.Text.Substring(2, 10)
            tag_seq = scan_qty.Text.Substring(59, 3)
            qty = scan_qty.Text.Substring(51, 8)
        End If
        Dim count As String = "0"
        Dim check_com_flg As String = "NO_DATA"
        Dim id As String = "NO_DATA"
        'reader.Close()
        Dim strCommand As String = "SELECT COUNT(id) as c, com_flg  as com_flg   FROM check_qr_part   where item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' and scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and line_cd = '" & scan_qty.Text.Substring(2, 6) & "'and com_flg = '1' group by com_flg"
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        ' MsgBox("strCommand == >" & strCommand)
        'Dim LI As New ListViewItem 'new obj ''
        Do While reader.Read = True
            count = reader("c").ToString()
        Loop
        reader.Close()
        Dim check_reprint_c As Integer = 0

        If count = 1 Then
            text_tmp.Text = scan_qty_total
            ' MsgBox("value = " & scan_qty_total)
        End If
        check_count__data = scan_qty_total
        ' MsgBox("recheck check_qr_part = " & count)
        Return count
    End Function
    Public Sub WEB_POST_Cut_stock_frith_in_out(ByVal PO As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal tag_readed As String, ByVal tag_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String)
        Try
            Dim Code_suppier As String = "no_data"
            Code_suppier = tag_readed.Substring(37, 5)
            Dim qty_stock As String = "NODATA"
            Dim id As String = "NO_DATA"
            Dim qty_check As Integer = 0
            Dim qty_check_remain As Integer = 0
            Dim total_qty As Integer = 0
            Dim count As String = "0"

            Dim strCommand As String = "SELECT DISTINCT c.id , c.ITEM_CD, c.PO, c.CODE_SUPPIER, c.LT  , c.id as id , c.qty as qty FROM ( SELECT SUBSTRING (AB.LOT_RECEIVE, 0, 6) AS CODE_SUPPIER, AB.com_flg, AB.ITEM_CD AS ITEM_CD, AB.PUCH_ODR_CD AS PO, AB.LOT_RECEIVE AS LT  , AB.qty as qty  , AB.id as id  FROM sup_frith_in_out AS AB ) c WHERE c.com_flg = 0 AND c.PO = '" & PO & "' AND c.ITEM_CD = '" & F_item_cd & "' AND c.CODE_SUPPIER = '" & Code_suppier & "' "
            'MsgBox("strCommand =====<><><>>>>>>" & strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            'MsgBox("strCommand WEB == >" & strCommand)
            Do While reader.Read = True
                id = reader("id").ToString()
                qty_stock = reader("qty").ToString()
            Loop
            reader.Close()
            If com_flg_table = "1" Then
                total_qty = qty_stock - scan_qty
                '  MsgBox("ไม่มี REMAIN")
            ElseIf com_flg_table = "0" Then
                ' MsgBox("เหลือ REMAIN")
                total_qty = scan_qty - tag_remain_qty
                total_qty = qty_stock - total_qty
            End If
            Dim com_flg As Integer = 0
            If total_qty <= 0 Then
                total_qty = 0
                com_flg = 1
            End If


            Dim strCommand2 As String = "update sup_frith_in_out set qty  = '" & total_qty & "' ,com_flg = '" & com_flg & "' where id = '" & id & "'"
            'MsgBox("UPDATE WEBPOST = " & strCommand2)
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
            'MsgBox("COMPLETE CUT STOCK")
        Catch ex As Exception
            MsgBox("ERROR UPDATE CUT_STOCK FAILL " & vbNewLine & ex.Message, 16, "Status")
        End Try
    End Sub

    Public Function FW_Cut_stock_frith_in_out(ByVal WI As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal tag_readed As String, ByVal tag_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal scan_lot As String)
        Dim c_re_check As Integer = 0
        Return True
    End Function

    Public Function check_remain_in_detail_test(ByVal order_number As String, ByVal tag_seq As String, ByVal scan_qty As String)
        Try
            Dim count_data As String = Nothing
            Dim strCommand123 As String = "SELECT COUNT(id) as c,  id as i   FROM sup_scan_pick_detail  where com_flg = '0' and item_cd ='" & Module1.A_PAST_NO & "'  GROUP BY id"
            ' MsgBox("check_remain_in_detail_test == >" & strCommand123)
            Dim command123 As SqlCommand = New SqlCommand(strCommand123, myConn)
            Dim status As Integer = 0
            reader = command123.ExecuteReader() 'ติดบรรทัดนี้'
            Do While reader.Read
                If reader("c").ToString() = "0" Then
                    ID_table_detail = "NODATA"
                    status = 0
                Else
                    ID_table_detail = reader("i").ToString()
                    status = 1
                End If
            Loop
            reader.Close()

            If status = 1 Then
                Dim count_remain As String = Nothing
                Dim check_qe_status As Integer = Remain_check_qr_part(order_number, tag_seq)
                Dim strCommand1234 As String = "SELECT COUNT(id) as c   , scan_lot , tag_seq  FROM sup_scan_pick_detail  where scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and com_flg = '0' and scan_qty <> '" & scan_qty & "' GROUP BY scan_lot,tag_seq"
                Dim command1234 As SqlCommand = New SqlCommand(strCommand1234, myConn)
                reader = command1234.ExecuteReader()
                If check_qe_status = 4 Then
                    status = 0
                End If
                If reader.Read() Then
                    Do While reader.Read
                        count_remain = reader("c").ToString()
                        If count_remain <> Nothing Or count_remain = "0" Then
                            'MsgBox("count = " & reader("c").ToString())
                        Else
                            status = 1 'scan เจอ'
                        End If
                    Loop
                    reader.Close()
                Else
                    reader.Close()
                    If query_join_check(scan_qty) = 1 Then
                        status = 1
                    End If
                    If query_join_check(scan_qty) = 0 Then
                        status = 2
                    End If
                End If
            End If
            Return status
        Catch ex As Exception
            MsgBox("ERROR check_remain_in_detail_test DATA NULL" & vbNewLine & ex.Message, 16, "Status")
        End Try
        Return 0
    End Function
    Public Function query_join_check(ByVal scan_qty As String)
        Dim str_join As String = "SELECT COUNT (sp.id) AS c, cp.scan_qty AS tmp_qty, sp.scan_qty AS master_qty FROM sup_scan_pick_detail sp, check_qr_part cp WHERE sp.item_cd = cp.item_cd AND sp.scan_lot = cp.scan_lot AND sp.tag_seq = cp.tag_seq AND ( cp.scan_qty <> '" & scan_qty & "' OR sp.scan_qty <> '" & scan_qty & "' ) GROUP BY cp.scan_qty, sp.scan_qty "
        Dim command_join As SqlCommand = New SqlCommand(str_join, myConn)
        reader = command_join.ExecuteReader()
        Dim status = 0
        Do While reader.Read
            If reader("c").ToString() = "1" Then
                status = 1
            Else
                status = 0
            End If
        Loop
        reader.Close()
        Return status
    End Function
    Public Function Remain_check_qr_part(ByVal order_number As String, ByVal tag_seq As String)
        Try
            Dim status As Integer = 0
            Dim strCommand12345678 As String = "SELECT COUNT(id) as c_c   FROM check_qr_part  where scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "'  "
            '  MsgBox(strCommand1234)
            Dim strCommand1234567 As SqlCommand = New SqlCommand(strCommand12345678, myConn)
            reader = strCommand1234567.ExecuteReader()
            Do While reader.Read = True
                If reader("c_c").ToString() = "1" Then
                    status = 4
                Else
                    status = 3
                End If
            Loop
            reader.Close()
            Return status

        Catch ex As Exception
            MsgBox("ERROR Remain_check_qr_part")
        End Try
        Return 0
    End Function

    Public Function check_scan_detail_PO(ByVal scan_po As String, ByVal scan_code_suppier As String)
        Dim testLen As Integer = Len(scan_qty.Text)
        Dim num As Integer = 0
        Dim check_c As Integer = 0
        Dim status As Integer = 0
        Dim count As Integer = 0
        If testLen = 62 Then
            'MsgBox("FUNCTION check_scan_detail_PO")
            For Each key In Module1.arr_pick_detail_po
                count += 1
            Next
            If c_num <= (count - 1) Then
                ''''''''''''check ว่า ให้ยิงตามลำดับ''''''''''''''''''
                For Each key In Module1.arr_pick_detail_po
                    Dim check_code_suppier As String = Module1.arr_pick_detail_lot(check_c).ToString
                    check_code_suppier = check_code_suppier.Substring(0, 5)
                    Dim check_po As String = Module1.arr_pick_detail_po(check_c).ToString
                    If check_code_suppier = scan_code_suppier And check_po = scan_po Then
                        If check_c <> c_num Then 'g'
                            'MsgBox("!=")
                            status = 2
                            Return status
                        End If
                    Else
                        status = 0
                    End If
                    check_c = check_c + 1
                Next
                '''''''''''''''''''''''''''''''
                Dim code_suppier As String = Module1.arr_pick_detail_lot(c_num).ToString
                code_suppier = code_suppier.Substring(0, 5)
                Dim po As String = Module1.arr_pick_detail_po(c_num).ToString
                Dim QTY As String = Module1.arr_pick_detail_qty(c_num).ToString
                ' MsgBox(code_suppier & " = " & scan_code_suppier)
                ' MsgBox(po & " = " & scan_po)
                If code_suppier = scan_code_suppier And po = scan_po Then
                    ' MsgBox("==totall_qty_scan===" & totall_qty_scan)
                    'MsgBox("QTY fifo = " & QTY)
                    If totall_qty_scan >= QTY Then
                        If check_po_lot = "pick_ok" Then
                            ' MsgBox("PICK_OK ===>")
                            Dim qty_add As Double = 0.0
                            Dim data_remain As Double = 0.0
                            qty_add = CDbl(Val(QTY))
                            ' MsgBox("QTY PO" & QTY)
                            data_remain = totall_qty_scan - qty_add
                            If data_remain >= 0 Then 'การเก็บ remain'
                                'MsgBox(data_remain)
                                QTY_INSERT_LOT_PO = data_remain
                                arr_remain_qty.Add(data_remain)
                            End If
                            c_num += 1
                            totall_qty_scan = 0.0
                            check_po_lot = ""
                        End If
                    ElseIf totall_qty_scan < QTY Then
                    End If
                    status = 1 '1 = ถูก'
                End If
                num = num + 1
            ElseIf c_num > (count - 1) Then
                status = 3 'scan ครบแล้ว'
                '  MsgBox("สแกน ตาม fifo แล้ว")
            End If
        ElseIf testLen = 103 Then
            For Each key In Module1.arr_pick_detail_lot
                count += 1
            Next
            If c_num <= (count - 1) Then
                ''''''''''''check ว่า ให้ยิงตามลำดับ''''''''''''''''''
                Dim lot_scan As String = scan_qty.Text.Substring(58, 4)
                For Each key In Module1.arr_pick_detail_lot
                    'Dim check_lot As String = Module1.arr_pick_detail_lot(check_c).ToString
                    ' If lot_scan = check_lot Then
                    'If check_c <> c_num Then 'g'
                    'status = 2
                    ' Return status
                    ' End If
                    '  Else
                    'status = 0
                    'End If
                    check_c = check_c + 1
                Next
                status = 1 'ชั่วคราว'
                '''''''END'''''''''''''''''''''
                ' MsgBox("lot")
                ' For Each key In Module1.arr_pick_detail_lot
                Dim lot As String = Module1.arr_pick_detail_lot(c_num).ToString
                'Dim po As String = Module1.arr_pick_detail_po(num).ToString
                Dim QTY As String = Module1.arr_pick_detail_qty(c_num).ToString
                'MsgBox(code_suppier & " = " & scan_code_suppier)
                ' MsgBox(po & " = " & scan_po)
                QTY_INSERT_LOT_PO = 0
                If lot_scan = lot Then
                    ' MsgBox("==totall_qty_scan===" & totall_qty_scan)
                    If totall_qty_scan >= QTY Then
                        If check_po_lot = "pick_ok" Then
                            ' MsgBox("PICK_OK ===>")
                            Dim qty_add As Double = 0.0
                            Dim data_remain As Double = 0.0
                            qty_add = CDbl(Val(QTY))
                            data_remain = totall_qty_scan - qty_add

                            If data_remain >= 0 Then 'การเก็บ remain'
                                'MsgBox(data_remain)
                                QTY_INSERT_LOT_PO = data_remain
                                arr_remain_qty.Add(data_remain)
                            End If
                            c_num += 1
                            totall_qty_scan = 0.0
                            check_po_lot = ""
                        End If
                    ElseIf totall_qty_scan < QTY Then
                    End If
                    status = 1
                End If
                num = num + 1
            ElseIf c_num > (count - 1) Then
                status = 3
                '  MsgBox("สแกน ตาม fifo แล้ว")
            End If
            'Next
        End If
        If check_FA_TAG_FG() = False Then
            status = 4
        ElseIf check_FA_TAG_FG() = 1 Then
            status = 5
        End If
        'MsgBox("status = " & status)
        Return status
    End Function
    Public Sub set_show_supply()
        show_number_supply.Text = Module1.G_show_data_supply
        '  show_number_supply.Text = CDbl(Val(Module1.show_data_supply)) - CDbl(Val(Module1.show_data_remain))
        want_to_tag.Text = CDbl(Val(show_qty.Text.Substring(6))) - CDbl(Val(show_number_supply.Text))
    End Sub
    Public Sub set_show_remain()

        show_number_remain.Text = Module1.show_data_remain
    End Sub
    Private Sub PictureBox3_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox3.MouseDown

    End Sub

    Private Sub PictureBox3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub
    Public Sub next_image(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles text_box_success.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.F3
                Panel7.Visible = False
                alert_success.Visible = False
                'MsgBox("PRIVATE")
            Case System.Windows.Forms.Keys.F4

            Case System.Windows.Forms.Keys.Enter
                If check_process = "OK" Then
                    Try
                        status_alert_image = ""
                        'select_pick_add.load_stock()
                        Select_plan_fg.Show()
                    Catch ex As Exception
                        MsgBox("NO TRY")
                    End Try
                    Me.Close()
                End If
                If leng_scan_qty = 62 Then
                    If bool_check_scan = "scan_ok_pickdetail" Then
                        Panel7.Visible = False
                        alert_pickdetail_ok.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "pick_detail_number" Then
                        Panel7.Visible = False
                        alert_pickdetail_number.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If

                        If show_number_supply.Text = "0" Or show_number_supply.Text = 0 Then
                            text_tmp.Text = 0
                        End If

                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "HAVE_TAG_REMAIN" Then
                        Panel7.Visible = False
                        alert_tag_remain.Visible = False
                        text_tmp.Text = ""
                        Re_scan_default()
                    ElseIf bool_check_scan = "HAVE_Reprint" Then
                        Panel7.Visible = False
                        alert_reprint.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "Plase_scna_detail" Then
                        Panel7.Visible = False
                        alert_detail.Visible = False
                        Re_scan_default()
                    ElseIf status_alert_image = "Part_incorrect" Then
                        Panel7.Visible = False
                        alert_pa.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf status_alert_image = "success" Then
                        Panel7.Visible = False
                        alert_success.Visible = False
                    ElseIf status_alert_image = "success_remain" Then
                        Panel7.Visible = False
                        alert_success_remain.Visible = False
                    ElseIf status_alert_image = "loop" Then
                        Panel7.Visible = False
                        alert_loop.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf status_alert_image = "loop_re_scan" Then
                        Panel7.Visible = False
                        alert_loop.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    End If

                ElseIf leng_scan_qty = 76 Then
                    If status_alert_image = "alert_right_fa" Then
                        Panel7.Visible = False
                        alert_right_fa.Visible = False
                        'MsgBox("Please scan FA tag on the top right")
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    End If

                ElseIf leng_scan_qty = 103 Then
                    If bool_check_scan = "scan_ok_pickdetail" Then
                        Panel7.Visible = False
                        alert_pickdetail_ok.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "pick_detail_number" Then
                        Panel7.Visible = False
                        alert_pickdetail_number.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "HAVE_TAG_REMAIN" Then
                        Panel7.Visible = False
                        alert_tag_remain.Visible = False
                        Re_scan_default()
                    ElseIf bool_check_scan = "HAVE_Reprint" Then
                        alert_reprint.Visible = False
                        Panel7.Visible = False
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
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "Plase_scna_detail" Then
                        Panel7.Visible = False
                        alert_detail.Visible = False
                        Dim Len_length As Integer = Len(scan_qty.Text)
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
                        scan_qty.Text = ""
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Focus()
                    ElseIf status_alert_image = "Part_incorrect" Then
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
                        'text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        Panel7.Visible = False
                        alert_pa.Visible = False
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf status_alert_image = "success" Then
                        Panel7.Visible = False
                        alert_success.Visible = False
                    ElseIf status_alert_image = "success_remain" Then
                        Panel7.Visible = False
                        alert_success_remain.Visible = False
                    ElseIf status_alert_image = "loop" Then
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
                        'text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        Panel7.Visible = False
                        alert_loop.Visible = False
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf status_alert_image = "loop_re_scan_fa" Then
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
                        'text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        Panel7.Visible = False
                        alert_loop.Visible = False
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    ElseIf bool_check_scan = "NO_data_tranfer" Then
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
                        'text_tmp.Text = Module1.SCAN_QTY_TOTAL
                        Panel7.Visible = False
                        alert_no_tranfer_data.Visible = False
                        If text_tmp.Text = "0" Then
                            text_tmp.Text = 0
                        Else
                            text_tmp.Text = scan_qty_total
                        End If
                        scan_qty.Text = ""
                        scan_qty.Focus()
                    End If
                End If
        End Select
    End Sub

    Private Sub PictureBox3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub alert_detail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub alert_tag_remain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label5_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'show_pln.Visible = False
        'scan_qty.Focus()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub alert_pa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub alert_loop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub scan_qty_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles scan_qty.TextChanged

    End Sub

    Private Sub PictureBox3_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click

    End Sub

    Private Sub end_pa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel4.Visible = False
    End Sub

    Private Sub PictureBox4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Panel4.Visible = False
        scan_qty.Focus()
    End Sub

    Private Sub user_id_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles user_id.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                password.Focus()
        End Select
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        Dim strCommand12345678 As String = "select count(su.su_id) as c_id  , su.sug_id as per from sys_users su  , sys_user_groups sug  where su.emp_id = '" & user_id.Text & "' and su.sys_pass = '" & password.Text & "' and su.sug_id  = sug.sug_id and su.enable = '1' GROUP BY su.sug_id"
        'MsgBox(strCommand12345678)
        Dim cmd As SqlCommand = New SqlCommand(strCommand12345678, myConn)
        reader = cmd.ExecuteReader()

        If reader.Read() Then
            If reader("c_id").ToString() = "1" And reader("per").ToString() <> "3" Then
                Module1.user_reprint = user_id.Text()
                Label11.Text = "QTY BEFORE : 0"
                Label10.Text = "QTY AFTER : 0"
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

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
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
            Dim check = check_reprint_stock("62", old_qty, TextBox1.Text, re_qty_number)
            If check = "SUCCESS" Then
loop_check_open_bt:
                If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
                    Dim PO As String = TextBox1.Text.Substring(2, 10)
                    Dim seq_text As String = TextBox1.Text.Substring(59, 3)
                    Dim textbox_split = TextBox1.Text.Split(" ")
                    Dim item_cd As String = textbox_split(0)
                    Dim data_item_cd As String = item_cd.Substring(12)
                    insert_log(old_qty, "1", PO, seq_text, data_item_cd)
                    Bluetooth_Reprint(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, Model_detail, re_qty_number, loc_detail, user_detail, now_date_detail, now_time_detail, new_qr_re_print, SEQ)
                Else
                    'MsgBox("connect faill")
                    MsgBox("กรุณาเปิดเครื่องปริ้น")
                    GoTo loop_check_open_bt
                End If
            ElseIf check = "FAILL" Then
                MsgBox("QTY ใน stock ไม่เพียงพอต่อการ reprint ")
            ElseIf check = "NO_DATA" Then
                MsgBox("ไม่พบข้อมูล QR นี้")
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
            stInfoSet1.addr = main.number_printter_bt
            M_reprint = "WEB_POST"
            Dim pinlen As UInt32 = CType(pin.Length, UInt32)
            M_reprint = "FW"
            Dim old_qty As String = TextBox1.Text.Substring(52, 6)
            Dim check = check_reprint_stock("103", Trim(old_qty), TextBox1.Text, re_qty_number)
            'MsgBox(check)
            If check = "SUCCESS" Then
loop_open_printer:
                If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
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
                    Bluetooth_Reprint(stInfoSet, pin, pinlen1, part_no_detail, part_name_detail, Model_detail, re_qty_number, loc_detail, user_detail, now_date_detail, now_time_detail, new_qr_re_print, SEQ)
                Else
                    'MsgBox("connect faill")
                    MsgBox("กรุณาเปิดเครื่องปริ้น")
                    GoTo loop_open_printer
                End If
            ElseIf check = "FAILL" Then
                MsgBox("QTY ใน stock ไม่เพียงพอต่อการ reprint ")
            ElseIf check = "NO_DATA" Then
                MsgBox("ไม่พบข้อมูล QR นี้")
            End If
        Else
            MsgBox("QR CODE ไม่ถูกต้อง")
        End If

    End Sub

    Private Sub ListView2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ListView2_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView2.SelectedIndexChanged

    End Sub
    Public Sub get_data_tetail()
        scan_qty.Visible = False
        Panel6.Visible = True
        Dim x As ListViewItem = New ListViewItem()
        ListView2.Items.Clear()
        Try
            Dim check As Integer = 0
            Dim num As Integer = 0
            For Each key In Module1.arr_pick_detail_lot
                Dim lot As String = Module1.arr_pick_detail_lot(num).ToString
                Dim QTY As String = Module1.arr_pick_detail_qty(num).ToString
                x = New ListViewItem(" - ")
                x.SubItems.Add(Module1.arr_pick_detail_lot(num).ToString)
                x.SubItems.Add(Module1.arr_pick_detail_qty(num).ToString)
                x.SubItems.Add("0")
                ListView2.Items.Add(x)
                If Module1.check_pick_detail <> 0 Then
                    Dim count_scan As Integer = 0
                    For Each key2 In Module1.arr_check_lot_scan
                        If lot = Module1.arr_check_lot_scan(count_scan).ToString Then
                            Dim i As Integer = 0
                            Dim total_qty As Integer = 0
                            For Each key3 In Module1.arr_check_QTY_scan
                                If lot = Module1.arr_check_lot_scan(i).ToString Then
                                    ' MsgBox(Module1.arr_check_QTY_scan(i))
                                    total_qty = CDbl(Val(total_qty)) + CDbl(Val(Module1.arr_check_QTY_scan(i)))
                                    If CDbl(Val(QTY)) <= CDbl(Val(total_qty)) Then
                                        ListView2.Items(num).BackColor = Color.FromArgb(103, 255, 103)
                                        'Dim val = Module1.M_CHECK_LOT_COUNT_FW(num)
                                    ElseIf CDbl(Val(QTY)) > CDbl(Val(total_qty)) Then
                                        ListView2.Items(num).BackColor = Color.Yellow
                                        ' Dim val = Module1.M_CHECK_LOT_COUNT_FW(num)
                                    End If
                                End If
                                If total_qty > QTY Then
                                    total_qty = QTY
                                End If
                                ListView2.Items(num).SubItems(3).Text = total_qty
exit_loop:
                                i = i + 1
                            Next
                        End If
Exit_count2:
                        count_scan = count_scan + 1
                    Next
                End If
                num = num + 1
            Next
        Catch ex As Exception
            MsgBox("ERROR LOT FAILL FROM CODE SUPPIER " & vbNewLine & ex.Message, 16, "Status in")
        End Try

    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        scan_qty.Visible = True
        Panel6.Visible = False
        scan_qty.Focus()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Label11.Text = "QTY BEFORE : 0"
        TextBox1.Text = ""
        TextBox2.Text = ""
        Label10.Text = "QTY AFTER : 0"
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_Text_key_down(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                ' MsgBox("001")
                If Len(TextBox1.Text) = "62" Then
                    ' MsgBox("1")
                    Dim n_old As Double = 0.0
                    n_old = CDbl(Val(TextBox1.Text.Substring(51, 8)))
                    Dim qty_old As String = n_old
                    'check_sock("62", qty_old)
                    Label11.Text = "QTY BEFORE :" & qty_old
                ElseIf Len(TextBox1.Text) = "103" Then
                    ' MsgBox("2")
                    Dim qty_old As String = TextBox1.Text.Substring(52, 6)
                    'check_sock("103", qty_old)
                    Label11.Text = "QTY BEFORE :" & Trim(qty_old)
                Else
                    Label11.Text = "QTY BEFORE : 0"
                End If
                TextBox2.Focus()
        End Select
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click, PictureBox6.Click
        user_id.Text = ""
        password.Text = ""
        scan_qty.Focus()
        Panel5.Visible = False
        Panel4.Visible = False
    End Sub

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

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Panel5_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel5.GotFocus

    End Sub

    Public Sub insert_pick_log(ByVal REMAIN_ID As String, ByVal F_wi As String, ByVal used_qty As Double, ByVal create_date As String, ByVal create_by As String, ByVal updated_date As String, ByVal updated_by As String, ByVal scan_lot As String, ByVal updated_seq As String, ByVal F_item_cd As String)
        Try
            Dim strCommand123 As String = "SELECT   id as i   FROM sup_scan_pick_detail  where scan_lot = '" & scan_lot & "' and tag_seq = '" & updated_seq & "'  and item_cd ='" & F_item_cd & "'"
            Dim command123 As SqlCommand = New SqlCommand(strCommand123, myConn)
            reader = command123.ExecuteReader() 'ติดบรรทัดนี้'
            Dim id As Integer = 0
            REMAIN_ID = 0
            If reader.Read Then
                REMAIN_ID = reader("i").ToString()

            Else
                MsgBox("QUERY sup_scan_pick_detail NO ID")
            End If
            reader.Close()
            Dim str_insert_log = "INSERT INTO sup_pick_log (REF_ID , WI_NO , USED_QTY , CREATED_DATE , CREATED_BY , UPDATED_DATE , UPDATED_BY , ADD_FLG) VALUES ('" & REMAIN_ID & "','" & F_wi & "' , '" & used_qty & "','" & create_date & "','" & create_by & "' , '" & updated_date & "','" & updated_by & "' , '2')"
            Dim command2 As SqlCommand = New SqlCommand(str_insert_log, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
            Dim get_id_log = "select * from sup_pick_log where REF_ID = '" & REMAIN_ID & "' and WI_NO = '" & F_wi & "' and USED_QTY = '" & used_qty & "'  and CREATED_DATE = '" & create_date & "' and CREATED_BY = '" & create_by & "' and UPDATED_DATE = '" & updated_date & "' and updated_by = '" & updated_by & "'"
            Dim cmd_get As SqlCommand = New SqlCommand(get_id_log, myConn)
            reader = cmd_get.ExecuteReader() 'ติดบรรทัดนี้'

            If reader.Read Then
                id_pick_log_supply = reader("ID").ToString()
                arr_pick_log.Add(id_pick_log_supply)
            Else
                MsgBox("DATA SET FAIL INSERT_PICK_LOG")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox("insert_pick_log_erro" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub
    Public Sub cut_stock_FASYSTEM(ByVal used_qty As String, ByVal F_item_cd As String, ByVal updated_seq As String, ByVal scan_lot As String, ByVal tag_read As String)
        Try
            Dim data_key_up = "NO_DATA"
            If tag_read.Substring(73, 2) = "E2" Then
                Dim arr_qty = tag_read.Split(" ")
                If tag_read.Substring(30, 2) = "E2" And tag_read.Substring(73, 2) = "E2" Then
                    data_key_up = arr_qty(28)
                ElseIf tag_read.Substring(73, 2) = "E2" Then
                    data_key_up = arr_qty(30)
                Else
                    data_key_up = arr_qty(32)
                End If
            ElseIf tag_read.Substring(73, 2) = "E " Then
                Dim arr_qty = tag_read.Split(" ")
                data_key_up = arr_qty(31)
            Else
                'MsgBox("READY")
                Dim arr_qty = tag_read.Split(" ")
                'MsgBox("READY2")
                'MsgBox(tag_read.Substring(56, 1))
                'MsgBox("end")
                If tag_read.Substring(56, 1) <> " " Then
                    'MsgBox("INto")
                    Try
                        If F_item_cd.Substring(11, 1) = "G" Then
                            data_key_up = arr_qty(31)
                        Else
                            data_key_up = arr_qty(32)
                            'MsgBox(arr_qty(32))
                            'MsgBox("INto2")
                        End If
                    Catch ex As Exception
                        data_key_up = arr_qty(32)
                    End Try

                Else
                    'MsgBox("OUT")
                    data_key_up = arr_qty(33)
                    'MsgBox(arr_qty(33))
                    'MsgBox("OUT2")
                End If
            End If
            'MsgBox("0")
            Dim SEQ = "NO_DATA"
            If tag_read.Substring(94, 1) <> " " Then
                SEQ = updated_seq.Substring(8, 3)
            End If
            'MsgBox("1")
            If tag_read.Substring(94, 1) = " " Then
                'MsgBox("2")
                data_key_up = tag_read.Substring(95, 8)
                'MsgBox("3")
                SEQ = tag_read.Substring(95, 3)
                'MsgBox("4")
                'MsgBox("---->>>>>" & tag_read.Substring(95, 1))
                'MsgBox("5")
                If tag_read.Substring(95, 1) = " " Then
                    'MsgBox("--->")
                    data_key_up = data_key_up.Substring(3)
                    ' MsgBox("-----<")
                    SEQ = " "
                End If
            End If
            'MsgBox("7")
            'MsgBox(" item_cd = " & F_item_cd)
            'MsgBox("SEQ  = " & SEQ)
            'MsgBox("LOT  = " & scan_lot)
            'MsgBox("data_key_up = " & data_key_up)
            'MsgBox(F_item_cd)
            Try
                If F_item_cd.Substring(11, 1) = "G" Then
                    F_item_cd = F_item_cd.Substring(0, 11)
                End If
            Catch ex As Exception

            End Try
            
            Dim get_id_log = "select * from FA_TAG_FG where ITEM_CD = '" & F_item_cd & "' AND TAG_SEQ = '" & SEQ & "'AND LOT_NO = '" & scan_lot & "' and KEY_UP = '" & data_key_up & "' and LINE_CD = '" & tag_read.Substring(2, 6) & "'"
            'MsgBox(get_id_log)
            Dim cmd_get As SqlCommand = New SqlCommand(get_id_log, myConn_fa)
            reader = cmd_get.ExecuteReader()
            Dim update_qty As Double = 0.0
            If reader.Read Then
                update_qty = CDbl(Val(reader("TAG_QTY").ToString())) - CDbl(Val(used_qty))
            End If
            reader.Close()
            'MsgBox("update_qty = " & update_qty)
            Dim FLG_STATUS As String = "0"
            If update_qty <= 0 Then
                update_qty = 0
                FLG_STATUS = "1"
            Else
                FLG_STATUS = "2"
            End If
            Dim str_update_qty = "EXEC [dbo].[cut_stock_pick_fg] @qty = '" & used_qty & "'  , @flg_status = '" & FLG_STATUS & "' , @item_cd = '" & F_item_cd & "' , @seq = '" & SEQ & "' , @lot = '" & scan_lot & "' , @KEY_UP = '" & data_key_up & "' , @LINE_CD = '" & tag_read.Substring(2, 6) & "'"
            'MsgBox(str_update_qty)
            Dim cmd_update As SqlCommand = New SqlCommand(str_update_qty, myConn_fa)
            reader = cmd_update.ExecuteReader()
            reader.Close()
            'MsgBox("CUT STOCK OK")
        Catch ex As Exception
            MsgBox("FAILL cut_stock_FASYSTEM" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub

    Public Function check_FA_TAG_FG()
        'MsgBox("111")
        Try
            Dim data_key_up As String = ""
            Dim KEY_UP = scan_qty.Text.Substring(58, 4)
            Dim SEQ As String = ""
            'MsgBox("READY ")
            Dim arr_qty = scan_qty.Text.Split(" ")

            If scan_qty.Text.Substring(19, 14) = "49373-25591-EP" Then
                GoTo query
            End If

            If scan_qty.Text.Substring(73, 2) = "E2" Then
                arr_qty = scan_qty.Text.Split(" ")

                If scan_qty.Text.Substring(30, 2) = "E2" And scan_qty.Text.Substring(73, 2) = "E2" Then
                    data_key_up = arr_qty(28)
                ElseIf scan_qty.Text.Substring(73, 2) = "E2" Then
                    data_key_up = arr_qty(30)
                Else
                    data_key_up = arr_qty(32)
                End If
                'MsgBox("READY2 ")
            ElseIf scan_qty.Text.Substring(73, 2) = "E " Then
                arr_qty = scan_qty.Text.Split(" ")
                data_key_up = arr_qty(31)
            Else
                'MsgBox("READY3")
                arr_qty = scan_qty.Text.Split(" ")
                ' MsgBox("---->")
                If scan_qty.Text.Substring(73, 1) = "G" Then
                    data_key_up = arr_qty(31)
                    'MsgBox("311111")
                ElseIf scan_qty.Text.Substring(56, 1) <> " " Then
                    data_key_up = arr_qty(32)
                    ' MsgBox("set2")
                Else
                    data_key_up = arr_qty(33)
                End If
            End If
            'MsgBox("READY4")
            ' MsgBox("01")
            Dim ITEM_CD = ""
            Dim lot_sep = ""
            'MsgBox("00")
            If scan_qty.Text.Substring(94, 1) <> " " Then
                Dim arr_item_cd = scan_qty.Text.Split(" ")
                'MsgBox("01")
                ITEM_CD = arr_item_cd(0).Substring(19)
                'MsgBox("02")
                lot_sep = scan_qty.Text.Substring(58, 4)
            Else
                Dim arr_item_cd = scan_qty.Text.Split(" ")
                ' MsgBox("0000")
                ITEM_CD = arr_item_cd(11)
                '  MsgBox("11111")
                lot_sep = scan_qty.Text.Substring(58, 4)
            End If
            'MsgBox("03")
            If scan_qty.Text.Substring(94, 1) <> " " Then
                'MsgBox("SEQ")
                SEQ = data_key_up.Substring(8, 3)
                'MsgBox("SEQ2")
            End If
            'MsgBox("---0")
            If scan_qty.Text.Substring(94, 1) = " " Then 'ตรวจสอบ TAG ว่าเป็น TAG REPRINT หรือไม่'
                ' MsgBox("SEQ3")
                If scan_qty.Text.Substring(95, 1) <> " " Then 'reprint ถูก'
                    'MsgBox("SEQ4")
                    data_key_up = scan_qty.Text.Substring(95, 8)
                    'MsgBox("SEQ5")
                    SEQ = scan_qty.Text.Substring(95, 3)
                    'MsgBox("SEQ6")
                End If
                Dim check_arr = data_key_up.Split(" ")
                'MsgBox("SEQ7")
                Dim i As Integer = 0
                If scan_qty.Text.Substring(95, 1) = " " Then 'TAG REPRINT ผิด'
                    '     MsgBox("SEQ8")
                    '      MsgBox(scan_qty.Text.Substring(98, 5))
                    data_key_up = scan_qty.Text.Substring(98, 5) 'ผิดตรงนี้ TAG REPRINT '
                    SEQ = " "
                End If
                '   MsgBox("data_key_up = " & data_key_up)
            End If
            'MsgBox("---01")
            'MsgBox(ITEM_CD)
            'MsgBox(ITEM_CD.Substring(11, 1))
            'MsgBox(ITEM_CD.Substr-ing(0, 11))
query:
            If scan_qty.Text.Substring(19, 14) = "49373-25591-EP" Then
                data_key_up = arr_qty(40)
                lot_sep = arr_qty(15).Substring(2, 4)
                ITEM_CD = scan_qty.Text.Substring(19, 14)
                SEQ = data_key_up.Substring(8, 3)
            End If
            Dim sql As String = "select IND , FLG_STATUS from FA_TAG_FG where LOT_NO = '" & lot_sep & "' and ITEM_CD = '" & ITEM_CD & "' and TAG_SEQ = '" & SEQ & "' and LINE_CD = '" & scan_qty.Text.Substring(2, 6) & "' and KEY_UP = '" & data_key_up & "'"
            'MsgBox(sql)
            Dim command2 As SqlCommand = New SqlCommand(sql, myConn_fa)
            reader = command2.ExecuteReader()
            Dim count As String = "NO_DATA"
            Dim flg_status As String = "NO_DATA"
            Do While reader.Read = True
                count = reader("IND").ToString()
                flg_status = reader("FLG_STATUS").ToString()
            Loop
            reader.Close()
            'MsgBox(count)
            'MsgBox(flg_status)
            If count <> "NO_DATA" Then
                If flg_status <> "2" Then
                    'bool_check_scan = "ever"
                    'Module1.check_count = 1
                    '       MsgBox("1")
                    Return 1
                Else
                    '        MsgBox("2")
                    Return 2
                End If
            Else
                '     MsgBox("FALSE")
                Return False
            End If
        Catch ex As Exception
            MsgBox("FALL please check")
        End Try
        Return 0
    End Function

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
    Private Sub TextBox2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Label10.Text = "QTY AFTER : " & TextBox2.Text
    End Sub

    Private Sub check_qr_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles check_qr.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                check_qr.Visible = False
                Panel7.Visible = False
                alert_no_tranfer_data.Visible = False
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
                If text_tmp.Text = "0" Then
                    text_tmp.Text = 0
                Else
                    text_tmp.Text = scan_qty_total
                End If
                scan_qty.Text = ""
                scan_qty.Focus()
        End Select
            End Sub

    Private Sub alert_pickdetail_number_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles alert_pickdetail_number.Click

    End Sub
End Class