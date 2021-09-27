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
    'PHASE 8 TRUE'
    ' Inherits Form
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
    'Public PD_ADD_PART As select_pick_add
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
    Public count_box_check_qr_part As Integer = 0
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
    Dim status_check_washing As Integer = 0
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
    Public count_net As Integer = 0
    Private Sub part_detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
re_connect:
            If Api.check_net() = True Then
                Dim connect_db = New connect()
                myConn = connect_db.conn()
                myConn_fa = connect_db.conn_fa()
                Timer1.Enabled = True
            Else
                Timer1.Enabled = False
                'MsgBox("อินเตอร์เน็ตไม่เสถียร กรุณา กด ENT เพื่อ รอ INTERNET")
                GoTo re_connect
            End If
        Catch
            MsgBox("connect poen please check open")
            If reader.Read() = True Then
                reader.Close()
                GoTo re_connect
            End If
        Finally
            'Panel4.Visible = False
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
            text_box_success.Visible = False
            alert_success.Visible = False
            alert_success_remain.Visible = False
            alert_loop.Visible = False
            alert_right_fa.Visible = False
            alert_reprint.Visible = False
            alert_open_printer.Visible = False
            alert_no_tranfer_data.Visible = False
            alert_14_day.Visible = False
            Panel7.Visible = False
            'ชั่วคราว'
            ' Panel6.Visible = False
            btn_detail_part.Visible = True
            alert_14_day.Visible = False
            set_qty_check_qr_part()
            'get_data_tetail() 'ปิดโชว์ FIFO
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = False
        set_default_data()
        QTY_INSERT_LOT_PO = 0.0
        Module1.G_show_data_supply = 0.0
        'delete_data_check_qr_part()
        Module1.arr_check_lot_scan = New ArrayList()
        Module1.arr_check_PO_scan = New ArrayList()
        Module1.arr_check_QTY_scan = New ArrayList()
        Module1.show_data_supply = 0.0
        Module1.show_data_remain = 0.0
        Module1.check_pick_detail = 0
        Module1.total_qty = 0
        Module1.total_database = 0
        delete_data_check_qr_part() 'ยังไม่ได้สร้าง function'
        Select_plan_fg.Show()
        Me.Close()
    End Sub
    Public Sub check_qr_part()
        Dim ps = Part_No.Text.Substring(16)
        Dim scan_p = text_tmp.Text.Substring(19)
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
                Dim testLen As Integer = Len(testString)
                length = testLen
                leng_scan_qty = length
                Dim req_qty As Double = 0.0
                req_qty = Val(Module1.FG_QTY.Substring(6))
                Dim number_remain As Double = 0.0
                number_remain = CDbl(Val(show_number_remain.Text))
                Dim ps = Part_No.Text.Substring(16)
                Dim QTY_show As Integer = show_qty.Text.Substring(6)
                If comp_flg = "0" Then
                    If testLen = 62 Then
                        MsgBox("NO SUPPROT WEBPOST")
                    ElseIf testLen = 76 Then
                        status_alert_image = "alert_right_fa"
                        Panel7.Visible = True
                        alert_right_fa.Visible = True
                        text_box_success.Focus()
                    ElseIf testLen = 103 Then
                        Try
                            Dim SearchForThis As String = " "
                            Dim FirstCharacter As Integer = testString.IndexOf(SearchForThis)
                            item_cd_scan = scan_qty.Text.Substring(19, FirstCharacter - 19)

                            fa_line_cd = scan_qty.Text.Substring(2, 6)
                            fa_act_date = scan_qty.Text.Substring(8, 8)
                            fa_lot_seq = scan_qty.Text.Substring(16, 3)
                            fa_qty = scan_qty.Text.Substring(52, 6)
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
                            'เคสเหลือจาก Tag
                            If Ck_dup(ListBox, order_number & supp_seq) = True Then
                                Dim check_w = check_washing()
                                If check_w = 1 Then
                                    bool_check_scan = "Over_14_days"
go_Over_14_days:
                                    Panel7.Visible = True
                                    alert_14_day.Visible = True
                                    check_qr.Visible = True
                                    check_qr.Focus()
                                    GoTo exit_keydown
                                ElseIf check_w = 2 Then
                                    bool_check_scan = "No_data_tranfer"
                                    GoTo go_No_data_tranfer
                                Else
                                    text_tmp.Text = fa_qty
                                    ' scan_qty.Text = ""
                                    ' scan_qty.Focus()
                                End If
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
                                Dim result = check_washing()
                                If result = 1 Then
                                    bool_check_scan = "Over_14_days"
                                    GoTo go_Over_14_days
                                ElseIf result = 0 Then
                                    If show_number_supply.Text > req_qty And firstscan = "0" And number_remain > 0 Then
                                        ' MsgBox("คุณสแกนครบแล้ว และมีเศษในกล่องชิ้นงาน", 16, "Alert")
                                        text_tmp.Text = fa_qty
                                        remain_qty1 = fa_tag_qty - req_qty
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
                                        fa_seq = fa_seq + 1
                                        sup_list.Add(fa_shift_seq & fa_lot_no & fa_tag_seq)
                                        Dim num As Integer = fa_seq
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
                                Else
                                    'MsgBox("Part incorrect")
                                    status_alert_image = "Part_incorrect"
                                    Panel7.Visible = True
                                    alert_pa.Visible = True
                                    text_box_success.Focus()
                                End If
                            End If
                        End If

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
                If check_arr(i) <> "" Then
                    tag_seq = check_arr(i)
                    GoTo out
                End If
                i = i + 1
            Next
out:
            order_number = scan_qty.Text.Substring(58, 4)
            qty_scan = scan_qty.Text.Substring(52, 6)
            seq_check_reprint = plan_seq
        ElseIf Len_length = 62 Then 'web post'
            MsgBox("NO SUPPROT WEBPOST")
        End If
        Dim strCommand3 As String = "SELECT top 1  COUNT(id) as c, com_flg  as com_flg , id as i  , scan_qty as qty FROM sup_scan_pick_detail  where item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' and scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and line_cd = '" & scan_qty.Text.Substring(2, 6) & "' and scan_qty >= '" & qty_scan & "' group by com_flg , id , scan_qty"
        Dim command3 As SqlCommand = New SqlCommand(strCommand3, myConn)
        reader = command3.ExecuteReader()
        Do While reader.Read = True
            count = reader("c").ToString()
            check_com_flg = reader("com_flg").ToString()
            id = reader("i").ToString()
            qty = reader("qty").ToString()
        Loop
        reader.Close()
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
                    bool_check_scan = "no_ever"
                    Module1.check_count = 0
                Else
                    bool_check_scan = "ever"
                    Module1.check_count = 1
                End If
                Return 0
            End If
            If count = 0 Then
                If check_qr_part_in_table() = True Then 'True หมายถึง ไม่มีข้อมูล ใน ตาราง check qr  '
                    bool_check_scan = "no_ever"
                    Module1.check_count = 0
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
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button2.Visible = False
        hidden_text_qr_code()
        Module1.M_QTY_STOP_SCAN = text_tmp.Text
        Dim total_qty = text_tmp.Text - Module1.check_QTY
        Button2.Visible = False
        btn_detail_part.Visible = False
        Button1.Visible = False
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
        Dim itemNStrqr As String = part_name_detail
        Dim strNCount As Integer = Len(part_name_detail)
        Dim numNCountTemp As Integer = 25 - strNCount
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
        If Bluetooth.btBluetoothOpen = True Then
            Bluetooth.btBluetoothClose()
        End If
        If Bluetooth_Connect_MB200i(stInfoSet, pin, pinlen) = True Then
            Dim stInfoSet1 As New LibDef.BT_BLUETOOTH_TARGET()   '  Bluetooth device information
            stInfoSet1.addr = main.number_printter_bt
            Dim pin1 As StringBuilder = New StringBuilder("0000")
            Dim pinlen1 As UInt32 = CType(pin1.Length, UInt32)
            Dim c As Integer = 0
            Dim text_temp_del_remain As Double = 0.0
            text_temp_del_remain = CDbl(Val(text_tmp.Text))
            Dim qrdetailSupply As String = "SUP-FG " & Module1.SLIP_CD & " " & Module1.FG_CUS_ORDER_ID & " " & itemStrqr & " " & Module1.G_show_data_supply & " " & Module1.A_USER_ID & " " & id_pick_log_supply & " " & date_sup & " " & time_sup
            Dim qr_detail_remain As String = "nodata"
            Dim index As Integer = 0
            Bluetooth_Print_MB200i(stInfoSet, pin, pinlen1, Module1.FG_PART_CD.Substring(16), Module1.FG_PART_NAME.Substring(12), Module1.FG_CUS_ORDER_ID, qty_detail, Module1.FG_LINE, Module1.A_USER_ID, now_date_detail, now_time_detail, qrdetailSupply)
            Timer1.Enabled = False
            Dim num As Integer = 0
            num += 1
        Else 'check ถ้าไม่เปิดเครื่องปริ้นให้แสดง POP UP'
            Panel7.Visible = True
            alert_open_printer.Visible = True
            GoTo loop_check_open_printer
        End If
        Try
            Dim strCommand As String = ""
            Dim str_plus As String = "select top 1 * from EXP_ORDER_FG where ORDER_FG_ID = '" & Module1.FG_ORDER_ID & "' "
            Dim cmd_plus As SqlCommand = New SqlCommand(str_plus, myConn_fa)
            reader = cmd_plus.ExecuteReader()
            Dim total_pig_qty As Double = 0.0
            Dim remain_qty As Double = 0.0
            Do While reader.Read()
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
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn_fa)
            reader = command.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("Can not update into database")
        End Try
        Try
            Dim com_flg As Integer = 0
            If total_qty = 0 Then
                com_flg = 1
            End If
            delete_data_check_qr_part()
        Catch ex As Exception

        End Try
        scan_qty_arrlist.Clear()
        scan_lot_arrlist.Clear()
        scan_read_arrlist.Clear()
        scan_seq_arrlist.Clear()
        scan_location.text_box_location.Text = ""
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

    Private Function Bluetooth_Connect_MB200i(ByVal stInfoSet As LibDef.BT_BLUETOOTH_TARGET, ByVal pin As StringBuilder, ByVal pinlen As UInt32) As [Boolean]
        Dim bRet As [Boolean] = False
        Dim ret As Int32 = 0
        Dim disp As [String] = ""
        Try
            ret = Bluetooth.btBluetoothOpen()
            If ret <> LibDef.BT_OK Then
            End If
            ret = Bluetooth.btBluetoothPairing(stInfoSet, pinlen, pin)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothPairing error ret[" & ret & "]"
                Return bRet
            End If
            ret = Bluetooth.btBluetoothSPPConnect(stInfoSet, 30000)
            If ret <> LibDef.BT_OK Then
                disp = "btBluetoothSPPConnect error ret[" & ret & "]"
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
            result_qty = Module1.G_show_data_supply
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "Pick : " & show_number_supply.Text & " pcs." & "  " & Module1.FG_MODEL)
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
            Dim del = Module1.delivery_date.Split()
            bBufWork = System.Text.Encoding.GetEncoding(932).GetBytes("K9B" & "DEL : " & del(0) & " | " & del(1))
            bBufWork.CopyTo(bBuf, len)
            len = len + bBufWork.Length
            '---------------------------------------------------------------------------------------------------------'
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
                bESC.CopyTo(bBuf, len)
                len = len + bESC.Length
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
            Dim printflg As [Boolean] = False
            While True
                Dim recvFlg As [Boolean] = False
                For i As Int32 = 0 To 9
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
                    bBuf = New [Byte]() {ENQ}
                    If SppSend(bBuf, ssizeGet) = False Then
                        GoTo L_END1
                    End If
                ElseIf bBufGet(0) = NAK Then
                    GoTo L_END1
                ElseIf bBufGet(0) = STX Then
                    bBufGet = New [Byte](4094) {}
                    If SppRecv(bBufGet, rsizeGet) = False Then
                        GoTo L_END1
                    End If
                    If bBufGet(9) <> ETX Then
                        GoTo L_END1
                    End If
                    If bBufGet(2) = &H47 OrElse bBufGet(2) = &H48 OrElse bBufGet(2) = &H53 OrElse bBufGet(2) = &H54 Then
                        Thread.Sleep(200)
                        bBuf = New [Byte]() {ENQ}
                        If SppSend(bBuf, ssizeGet) = False Then
                            GoTo L_END1
                        End If
                        Continue While
                    ElseIf (bBufGet(2) <> &H0) AndAlso (bBufGet(2) <> &H1) AndAlso (bBufGet(2) <> &H41) AndAlso (bBufGet(2) <> &H42) AndAlso (bBufGet(2) <> &H4E) AndAlso (bBufGet(2) <> &H4D) Then
                        Exit While
                    End If
                    printflg = True
                    Exit While
                End If
            End While
            If printflg = True Then
                disp = "Printing Successfully."
                MessageBox.Show(disp, "Printing complete")
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
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            Do While reader.Read = True
                count = reader("c").ToString()
                check_com_flg = reader("com_flg").ToString()
                id = reader("i").ToString()
                qty = reader("qty").ToString()
            Loop
            reader.Close()
            If check_com_flg = "0" Then
                Module1.check_count = 0
                inset_check_qr_part()
                Return 0
            End If

            If count = 0 Then
                Module1.check_count2 = 0
                boll = True
            Else
                text_tmp.Text = scan_qty_total
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
            Dim best_qty = show_qty.Text.Substring(6)
            Dim qty_s As Integer = 0
            Dim total_qty As Integer = 0
            Dim t As Integer = 0
            Dim plan_seq As String = ""
            Dim lot_sep As String = ""
            Dim tag_number As String = ""
            If Len_length = 103 Then 'Fa '
                qty_s = scan_qty.Text.Substring(55, 3)
                plan_seq = scan_qty.Text.Substring(16, 3)
                lot_sep = scan_qty.Text.Substring(58, 4)
                tag_number = scan_qty.Text.Substring(100, 3)
                tag_seq = scan_qr.Substring(87, 16) 'plan_seq & lot_sep & tag_number
                Dim check_arr = tag_seq.Split(" ")
                Dim i As Integer = 0
                For Each value As String In check_arr
                    If check_arr(i) <> "" Then
                        tag_seq = check_arr(i)
                        GoTo out
                    End If
                    i = i + 1
                Next
out:
                order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
            ElseIf Len_length = 62 Then 'web post'
                qty_s = scan_qty.Text.Substring(51, 8)
                tag_seq = scan_qty.Text.Substring(59, 3)
                order_number = scan_qty.Text.Substring(2, 10)
            End If
            If Module1.total_database = 0 Then
                t = best_qty - qty_s
                Module1.total_qty = best_qty - qty_s
                Module1.total_database = Module1.total_database + 1
            Else
                'MsgBox("รอบมากกว่า1")
                Module1.total_qty = Module1.total_qty - qty_s
                t = Module1.total_qty
                'MsgBox("ค่าคงเหลือรอบปัจจุบัน = " & Module1.total_qty)
                Module1.total_database = Module1.total_database + 1
            End If
            If Module1.total_qty = 0 Then
                t = 0
                com_flg = 1
                Dim pase_number_qty_scan As Integer = 0
                If Len(scan_qty.Text) = 62 Then
                    pase_number_qty_scan = scan_qty.Text.Substring(51, 8)
                    tag_seq = scan_qty.Text.Substring(59, 3)
                    text_tmp.Text = pase_number_qty_scan
                    order_number = scan_qty.Text.Substring(2, 10)
                ElseIf Len(scan_qty.Text) = 103 Then
                    order_number = scan_qty.Text.Substring(58, 4) 'LOT FA'
                End If
            End If 'เพิ่มมาใหม่ BTN2 '
            If Module1.total_qty > 0 Then
                t = 0
                com_flg = 1
            End If
            Dim t_string As String = "no data"
            If t < 0 Then
                t_string = t
                t = t_string.Substring(1)
            End If
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
            Module1.show_data_supply = Module1.G_show_data_supply
            set_show_remain() 'show RM'
            set_show_supply()
            Dim box_control As Integer = 0
            Dim status_box As Integer = 0
            Try
                Dim check_box = "SELECT CASE WHEN MAX (BOX_CONTROL)  IS NULL THEN 0 ELSE MAX (BOX_CONTROL) END AS box FROM sup_scan_pick_detail WHERE SLIP_CD = '" & Module1.SLIP_CD & "'"
                Dim cmd_check_box As SqlCommand = New SqlCommand(check_box, myConn)
                reader = cmd_check_box.ExecuteReader()
                If reader.Read() Then
                    If reader("box").ToString() = "0" Then
                        status_box = 1
                        reader.Close()
                        GoTo check_box_part
                    Else
                        count_box_check_qr_part += 1
                        If count_box_check_qr_part = 1 Then
                            box_control = CDbl(Val(reader("box").ToString())) + 1
                        Else
                            reader.Close()
                            GoTo plus_box
                        End If
                        reader.Close()
                    End If

                End If
check_box_part:
                If status_box = 1 Then
plus_box:
                    Dim check_box_part = "SELECT CASE WHEN MAX (BOX_CONTROL) IS NULL THEN 0 ELSE MAX (BOX_CONTROL) END AS box FROM check_qr_part WHERE SLIP_CD = '" & Module1.SLIP_CD & "'"
                    Dim cmd_check_box_part As SqlCommand = New SqlCommand(check_box_part, myConn)
                    reader = cmd_check_box_part.ExecuteReader()
                    If reader.Read() Then
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
            Dim sql_get_date = "SELECT GETDATE() as date_now"
            Dim cmd_get_date As SqlCommand = New SqlCommand(sql_get_date, myConn)
            reader = cmd_get_date.ExecuteReader()
            Dim date_now_database As String = ""
            If reader.Read() Then
                date_now_database = reader("date_now").ToString()
            End If
            reader.Close()
            strCommand2 = "EXEC [dbo].[INSERT_CHECK_QR_PART] @wi = '" & Module1.FG_CUS_ORDER_ID & "'  , @item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' , @scan_qty = '" & Trim(scan_qty.Text.Substring(52, 6)) & "' , @scan_lot = '" & order_number & "' , @tag_typ = '1' , @tag_readed = '" & scan_qr & "' , @scan_emp = '" & Module1.A_USER_ID & "',  @term_cd = '" & S_number & "', @updated_date = '" & date_now_database & "', @tag_seq = '" & tag_seq & "', @tag_remain_qty = '" & t & "', @MENU_ID = '" & Module1.MENU_ID & "', @line_cd = '" & scan_qty.Text.Substring(2, 6) & "', @delivery_date = '" & delivery_date & "', @SLIP_CD = '" & Module1.SLIP_CD & "', @flg = '" & com_flg & "', @BOX_CONTROL = '" & box_control & "'"
            'strCommand2 = "INSERT INTO check_qr_part (wi,item_cd,scan_qty,scan_lot,tag_typ,tag_readed,scan_emp,term_cd,updated_date,updated_by,tag_seq,S_number , com_flg ,tag_remain_qty  , CREATE_DATE , CREATE_BY , MENU_ID , line_cd , Delivery_date , SLIP_CD , BOX_CONTROL) VALUES ('" & Module1.FG_CUS_ORDER_ID & "','" & Module1.FG_PART_CD.Substring(16) & "','" & Trim(scan_qty.Text.Substring(52, 6)) & "','" & order_number & "','1','" & scan_qr & "','" & Module1.A_USER_ID & "','" & S_number & "','" & date_now & "','" & Module1.A_USER_ID & "','" & tag_seq & "','" & S_number & "','" & com_flg & "','" & t & "' , '" & date_now & "' , '" & Module1.A_USER_ID & "', '" & Module1.MENU_ID & "' , '" & scan_qty.Text.Substring(2, 6) & "' , '" & Module1.delivery_date & "' , '" &  & "', '" & box_control & "')"
            'reader.Close()
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
            arr_check_qr_remain_lot.Add(order_number)
            arr_check_qr_remain_seq.Add(tag_seq)
            If Len_length = 62 Then
                count_scan = count_scan + 1
                check_lot_scan_web_post()
            Else
                check_lot_scan_fw()
            End If
        Catch ex As Exception
            MsgBox("SORRY Insert ERROR" & vbNewLine & ex.Message)
            MsgBox(strCommand2)
        End Try
    End Sub
    Public Function sup_scan_pick_detail(ByVal count As String, ByVal F_wi As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal scan_lot As String, ByVal tag_typ As String, ByVal tag_readed As String, ByVal scan_emp As String, ByVal term_cd As String, ByVal updated_date As String, ByVal updated_by As String, ByVal updated_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal create_date As String, ByVal create_by As String, ByVal line_cd As String, ByVal delivery_date As String, ByVal box_control As String)
        Dim Len_length As Integer = length
        Dim strCommand2 As String = "no data"
        Dim PO = scan_lot 'lot คือ PO'
        Try
            ' strCommand2 = "INSERT INTO sup_scan_pick_detail (wi , item_cd , scan_qty ,scan_lot , tag_typ , tag_readed , scan_emp, term_cd , updated_date , updated_by , tag_seq  , com_flg , tag_remain_qty , CREATE_DATE , CREATE_BY , MENU_ID , line_cd , delivery_date , SLIP_CD , BOX_CONTROL )VALUES ('" & F_wi & "' ,'" & F_item_cd & "','" & scan_qty & "' ,'" & scan_lot & "','" & tag_typ & "','" & tag_readed & "','" & scan_emp & "','" & term_cd & "','" & updated_date & "','" & updated_by & "','" & updated_seq & "','" & com_flg_table & "','" & tag_remain_qty & "' , '" & create_date & "' , '" & create_by & "', '" & Module1.MENU_ID & "' , '" & line_cd & "', '" & delivery_date & "' , '" & Module1.SLIP_CD & "', '" & box_control & "')"
            Dim sql_get_date = "SELECT GETDATE() as date_now"
            Dim cmd_get_date As SqlCommand = New SqlCommand(sql_get_date, myConn)
            reader = cmd_get_date.ExecuteReader()
            Dim date_now_database As String = ""
            If reader.Read() Then
                date_now_database = reader("date_now").ToString()
            End If
            reader.Close()
            strCommand2 = "EXEC [dbo].[INSERT_SUP_SCAN_PICK_DETAIL] @wi = '" & F_wi & "'  , @item_cd = '" & F_item_cd & "' , @scan_qty = '" & scan_qty & "' , @scan_lot = '" & scan_lot & "' , @tag_typ = '" & tag_typ & "' , @tag_readed = '" & tag_readed & "' , @scan_emp = '" & scan_emp & "',  @term_cd = '" & term_cd & "', @updated_date = '" & date_now_database & "', @tag_seq = '" & updated_seq & "', @tag_remain_qty = '" & tag_remain_qty & "', @MENU_ID = '" & Module1.MENU_ID & "', @line_cd = '" & line_cd & "', @delivery_date = '" & delivery_date & "', @SLIP_CD = '" & Module1.SLIP_CD & "', @flg = '" & com_flg_table & "', @BOX_CONTROL = '" & box_control & "'"
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
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
        Panel7.Visible = True
        alert_loop.Visible = True
        status_alert_image = "loop_re_scan_fa"
        text_box_success.Visible = True
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
        Dim S_number As String = main.scan_terminal_id
        Dim qty_s As Integer = 0
        If Len_length = 103 Then 'Fa '
            qty_s = scan_qty.Text.Substring(55, 3)
        ElseIf Len_length = 62 Then 'web post'
            qty_s = scan_qty.Text.Substring(51, 8)
        End If
        Dim sum_qty = qty_database - qty_s
        Dim com_flg As String = "no data"
        Try
            If sum_qty > 0 Then
                sum_qty = 0
                com_flg = 1
            End If
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
            reader.Close()
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)

            Dim str_update As String = "update sup_scan_pick_detail set com_flg  = '" & com_flg & "' ,tag_remain_qty = '" & sum_qty & "'  , updated_date = '" & up_date_date & "' , updated_by = '" & up_date_by & "' where id = '" & id & "'"
            Dim command2 As SqlCommand = New SqlCommand(str_update, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
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
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
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
        count_box_check_qr_part = 0
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
        'PD_ADD_PART.Show()
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
                Dim command123 As SqlCommand = New SqlCommand(strCommand123, myConn)
                reader = command123.ExecuteReader() 'ติดบรรทัดนี้'
                Dim id As Integer = 0
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
        Dim strCommand As String = "SELECT COUNT(id) as c, com_flg  as com_flg   FROM check_qr_part   where item_cd = '" & Module1.FG_PART_CD.Substring(16) & "' and scan_lot = '" & order_number & "' and tag_seq = '" & tag_seq & "' and line_cd = '" & scan_qty.Text.Substring(2, 6) & "'and com_flg = '1' group by com_flg"
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read = True
            count = reader("c").ToString()
        Loop
        reader.Close()
        Dim check_reprint_c As Integer = 0

        If count = 1 Then
            text_tmp.Text = scan_qty_total
        End If
        check_count__data = scan_qty_total
        Return count
    End Function
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
                If code_suppier = scan_code_suppier And po = scan_po Then
                    If totall_qty_scan >= QTY Then
                        If check_po_lot = "pick_ok" Then
                            Dim qty_add As Double = 0.0
                            Dim data_remain As Double = 0.0
                            qty_add = CDbl(Val(QTY))
                            data_remain = totall_qty_scan - qty_add
                            If data_remain >= 0 Then 'การเก็บ remain'
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
                ' For Each key In Module1.arr_pick_detail_lot
                Dim lot As String = Module1.arr_pick_detail_lot(c_num).ToString
                'Dim po As String = Module1.arr_pick_detail_po(num).ToString
                Dim QTY As String = Module1.arr_pick_detail_qty(c_num).ToString
                QTY_INSERT_LOT_PO = 0
                If lot_scan = lot Then
                    If totall_qty_scan >= QTY Then
                        If check_po_lot = "pick_ok" Then
                            Dim qty_add As Double = 0.0
                            Dim data_remain As Double = 0.0
                            qty_add = CDbl(Val(QTY))
                            data_remain = totall_qty_scan - qty_add

                            If data_remain >= 0 Then 'การเก็บ remain'
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
        Dim check_fa = check_FA_TAG_FG()
        If check_fa = False Then
            status = 4
        ElseIf check_fa = 1 Then
            status = 5
        End If
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
    Public Sub next_image(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles text_box_success.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.F3
                Panel7.Visible = False
                alert_success.Visible = False
            Case System.Windows.Forms.Keys.F4

            Case System.Windows.Forms.Keys.Enter
                If check_process = "OK" Then
                    Try
                        status_alert_image = ""
                        Dim plan_fg = New Select_plan_fg()
                        plan_fg.Show()
                        Me.Close()
                    Catch ex As Exception
                        MsgBox("NO TRY")
                    End Try
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

    Public Sub insert_pick_log(ByVal REMAIN_ID As String, ByVal F_wi As String, ByVal used_qty As Double, ByVal create_date As String, ByVal create_by As String, ByVal updated_date As String, ByVal updated_by As String, ByVal scan_lot As String, ByVal updated_seq As String, ByVal F_item_cd As String)
        Try
            Dim strCommand123 As String = "SELECT   top 1 id as i   FROM sup_scan_pick_detail  where scan_lot = '" & scan_lot & "' and tag_seq = '" & updated_seq & "'  and item_cd ='" & F_item_cd & "'"
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
            Dim get_id_log = "select top 1 * from sup_pick_log where REF_ID = '" & REMAIN_ID & "' and WI_NO = '" & F_wi & "' and USED_QTY = '" & used_qty & "'  and CREATED_DATE = '" & create_date & "' and CREATED_BY = '" & create_by & "' and UPDATED_DATE = '" & updated_date & "' and updated_by = '" & updated_by & "'"
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
                Dim arr_qty = tag_read.Split(" ")
                If tag_read.Substring(56, 1) <> " " Then
                    Try
                        If F_item_cd.Substring(11, 1) = "G" Then
                            data_key_up = arr_qty(31)
                        Else
                            data_key_up = arr_qty(32)
                        End If
                    Catch ex As Exception
                        data_key_up = arr_qty(32)
                    End Try
                Else
                    data_key_up = arr_qty(33)
                End If
            End If
            Dim SEQ = "NO_DATA"
            If tag_read.Substring(94, 1) <> " " Then
                SEQ = updated_seq.Substring(8, 3)
            End If
            If tag_read.Substring(94, 1) = " " Then
                data_key_up = tag_read.Substring(95, 8)
                SEQ = tag_read.Substring(95, 3)
                If tag_read.Substring(95, 1) = " " Then
                    data_key_up = data_key_up.Substring(3)
                    SEQ = " "
                End If
            End If
            Try
                If F_item_cd.Substring(11, 1) = "G" Then
                    F_item_cd = F_item_cd.Substring(0, 11)
                End If
            Catch ex As Exception
            End Try
            Dim get_id_log = "select top 1 * from FA_TAG_FG where ITEM_CD = '" & F_item_cd & "' AND TAG_SEQ = '" & SEQ & "'AND LOT_NO = '" & scan_lot & "' and KEY_UP = '" & data_key_up & "' and LINE_CD = '" & tag_read.Substring(2, 6) & "'"
            Dim cmd_get As SqlCommand = New SqlCommand(get_id_log, myConn_fa)
            reader = cmd_get.ExecuteReader()
            Dim update_qty As Double = 0.0
            If reader.Read Then
                update_qty = CDbl(Val(reader("TAG_QTY").ToString())) - CDbl(Val(used_qty))
            End If
            reader.Close()
            Dim FLG_STATUS As String = "0"
            If update_qty <= 0 Then
                update_qty = 0
                FLG_STATUS = "1"
            Else
                FLG_STATUS = "2"
            End If
            Dim str_update_qty = "EXEC [dbo].[cut_stock_pick_fg] @qty = '" & used_qty & "'  , @flg_status = '" & FLG_STATUS & "' , @item_cd = '" & F_item_cd & "' , @seq = '" & SEQ & "' , @lot = '" & scan_lot & "' , @KEY_UP = '" & data_key_up & "' , @LINE_CD = '" & tag_read.Substring(2, 6) & "'"
            Dim cmd_update As SqlCommand = New SqlCommand(str_update_qty, myConn_fa)
            reader = cmd_update.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("FAILL cut_stock_FASYSTEM" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub

    Public Function check_FA_TAG_FG()
        Try
            Dim data_key_up As String = ""
            Dim KEY_UP = scan_qty.Text.Substring(58, 4)
            Dim SEQ As String = ""
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
            ElseIf scan_qty.Text.Substring(73, 2) = "E " Then
                arr_qty = scan_qty.Text.Split(" ")
                data_key_up = arr_qty(31)
            Else
                arr_qty = scan_qty.Text.Split(" ")
                If scan_qty.Text.Substring(73, 1) = "G" Then
                    data_key_up = arr_qty(31)
                ElseIf scan_qty.Text.Substring(56, 1) <> " " Then
                    data_key_up = arr_qty(32)
                Else
                    data_key_up = arr_qty(33)
                End If
            End If
            Dim ITEM_CD = ""
            Dim lot_sep = ""
            If scan_qty.Text.Substring(94, 1) <> " " Then
                Dim arr_item_cd = scan_qty.Text.Split(" ")
                ITEM_CD = arr_item_cd(0).Substring(19)
                lot_sep = scan_qty.Text.Substring(58, 4)
            Else
                Dim arr_item_cd = scan_qty.Text.Split(" ")
                ITEM_CD = arr_item_cd(11)
                lot_sep = scan_qty.Text.Substring(58, 4)
            End If
            If scan_qty.Text.Substring(94, 1) <> " " Then
                SEQ = data_key_up.Substring(8, 3)
            End If
            If scan_qty.Text.Substring(94, 1) = " " Then 'ตรวจสอบ TAG ว่าเป็น TAG REPRINT หรือไม่'
                If scan_qty.Text.Substring(95, 1) <> " " Then 'reprint ถูก'
                    data_key_up = scan_qty.Text.Substring(95, 8)
                    SEQ = scan_qty.Text.Substring(95, 3)
                End If
                Dim check_arr = data_key_up.Split(" ")
                Dim i As Integer = 0
                If scan_qty.Text.Substring(95, 1) = " " Then 'TAG REPRINT ผิด'
                    data_key_up = scan_qty.Text.Substring(98, 5) 'ผิดตรงนี้ TAG REPRINT '
                    SEQ = " "
                End If
            End If
query:
            If scan_qty.Text.Substring(19, 14) = "49373-25591-EP" Then
                data_key_up = arr_qty(40)
                lot_sep = arr_qty(15).Substring(2, 4)
                ITEM_CD = scan_qty.Text.Substring(19, 14)
                SEQ = data_key_up.Substring(8, 3)
            End If
            Dim sql As String = "select top 1 IND , FLG_STATUS from FA_TAG_FG where LOT_NO = '" & lot_sep & "' and ITEM_CD = '" & ITEM_CD & "' and TAG_SEQ = '" & SEQ & "' and LINE_CD = '" & scan_qty.Text.Substring(2, 6) & "' and KEY_UP = '" & data_key_up & "'"
            Dim command2 As SqlCommand = New SqlCommand(sql, myConn_fa)
            reader = command2.ExecuteReader()
            Dim count As String = "NO_DATA"
            Dim flg_status As String = "NO_DATA"
            Do While reader.Read = True
                count = reader("IND").ToString()
                flg_status = reader("FLG_STATUS").ToString()
            Loop
            reader.Close()
            If count <> "NO_DATA" Then
                If flg_status <> "2" Then
                    Return 1
                Else
                    Return 2
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("FALL please check")
        End Try
        Return 0
    End Function

    Private Sub TextBox2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub



    Private Sub check_qr_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles check_qr.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                check_qr.Visible = False
                Panel7.Visible = False
                alert_no_tranfer_data.Visible = False
                alert_14_day.Visible = False
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
    Private Sub alert_washing_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles check_qr.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                check_qr.Visible = False
                Panel7.Visible = False
                alert_14_day.Visible = False
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
    Public Function check_washing()
        Dim strCommand = "select MAX(IND) , UPDATED_DATE from FA_TAG_FG  where READ_QR= '" & scan_qty.Text & "' and FLG_STATUS = '2' group by UPDATED_DATE"
        Dim command2 As SqlCommand = New SqlCommand(strCommand, myConn_fa)
        reader = command2.ExecuteReader()
        Dim date_scan_data As String = scan_qty.Text.Substring(44, 8)
        Dim status_check_data As String = "0"
        If reader.Read = True Then
            If reader("UPDATED_DATE").ToString() IsNot Nothing Then
                date_scan_data = reader("UPDATED_DATE").ToString()
                status_check_data = "1"
            End If
        End If
        reader.Close()
        If status_check_data = "1" Then
            Dim time As DateTime = DateTime.Now
            Dim date_washing_string As String = date_scan_data.Substring(0, 4) & "-" & date_scan_data.Substring(4, 2) & "-" & date_scan_data.Substring(6, 2)
            Dim date_washing As DateTime = date_washing_string
            Dim format As String = "yyyy-MM-dd"
            Dim date_washing_format = time.ToString(format)
            Dim time_washing As DateTime = date_washing.AddDays(14)
            Dim format_tommorow = "yyyy-MM-dd"
            Dim date_washing_check = time_washing.ToString(format_tommorow)
            Dim status As Integer = 0
            Dim time_now As DateTime = DateTime.Now
            Dim format_now As String = "yyyy-MM-dd"
            Dim date_now = time_now.ToString(format_now)
            If date_now >= date_washing_check Then
                status_check_washing = 1
            Else
                status_check_washing = 0
            End If
        Else
            status_check_washing = 2
        End If
        status_check_washing = 0 'ชั่วคราว'
        Return status_check_washing
        'Return 0
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
recheck_net:
        If count_net = 5000 Then
            If Api.check_net <> True Then
                Timer1.Enabled = False
                ' MsgBox("อินเตอร์เน็ตไม่เสถียร กรุณา กด ENT เพื่อ รอ INTERNET")
                Timer1.Enabled = True
                GoTo recheck_net
            Else
                count_net = 0
            End If
        Else
            count_net += 1
        End If
    End Sub
    Public Sub set_qty_check_qr_part()
        Dim sql As String = "select sum(scan_qty) as TOTAL_QTY from check_qr_part where item_cd = '" & Part_No.Text.Substring(16) & "' and term_cd = '" & main.scan_terminal_id & "'"
        Dim command2 As SqlCommand = New SqlCommand(sql, myConn)
        reader = command2.ExecuteReader()
        Dim count As String = "NO_DATA"
        Dim flg_status As String = "NO_DATA"
        Dim check_qty As Double = 0
        Do While reader.Read = True
            text_tmp.Text = reader("TOTAL_QTY").ToString()
            check_qty = CDbl(Val(reader("TOTAL_QTY").ToString()))
        Loop
        reader.Close()
        If check_qty > 0 Then
            show_number_supply.Text = text_tmp.Text
            show_number_remain.Text = CDbl(Val(show_qty.Text.Substring(6))) - CDbl(Val(show_number_supply.Text))
            want_to_tag.Text = CDbl(Val(show_qty.Text.Substring(6))) - CDbl(Val(show_number_supply.Text))
            scan_qty_total = show_number_supply.Text
        Else
            show_number_supply.Text = 0
        End If
    End Sub
End Class