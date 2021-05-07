Imports System.Data.SqlClient
Imports PICKING_SYSTEM.Page_projects
Imports System.Data
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System
Imports System.IO
Imports System.Net

Public Class Select_plan_fg
    'Public result_get_CUST_ODR_NO As String = "NO_DATA"
    'Public result_get_ITEM_NO As String = "NO_DATA"
    'Public result_get_REMAIN_SHIP_QTY As String = "NO_DATA"
    'Public result_get_part_name As String = "NO_DATA"
    Dim result_order_id As ArrayList = New ArrayList()
    Dim result_Line As ArrayList = New ArrayList()
    Dim result_get_CUST_ODR_NO As ArrayList = New ArrayList()
    Dim result_get_ITEM_NO As ArrayList = New ArrayList()
    Dim result_get_REMAIN_SHIP_QTY As ArrayList = New ArrayList()
    Dim result_get_part_name As ArrayList = New ArrayList()
    Dim result_get_model As ArrayList = New ArrayList()
    Dim SLIP_CD As ArrayList = New ArrayList()
    Public status As Integer = 0
    Public ml As Integer = 0
    Public count_time As Integer = 0
    'Dim myConn As SqlConnection
    Public myConn = "NO"
    Public myconn_fa = "NO"
    Dim x As ListViewItem
    Dim sel_itemSpa As String = "                        "
    Dim sel_where As String = ""
    Public Part_Selected As String = ""
    Public Part_Name As String = ""
    Public lo As String = ""
    Public QTY As String = ""
    Public Lot_No As String = ""
    Public PD3 As Select_PD
    Public index_
    Public FG_part_id As String = "NO_DATA"
    Public FG_part_name As String = "NO_DATA"
    Public FG_QTY As String = "NO_DATA"
    Public Trip As String = "NO_DATA"
    Dim wi As String = ""
    Public part As part_detail
    Public line, counter As String
    Dim arr_LVL As ArrayList = New ArrayList()
    Dim arr_QTY As ArrayList = New ArrayList()
    Dim arr_com_flg As ArrayList = New ArrayList()
    Dim arr_order_no As ArrayList = New ArrayList()
    Dim arr_item_cd As ArrayList = New ArrayList()
    Dim arr_QTY_FG As ArrayList = New ArrayList()
    Dim delivery_date As ArrayList = New ArrayList()
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
    Dim count_arr_fw As Integer = 0
    Dim F_id_sup_ As ArrayList = New ArrayList()
    Dim F_menu_CANCLE As ArrayList = New ArrayList()
    Dim F_DEL_CANCLE As ArrayList = New ArrayList()
    Dim F_SLIP_CD_CANCLE As ArrayList = New ArrayList()
    Dim F_Box_control_CANCLE As ArrayList = New ArrayList()
    Public F_DELIVERY_DATE As ArrayList = New ArrayList()
    Dim check_data As ArrayList = New ArrayList()
    Public check_action As Integer = 0
    Public index_list_view As Integer = 0
    Public g_index As Integer = 0
    Dim reader As SqlDataReader
    Dim F_id_sup As ArrayList = New ArrayList()
    Dim dat As String = String.Empty
    Dim M_SLIP_CD As String = "NO_DATA"
    Public count_net As Integer = 0


    Private Sub selLine_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
LOOP_MAIN_OPEN:
re_connect:
            If Api.check_net() = True Then
                Dim connect_db = New connect()
                myConn = connect_db.conn()
                myconn_fa = connect_db.conn_fa()
            Else
                Timer1.Enabled = False
                MsgBox("อินเตอร์เน็ตไม่เสถียร กรุณา กด ENT เพื่อ รอ INTERNET")
                GoTo re_connect
            End If
        Finally
            'Module1.M_QTY_LOT_ALL = 0
            'combobox_line()
            'load_data()
            set_day()
            days.SelectedIndex = 1
            alert_check_part_no.Visible = False
            p_show_confrim.Visible = False
            Line_Emp_cd.Text = main.show_code_id_user()
            check_action = 2
            Panel3.Visible = False
            Panel8.Visible = False
            OK_CON.Visible = False
            ComboBox1.Items.Add("8")
            ComboBox1.Items.Add("10")
            If Module1.PHASE = "51" Then
                ComboBox1.SelectedIndex = 1
            ElseIf Module1.PHASE = "52" Then
                ComboBox1.SelectedIndex = 0
            End If
            ComboBox2.Items.Add("09:20")
            ComboBox2.Items.Add("16:00")
            ComboBox2.Items.Add("22:40")
            ComboBox2.Items.Add("04:00")
            ComboBox2.SelectedIndex = 0
            load_combobox1()
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Public Function get_data_qty_po_lot_fa(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String, ByVal QTY As String, ByVal data_total As Integer)
        Module1.M_CHECK_TYPE = "FW"
        Dim strCommand As String = "SELECT B.fa_lot AS LT, SUM (B.fa_total) AS fa_total, SUM (B.fa_use) AS fa_use FROM sup_frith_in_out_fa B WHERE B.fa_item_cd = '" & code_part & "' AND B.fa_com_flg = '0' GROUP BY B.fa_lot, B.fa_item_cd" 'แบบ ไม่ check wi FW
        'Dim strCommand As String = "SELECT B.fa_lot AS LT, SUM (B.fa_total) AS fa_total, SUM (B.fa_use) AS fa_use FROM sup_frith_in_out_fa B WHERE B.fa_item_cd = '" & code_part & "' AND B.fa_com_flg = '0' GROUP BY B.fa_lot, B.fa_item_cd" 'แบบ check wi FW
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim number As Integer = 1
        Dim RESULT_QTY As Integer = Module1.check_QTY
        Dim Lot_plus_qty As Integer = 0
        Dim QTY_PROCESS As Integer = 0
        'MsgBox("Module1.check_QTY = " & Module1.check_QTY)
        Dim lot_qty As String = "0"
        Dim N_lot_qty As Integer = 0
        Dim N_fa_use As Integer = 0
        Dim NUM_QTY As Integer = 0
        Dim arr_count_lot As ArrayList = New ArrayList()
        Do While reader.Read()
            lot_qty = reader("fa_total").ToString
            'MsgBox("lot_qty = " & lot_qty)
            Dim fa_use = reader("fa_use").ToString
            'MsgBox("fa_use = " & fa_use)
            N_lot_qty = Integer.Parse(lot_qty)
            N_fa_use = Integer.Parse(fa_use)
            NUM_QTY = N_lot_qty - N_fa_use
            'MsgBox("NUM_QTY = " & NUM_QTY)
            Lot_plus_qty = Lot_plus_qty + NUM_QTY
            'MsgBox("Lot_plus_qty = " & Lot_plus_qty)
            'MsgBox("Lot_plus_qty = " & Lot_plus_qty)
            If Module1.check_QTY <= Lot_plus_qty.ToString Then 'check ว่า QTY ที่ จะไป pick กับ QTY ของ LOT คือ ถ้า QTY ที่จะเอาน้อยกว่าหรือ= QTY ใน LOT ก็เอาไปได้เลบ'
                'Module1.arr_pick_detail_po.Add(reader("WI").ToString)
                Module1.arr_pick_detail_lot.Add(reader("LT").ToString)
                arr_count_lot.Add(reader("LT").ToString)
                If number = 1 Then
                    Module1.arr_pick_detail_qty.Add(QTY)
                    ' MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                ElseIf number > 1 Then
                    Module1.arr_pick_detail_qty.Add(RESULT_QTY.ToString)
                    'MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                End If
                GoTo NEXT_END_FW
            End If 'กรณี มี LOT ที่ พอดีอันเดียว'
            If Module1.check_QTY > Lot_plus_qty.ToString Then
                'Module1.arr_pick_detail_po.Add(reader("WI").ToString)
                Module1.arr_pick_detail_qty.Add(NUM_QTY)
                Module1.arr_pick_detail_lot.Add(reader("LT").ToString)
                arr_count_lot.Add(reader("LT").ToString)
                RESULT_QTY = RESULT_QTY - NUM_QTY
            End If 'กรณี มี LOT ที่ มากกว่า 1'
            number = number + 1
        Loop
NEXT_END_FW:
        reader.Close()
        Dim c As Integer = 0
        For Each key5 In arr_count_lot
            Dim str_check_lot As String = "SELECT COUNT (fa_id) as c FROM sup_frith_in_out_fa WHERE fa_item_cd = '" & code_part & "' AND fa_lot = '" & arr_count_lot(c) & "' AND fa_com_flg = '0' GROUP BY fa_lot, fa_item_cd"
            Dim command_lot As SqlCommand = New SqlCommand(str_check_lot, myConn)
            reader = command_lot.ExecuteReader()
            If reader.Read Then
                If reader("c").ToString > 1 Then
                    Module1.M_CHECK_LOT_COUNT_FW.Add("2")
                Else
                    Module1.M_CHECK_LOT_COUNT_FW.Add("1")
                End If
            Else
                Module1.M_CHECK_LOT_COUNT_FW.Add("1")
            End If
            '  MsgBox("Module1.M_CHECK_LOT_COUNT_FW = " & Module1.M_CHECK_LOT_COUNT_FW(0))
            reader.Close()
            c = c + 1
        Next
        Return 0
    End Function
    Private Sub Lb2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Public Sub combobox_line()
        Dim subSelpd = Module1.data_combo
        Dim strCommand As String = "SELECT DISTINCT LINE_CD from sup_work_plan_supply_dev where PD like '%" & subSelpd & "%' order by LINE_CD asc"
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim num As Integer
        num = 0
        Dim status = 0
        Do While reader.Read()
            'ComboBox1.Items.Add(reader.Item(0))
            num = num + 1
            status = 1
        Loop
        reader.Close()
        If status = 1 Then
            ' ComboBox1.SelectedIndex = 0
        End If
    End Sub
    Public Function get_data_qty_po_lot(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String)
        Module1.M_CHECK_TYPE = "WEB_POST"
        Dim strCommand As String = "SELECT sd.*,F_I_O.LT ,F_I_O.QTY_OF_LOT , F_I_O.PO  FROM ( SELECT ss_f.item_cd, ss_f.com_flg,ss_f.LOT_RECEIVE as LT , ss_f.qty as QTY_OF_LOT ,ss_f.PUCH_ODR_CD as PO  , SUBSTRING (LOT_RECEIVE, 6, 8) AS date_lot FROM sup_frith_in_out ss_f WHERE ss_f.com_flg <> 1 ) F_I_O INNER JOIN sup_work_plan_supply_dev sd ON F_I_O.item_cd = sd.ITEM_CD WHERE sd.line_cd = '" & code_line & "' AND sd.item_cd = '" & code_part & "' AND sd.wi = '" & code_wi & "' AND ( ps_unit_numerator <> '' AND sd.location_part <> '' ) AND sd.pick_flg != 1 AND sd.WORK_ODR_DLV_DATE  = '" & Module1.date_now_database & "' ORDER BY sd.wi, F_I_O.date_lot ASC" 'ใช้ที่บ้านแบบใหม่
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim number As Integer = 1
        Dim RESULT_QTY As Integer = Module1.check_QTY
        Dim Lot_plus_qty As Integer = 0
        'MsgBox("Module1.check_QTY = " & Module1.check_QTY)
        Dim lot_qty As String = "0"
        Dim N_lot_qty As Integer = 0
        Dim totala_scan_qty As Double = 0.0

        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
            lot_qty = reader("QTY_OF_LOT").ToString
            N_lot_qty = Integer.Parse(lot_qty)
            Lot_plus_qty = Lot_plus_qty + N_lot_qty
            'MsgBox("Lot_plus_qty = " & Lot_plus_qty)
            If Module1.check_QTY <= Lot_plus_qty.ToString Then 'check ว่า QTY ที่ จะไป pick กับ QTY ของ LOT คือ ถ้า QTY ที่จะเอาน้อยกว่าหรือ= QTY ใน LOT ก็เอาไปได้เลบ'
                Module1.arr_pick_detail_po.Add(reader("PO").ToString)
                Module1.arr_pick_detail_lot.Add(reader("LT").ToString)
                If number = 1 Then
                    Module1.arr_pick_detail_qty.Add(totala_scan_qty)
                    ' MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                ElseIf number > 1 Then
                    Module1.arr_pick_detail_qty.Add(RESULT_QTY.ToString)
                    'MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                End If
                GoTo NEXT_END_WEB_POST
            End If 'กรณี มี LOT ที่ พอดีอันเดียว'
            If Module1.check_QTY > Lot_plus_qty.ToString Then
                Module1.arr_pick_detail_po.Add(reader("PO").ToString)
                Module1.arr_pick_detail_qty.Add(reader("QTY_OF_LOT").ToString)
                Module1.arr_pick_detail_lot.Add(reader("LT").ToString)
                RESULT_QTY = RESULT_QTY - reader("QTY_OF_LOT").ToString
            End If 'กรณี มี LOT ที่ มากกว่า 1'
            number = number + 1
        Loop
NEXT_END_WEB_POST:
        reader.Close()
        Return 0
    End Function

    Private Sub Panel2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel2.GotFocus

    End Sub
    Public Sub load_data()
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Panel9.Visible = True
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Panel3.Visible = False
    End Sub

    Private Sub Panel2_GotFocus_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ml += 1
        status += 5
recheck_net:
        'If count_net = 1000 Then
        'If Api.check_net <> True Then
        'MsgBox("อินเตอร์เน็ตไม่เสถียร กรุณา กด ENT เพื่อ รอ INTERNET")
        'GoTo recheck_net
        'Else
        'count_net = 0
        'End If
        'Else
        'count_net += 1
        'End If
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Panel9.Visible = False
    End Sub

    Private Sub Label8_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.ParentChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        load_combobox1()
    End Sub
    Public Sub load_combobox1()
        arr_com_flg = New ArrayList()
        Line_list_view.Items.Clear()
        result_order_id = New ArrayList
        result_get_CUST_ODR_NO = New ArrayList()
        result_get_ITEM_NO = New ArrayList()
        result_get_REMAIN_SHIP_QTY = New ArrayList()
        result_get_part_name = New ArrayList()
        delivery_date = New ArrayList()
        result_get_model = New ArrayList()
        SLIP_CD = New ArrayList()
        Try
            Dim time As DateTime = Today.AddDays("-1") 'DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)

            Dim time_tomorrow As DateTime = Today.AddDays(days.Text)
            Dim format_tommorow = "yyyy/MM/dd"
            Dim date_tommorow = time_tomorrow.ToString(format_tommorow)

            Dim time_detail As DateTime = DateTime.Now
            Dim format_time_detail As String = "HH:mm"
            Dim now_time_detail = time_detail.ToString(format_time_detail)

            Dim date_detail As DateTime = Today.AddDays("-1") 'DateTime.Now
            Dim format_date_detail As String = "yyyy/MM/dd"
            Dim now_date_detail = date_detail.ToString(format_date_detail)
            'result_get_CUST_ODR_NO = Api.Get_order_fg("http://192.168.161.102/exp_api3party/Api_get_order_fg/get_CUST_ODR_NO?phase=" & ComboBox1.Text & "")
            'result_get_ITEM_NO = Api.Get_order_fg("http://192.168.161.102/exp_api3party/Api_get_order_fg/get_ITEM_NO?phase=" & ComboBox1.Text & "")
            'result_get_REMAIN_SHIP_QTY = Api.Get_order_fg("http://192.168.161.102/exp_api3party/Api_get_order_fg/get_REMAIN_SHIP_QTY?phase=" & ComboBox1.Text & "")
            'result_get_part_name = Api.Get_order_fg("http://192.168.161.102/exp_api3party/Api_get_order_fg/get_name?phase=" & ComboBox1.Text & "")
            'Dim arr_cut_comma_ODR_NO = result_get_CUST_ODR_NO.Split(",")
            'Dim arr_cut_comma_item_no = result_get_ITEM_NO.Split(",")
            'Dim arr_cut_comma_REMAIN_SHIP_QTY = result_get_REMAIN_SHIP_QTY.Split(",")
            'Dim i As Integer = 0
            'For Each key In arr_cut_comma_item_no
            'x = New ListViewItem(arr_cut_comma_ODR_NO(i))
            'x.SubItems.Add(arr_cut_comma_item_no(i))
            'x.SubItems.Add(arr_cut_comma_REMAIN_SHIP_QTY(i))
            'Line_list_view.Items.Add(x)
            'i = i + 1
            'Next
            Dim date_start = now_date_detail & " 09:00"
            Dim date_end = date_tommorow & " 04:20"
            Dim phase As String = "NO_DATA"

            If ComboBox1.Text = "10" Then
                phase = "51"
            ElseIf ComboBox1.Text = "8" Then
                phase = "52"
            End If

            If ComboBox2.Text = "09:20" Then
                Trip = 1
            ElseIf ComboBox2.Text = "16:00" Then
                Trip = 2
            ElseIf ComboBox2.Text = "22:40" Then
                Trip = 3
            ElseIf ComboBox2.Text = "04:00" Then
                Trip = 4
            End If
            Dim sql As String = "EXEC dbo.Get_stock_fg_today  @phase ='" & phase & "', @date_start = '" & date_start & "' , @date_end = '" & date_end & "', @TRIP = '" & Trip & "'"
            Dim command_FA_TAG_FG As SqlCommand = New SqlCommand(sql, myconn_fa)
            reader = command_FA_TAG_FG.ExecuteReader()
            Dim i As Integer = 0
            Do While reader.Read()
                x = New ListViewItem(reader("CUST_ITEM_CD").ToString())
                x.SubItems.Add(reader("REMAIN_SHIP_QTY").ToString())
                x.SubItems.Add(reader("Delivery_date").ToString())
                delivery_date.Add(reader("Delivery_date").ToString())
                result_get_CUST_ODR_NO.Add(reader("CUST_ODR_NO").ToString())
                result_order_id.Add(reader("ORDER_FG_ID").ToString())
                result_get_ITEM_NO.Add(reader("CUST_ITEM_CD").ToString())
                result_get_REMAIN_SHIP_QTY.Add(reader("REMAIN_SHIP_QTY").ToString())
                result_get_part_name.Add(reader("CUST_ITEM_NAME").ToString())
                result_get_model.Add(reader("MODEL").ToString())
                arr_com_flg.Add(reader("COM_FLG").ToString)
                SLIP_CD.Add(reader("SLIP_CD").ToString)

435:
                Line_list_view.Items.Add(x)
                If reader("USE_QTY").ToString() <> "0" Then
                    Line_list_view.Items(i).BackColor = Color.Yellow
                End If
                If reader("COM_FLG").ToString() = "1" Then
                    Line_list_view.Items(i).BackColor = Color.FromArgb(136, 199, 249)
                End If
                If reader("COM_FLG").ToString() = "2" Then
                    Line_list_view.Items(i).BackColor = Color.FromArgb(103, 255, 103)
                End If
                i = i + 1
            Loop
            reader.Close()
            'set_data_home()
        Catch ex As Exception
            MsgBox("FAAAAAAAAAAAAALLLLLLLLLLLL ")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Timer1.Enabled = False
        main.Panel2.Visible = False
        Line_list_view.Items.Clear()
        Module1.M_CHECK_LOT_COUNT_FW = New ArrayList()
        main.Show()
        Me.Close()
    End Sub

    Private Sub btn_ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ok.Click
        Try
            Dim status = 0
            Dim FG_part As String = Module1.FG_PART_CD.Substring(16)
            'MsgBox(Module1.FG_PART_CD)
            Try
                If FG_part.Substring(11, 1) = "G" Then
                    FG_part = FG_part.Substring(0, 11)
                ElseIf FG_part.Substring(11, 1) = "E" Then
                    FG_part = FG_part.Substring(0, 11)
                End If
            Catch ex As Exception

            End Try
            '  Dim cut_fg = FG_part.Substring(0, 11)
            Dim sql As String = "select TOP 1.* from FA_TAG_FG where ITEM_CD = '" & FG_part & "'  order by LOT_NO asc"
            ' MsgBox("sql_check = " & sql)
            Dim command_FA_TAG_FG As SqlCommand = New SqlCommand(sql, myconn_fa)
            reader = command_FA_TAG_FG.ExecuteReader()
            If reader.Read Then
                Module1.FG_LOCATIONS = reader("LOCATION_PART").ToString
                Module1.FG_LINE = reader("LINE_CD").ToString
                status = 0
            Else
                status = 1
                alert_check_part_no.Visible = True
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
                ' MsgBox("READY")
                Bt.SysLib.Device.btBuzzer(1, stBuz)
                Bt.SysLib.Device.btVibrator(1, stVib)
                Bt.SysLib.Device.btLED(1, stLed)
                fo.Focus()
            End If
            reader.Close()
            Dim sql_stock As String = "select sum(TAG_QTY) as stock_qty from FA_TAG_FG where ITEM_CD = '" & FG_part & "' and FLG_STATUS = '2' "
            Dim commmand_stock As SqlCommand = New SqlCommand(sql_stock, myconn_fa)
            reader = commmand_stock.ExecuteReader()
            If reader.Read Then
                Module1.FG_ALL_QTY_STOCK = reader("stock_qty").ToString
            Else
            End If
            reader.Close()
            If status = 0 Then
                FIFO_FG()
                Dim p_scan As scan_location_fg = New scan_location_fg()

                If Module1.SLIP_CD <> "" Then
                    p_scan.Show()
                    Me.Hide()
                Else
                    MsgBox("NO SLIP_CD")
                End If

            End If
        Catch ex As Exception
            MsgBox("Please select Part")
        Finally
        End Try
    End Sub
    Private Sub Line_list_view_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Line_list_view.SelectedIndexChanged
        If IsNothing(Me.Line_list_view.FocusedItem) Then
        ElseIf Line_list_view.FocusedItem.Index >= 0 Then
            If Line_list_view.Items.Count > 0 Then
                Dim index As Integer = Line_list_view.FocusedItem.Index
                g_index = index
                '    MsgBox(index)
                'Dim(arr_selected = result_get_ITEM_NO.Split(","))
                'Dim arr_QTY_FG = result_get_REMAIN_SHIP_QTY.Split(",")
                'Dim arr_get_part_name = result_get_part_name.Split(",")
                Module1.FG_ORDER_ID = result_order_id(index)
                Module1.delivery_date = delivery_date(g_index)
                Module1.SLIP_CD = SLIP_CD(g_index)
                'MsgBox(Module1.SLIP_CD)
                Dim order_id = result_get_CUST_ODR_NO(index)
                Dim item_selected = result_get_ITEM_NO(index)
                Dim QTY_FG = result_get_REMAIN_SHIP_QTY(index)
                Dim part_name = result_get_part_name(index)
                Module1.FG_MODEL = result_get_model(index)
                Module1.FG_PHASE = ComboBox1.Text
                set_part(item_selected)
                set_QTY_FG(QTY_FG)
                set_part_name(part_name)
                set_CUST_ODR_NO(order_id)
                'MsgBox(arr_com_flg(index))
                If ComboBox2.Text = "09:20" Then
                    Trip = 1
                ElseIf ComboBox2.Text = "16:00" Then
                    Trip = 2
                ElseIf ComboBox2.Text = "22:40" Then
                    Trip = 3
                ElseIf ComboBox2.Text = "04:00" Then
                    Trip = 4
                End If
                Module1.Trip = Trip
                If arr_com_flg(index) = "0" Or arr_com_flg(index) = "2" Then
                    btn_ok.Visible = True
                    OK_CON.Visible = False
                End If
                If arr_com_flg(index) = "1" Then
                    btn_ok.Visible = False
                    OK_CON.Visible = True
                ElseIf arr_com_flg(index) = "2" Then
                    btn_ok.Visible = False
                    OK_CON.Visible = True
                End If

                'endddddddddddddddddddddddddd(b)'
                'set_part("49373-25590")
                'set_QTY_FG("400")
                'set_part_name("nameeeeee")
            End If
        Else
            MessageBox.Show("An Error has halted thid process")
        End If
    End Sub
    Public Sub set_CUST_ODR_NO(ByVal order_id As String)
        Module1.FG_CUS_ORDER_ID = order_id
    End Sub
    Public Sub set_part(ByVal item_selected As String)
        FG_part_id = "Part_Selected : " + item_selected
        Module1.FG_PART_CD = FG_part_id
    End Sub
    Public Sub set_QTY_FG(ByVal QTY As String)
        FG_QTY = "QTY : " + QTY
        Module1.FG_QTY = FG_QTY
    End Sub
    Public Sub FG_LINE(ByVal line As String)
        Module1.line = line
    End Sub
    Public Sub set_part_name(ByVal part_name As String)
        FG_part_name = "Part_Name : " + part_name
        Module1.FG_PART_NAME = FG_part_name
        'MsgBox("set FG_part_name" + FG_part_name)
    End Sub
    Public Function get_Part_Selected() As String
        Dim data = "Part_Selected : " + FG_part_id
        Return data
    End Function

    Public Function get_Part_Name() As String
        Dim data = "Part_Name : " + FG_part_name
        'MsgBox("get data" + FG_part_name)
        Return data
    End Function

    Public Function get_Location() As String
        Dim data = "Location : " + lo
        Return data
    End Function

    Public Function get_QTY_FG() As String
        Dim data = "QTY : " + FG_QTY
        Return data
    End Function

    Public Function get_Lot_No() As String
        Dim data = "Lot No : " + Lot_No
        Return data
    End Function

    Public Function get_wi() As String
        Dim data = wi
        Return data
    End Function

    Public Function FIFO_FG()
        Try
            Dim FG = Module1.FG_PART_CD.Substring(16)
            ' Dim cut_e = FG.Substring(0, 11)
            Try
                If FG.Substring(11, 1) = "G" Then
                    FG = FG.Substring(0, 11)
                ElseIf FG.Substring(11, 1) = "E" Then
                    FG = FG.Substring(0, 11)
                End If
            Catch ex As Exception

            End Try

            Dim stock_location As String = "NO_DATA"
            Dim sql_lot As String = "EXEC dbo.get_data_order_fg @FG = '" & FG & "'"
            Dim command_lot As SqlCommand = New SqlCommand(sql_lot, myconn_fa)
            reader = command_lot.ExecuteReader()
            Dim number As Integer = 1
            Dim RESULT_QTY As Integer = Module1.FG_QTY.Substring(6)
            Dim Lot_plus_qty As Integer = 0
            'MsgBox("Module1.check_QTY = " & Module1.check_QTY)
            Dim lot_qty As String = "0"
            Dim N_lot_qty As Integer = 0
            Dim totala_scan_qty As Double = 0.0
            Do While reader.Read()
                totala_scan_qty = CDbl(Val(reader("TAG_QTY").ToString))
                lot_qty = reader("QTY_OF_LOT").ToString
                N_lot_qty = Integer.Parse(lot_qty)
                Lot_plus_qty = Lot_plus_qty + N_lot_qty
                'MsgBox("Lot_plus_qty = " & Lot_plus_qty)
                If CDbl(Val(Module1.FG_QTY.Substring(6))) <= CDbl(Val(Lot_plus_qty.ToString)) Then 'check ว่า QTY ที่ จะไป pick กับ QTY ของ LOT คือ ถ้า QTY ที่จะเอาน้อยกว่าหรือ= QTY ใน LOT ก็เอาไปได้เลบ'
                    Module1.arr_pick_detail_po.Add("")
                    Module1.arr_pick_detail_lot.Add(reader("LOT_NO").ToString)
                    If number = 1 Then
                        If CDbl(Val(Module1.FG_QTY.Substring(6))) <= CDbl(Val(Lot_plus_qty.ToString)) Then
                            Module1.arr_pick_detail_qty.Add(Module1.FG_QTY.Substring(6))
                        Else '>='
                            Module1.arr_pick_detail_qty.Add(totala_scan_qty)
                        End If
                        ' MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                    ElseIf number > 1 Then
                        Module1.arr_pick_detail_qty.Add(RESULT_QTY.ToString)
                        'MsgBox("LOT 1 OK = " & reader("QTY").ToString)
                    End If
                    GoTo NEXT_END_WEB_POST
                End If 'กรณี มี LOT ที่ พอดีอันเดียว'
                If CDbl(Val(Module1.FG_QTY.Substring(6))) > CDbl(Val(Lot_plus_qty.ToString)) Then
                    Module1.arr_pick_detail_po.Add("")
                    Module1.arr_pick_detail_qty.Add(reader("QTY_OF_LOT").ToString)
                    Module1.arr_pick_detail_lot.Add(reader("LOT_NO").ToString)
                    RESULT_QTY = RESULT_QTY - reader("QTY_OF_LOT").ToString
                End If 'กรณี มี LOT ที่ มากกว่า 1'
                number = number + 1
            Loop
NEXT_END_WEB_POST:
            reader.Close()
        Catch ex As Exception
            MsgBox("FAILL cut_FIFO_FG" & vbNewLine & ex.Message, "FAILL")
        End Try
        Return 0
    End Function
    Public Sub set_data_home()
        Dim sql As String = "select * from FA_TAG_FG"
        Dim command As SqlCommand = New SqlCommand(sql, myconn_fa)
        reader = command.ExecuteReader()
        Do While reader.Read()
            x = New ListViewItem("12345678")
            x.SubItems.Add(reader("ITEM_CD").ToString)
            x.SubItems.Add("400")
            Line_list_view.Items.Add(x)
        Loop
        reader.Close()
    End Sub
    Private Sub fo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles fo.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                alert_check_part_no.Visible = False
        End Select
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        load_combobox1()
    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub

    Private Sub OK_CON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_CON.Click
        p_show_confrim.Visible = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        p_show_confrim.Visible = False
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Dim REMAIN_SHIP_QTY = 0
            Dim check_flg = "NO_DATA"
            Dim sql_get_data As String = "select * from  EXP_ORDER_FG  where SLIP_CD = '" & SLIP_CD(g_index) & "'"
            Dim cmd_get_fa As SqlCommand = New SqlCommand(sql_get_data, myconn_fa)
            reader = cmd_get_fa.ExecuteReader()
            Dim count As Integer = 0
            If reader.Read() Then
                check_flg = reader("COM_FLG").ToString
                REMAIN_SHIP_QTY = reader("ODR_QTY").ToString
            End If
            reader.Close()
            If check_flg = "1" Or check_flg = "2" Then
                Dim sql_get_picking As String = "select * from sup_scan_pick_detail where wi = '" & result_get_CUST_ODR_NO(g_index) & "' and item_cd = '" & result_get_ITEM_NO(g_index) & "' and Delivery_date = '" & delivery_date(g_index) & "'"
                Dim cmd_get_data As SqlCommand = New SqlCommand(sql_get_picking, myConn)
                reader = cmd_get_data.ExecuteReader()
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
                    F_com_flg.Add(reader.Item(12))
                    F_tag_remain_qty.Add(reader.Item(13))
                    F_Create_Date.Add(reader.Item(14))
                    F_Create_By.Add(reader.Item(15))
                    F_Line_cd.Add(reader.Item(17))
                    F_DELIVERY_DATE.Add(reader.Item(19))
                    F_id_sup.Add(reader.Item(0))
                    F_menu_CANCLE.Add(reader.Item(16))
                    F_DEL_CANCLE.Add(reader.Item(18))
                    F_SLIP_CD_CANCLE.Add(reader.Item(19))
                    F_Box_control_CANCLE.Add(reader.Item(20))
                    count += 1
                    count_arr_fw = count_arr_fw + 1
                Loop
                reader.Close()
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
                    Dim id_sup As String = F_id_sup(num)
                    Dim menu As String = F_menu_CANCLE(num)
                    Dim DEL As String = F_DEL_CANCLE(num)
                    Dim SLIP_CD_table As String = F_SLIP_CD_CANCLE(num)
                    Dim delivery_date As String = F_DELIVERY_DATE(num)
                    Dim box_control As String = F_Box_control_CANCLE(num)
                    M_SLIP_CD = SLIP_CD(g_index)
                    num += 1
                    log_cancle_fg(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, Line_cd, delivery_date, id_sup, menu, DEL, SLIP_CD_table, box_control)
                    return_flg(tag_readed, scan_qty)
                Next

                Dim str_update_qty = "EXEC [dbo].[cancle_pick_fg] @REMAIN_SHIP_QTY = '" & REMAIN_SHIP_QTY & "'  , @ITEM_CD = '" & result_get_ITEM_NO(g_index) & "'  , @DELIVERY_DATE = '" & delivery_date(g_index) & "' , @TRIP = '" & Trip & "' , @COM_FLG= '0'"
                Dim cmd_update As SqlCommand = New SqlCommand(str_update_qty, myconn_fa)
                reader = cmd_update.ExecuteReader()
                reader.Close()
                Dim str_update_flg = "update sup_scan_pick_detail set  com_flg = '9' where SLIP_CD = '" & M_SLIP_CD & "'"
                Dim cmd_update_flg As SqlCommand = New SqlCommand(str_update_flg, myConn)
                reader = cmd_update_flg.ExecuteReader()
                reader.Close()
                MsgBox("MOVE OK")
                p_show_confrim.Visible = False
            End If
            load_combobox1()
        Catch ex As Exception

        End Try
        'Try
        'Dim sql_get_picking As String = "select * from sup_scan_pick_detail where wi = '" & result_get_CUST_ODR_NO(g_index) & "' and item_cd = '" & result_get_ITEM_NO(g_index) & "' and Delivery_date = '" & delivery_date(g_index) & "'"
        'Dim cmd_get_data As SqlCommand = New SqlCommand(sql_get_picking, myConn)
        'reader = cmd_get_data.ExecuteReader()
        'Do While reader.Read()
        'F_wi.Add(reader.Item(1))
        'F_item_cd.Add(reader.Item(2))
        'F_scan_qty.Add(reader.Item(3))
        'F_scan_lot.Add(reader.Item(4))
        'F_updated_seq.Add(reader.Item(11))
        'F_com_flg.Add(reader.Item(12))
        'F_Line_cd.Add(reader.Item(17))
        'F_DELIVERY_DATE.Add(reader.Item(18))
        'Loop
        'reader.Close()
        'Dim delivery_date_check As String = "NO_DATA"
        'Dim num As Integer = 0
        'For Each key In F_wi
        'Dim wi As String = key
        'Dim item_cd As String = F_item_cd(num)
        'Dim scan_qty As String = F_scan_qty(num)
        'Dim lot As String = F_scan_lot(num)
        'Dim seq As String = F_updated_seq(num)
        'Dim com_flg As String = F_com_flg(num)
        'Dim line_cd As String = F_Line_cd(num)
        'delivery_date_check = F_DELIVERY_DATE(num)
        'Cut_stock(wi, item_cd, scan_qty, lot, seq, com_flg, line_cd, delivery_date_check)
        'num += 1
        'Next
        'Dim str_update_qty = "EXEC [dbo].[confirm_plan_fg] @ITEM_CD = '" & result_get_ITEM_NO(g_index) & "'  , @DELIVERY_DATE = '" & delivery_date(g_index) & "' , @TRIP = '" & Trip & "' ,@COM_FLG= '2'"
        'Dim cmd_update As SqlCommand = New SqlCommand(str_update_qty, myconn_fa)
        'reader = cmd_update.ExecuteReader()
        'reader.Close()
        'p_show_confrim.Visible = False
        'load_combobox1()
        'Catch ex As Exception
        'End Try
    End Sub
    Public Sub return_flg(ByVal tag_readed As String, ByVal scan_qty As String)
        Dim str_update_qr = "EXEC [dbo].[API_return_flg]  @scan_qty='" & scan_qty & "', @flg_status ='2' ,@read_qr = '" & tag_readed & "'"
        Dim cmd_update_qr As SqlCommand = New SqlCommand(str_update_qr, myconn_fa)
        reader = cmd_update_qr.ExecuteReader()
        reader.Close()
    End Sub
    Public Sub log_cancle_fg(ByVal count As String, ByVal F_wi As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal scan_lot As String, ByVal tag_typ As String, ByVal tag_readed As String, ByVal scan_emp As String, ByVal term_cd As String, ByVal updated_date As String, ByVal updated_by As String, ByVal updated_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal create_date As String, ByVal create_by As String, ByVal line_cd As String, ByVal delivery_date As String, ByVal id_sup As String, ByVal menu As String, ByVal DEL As String, ByVal SLIP_CD_TABLE As String, ByVal box_control As String)
        Try
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim sql_move = "INSERT INTO log_cancle_picking (id_sup , cancle_create_date , cancle_create_by , cancle_flg) VALUES ('" & id_sup & "' ,'" & date_now & "','" & Module1.A_USER_ID & "' ,'1')"
            'MsgBox(sql_move)
            Dim command2 As SqlCommand = New SqlCommand(sql_move, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("LOST" & vbNewLine & ex.Message, "FAILL")
        End Try

    End Sub
    Public Sub Cut_stock(ByVal wi As String, ByVal item_cd As String, ByVal scan_qty As String, ByVal lot As String, ByVal seq As String, ByVal com_flg As String, ByVal line_cd As String, ByVal delivery_date As String)
        Try
            ' MsgBox("11")
            Dim seq1 As String = "no_data"
            'MsgBox("seq length = " & seq.Length())
            If seq.Length() = "8" Then
                seq1 = seq.Substring(0, 3)
                'MsgBox("A")
            ElseIf seq.Length() = "5" Then
                seq1 = " "
                'MsgBox("B")
            Else
                seq1 = seq.Substring(8, 3)
                'MsgBox("C")
            End If
            'MsgBox("0002")
            'MsgBox("seq1 = " & seq1)
            Try
                If item_cd.Substring(11, 1) = "G" Then
                    item_cd = item_cd.Substring(0, 11)
                End If
            Catch ex As Exception

            End Try
            ' MsgBox("0003")
            Dim get_id_log = "select * from FA_TAG_FG where ITEM_CD = '" & item_cd & "' AND TAG_SEQ = '" & seq1 & "'AND LOT_NO = '" & lot & "' and KEY_UP = '" & seq & "' and LINE_CD = '" & line_cd & "'"
            'MsgBox("===>get_id_log" & get_id_log)
            Dim cmd_get As SqlCommand = New SqlCommand(get_id_log, myconn_fa)
            reader = cmd_get.ExecuteReader()
            Dim update_qty As Double = 0.0
            If reader.Read Then
                '    MsgBox("rrr")
                update_qty = CDbl(Val(reader("TAG_QTY").ToString())) - CDbl(Val(scan_qty))
            End If
            reader.Close()
            'MsgBox("-------------->>><<<")
            Dim FLG_STATUS As String = "0"
            If update_qty <= 0 Then
                update_qty = 0
                FLG_STATUS = "0"
            Else
                FLG_STATUS = "1"
            End If
            'MsgBox("----000")
            If seq.Length() = "8" Then
                seq1 = seq.Substring(0, 3)
            ElseIf seq.Length() = "5" Then
                seq1 = " "
            Else
                seq1 = seq.Substring(8, 3)
            End If
            Dim str_update_qty = "EXEC [dbo].[cut_stock_pick_fg] @qty = '" & update_qty & "'  , @flg_status = '" & FLG_STATUS & "' , @item_cd = '" & item_cd & "' , @seq = '" & seq1 & "' , @lot = '" & lot & "' , @KEY_UP = '" & seq & "' , @LINE_CD = '" & line_cd & "'"
            ' MsgBox("===>str" & str_update_qty)
            Dim cmd_update As SqlCommand = New SqlCommand(str_update_qty, myconn_fa)
            reader = cmd_update.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("FAILL Cut_stock" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub
    Private Sub p_show_confrim_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles p_show_confrim.GotFocus

    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.ParentChanged

    End Sub

    Private Sub Line_Emp_cd_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Line_Emp_cd.ParentChanged

    End Sub

    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged

    End Sub

    Private Sub days_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles days.SelectedIndexChanged
        load_combobox1()
    End Sub
    Public Sub delete_supscan_pick_detail()
        Dim del = "delete from sup_scan_pick_detail where SLIP_CD = '" & M_SLIP_CD & "'"
        Dim command2 As SqlCommand = New SqlCommand(del, myConn)
        reader = command2.ExecuteReader()
        reader.Close()
    End Sub
    Public Sub set_day()
        'days.Items.Add("-14")
        'days.Items.Add("-13")
        'days.Items.Add("-12")
        'days.Items.Add("-11")
        ' days.Items.Add("-10")
        ' days.Items.Add("-9")
        ' days.Items.Add("-8")
        ' days.Items.Add("-7")
        ' days.Items.Add("-6")
        ' days.Items.Add("-5")
        ' days.Items.Add("-4")
        ' days.Items.Add("-3")
        ' days.Items.Add("-2")
        'days.Items.Add("-1")
        days.Items.Add("1")
        days.Items.Add("2")
        days.Items.Add("3")
        days.Items.Add("4")
        days.Items.Add("5")
        days.Items.Add("6")
        days.Items.Add("7")
    End Sub


End Class