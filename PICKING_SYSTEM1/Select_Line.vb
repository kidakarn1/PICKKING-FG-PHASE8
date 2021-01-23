Imports System.Data.SqlClient
Imports PICKING_SYSTEM.Page_projects
Imports System.Data
Public Class Select_Line
    Public status As Integer = 0
    Public ml As Integer = 0
    Public count_time As Integer = 0
    'Dim myConn As SqlConnection
    Public myConn = "NOO"
    Dim x As ListViewItem
    Dim sel_itemSpa As String = "                        "
    Dim sel_where As String = ""
    Public Part_Selected As String = ""
    Public Part_Name As String = ""
    Public lo As String = ""
    Public QTY As String = ""
    Public Lot_No As String = ""
    Public PD3 As Select_PD
    Dim wi As String = ""
    Public part As part_detail
    Dim arr_LVL As ArrayList = New ArrayList()
    Dim arr_QTY As ArrayList = New ArrayList()
    Dim arr_com_flg As ArrayList = New ArrayList()
    Public check_action As Integer = 0
    Public index_list_view As Integer = 0
    Public g_index As Integer = 0
    Dim reader As SqlDataReader
    Dim dat As String = String.Empty

    Private Sub selLine_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
LOOP_MAIN_OPEN:
            Dim connect_db = New connect()
            myConn = connect_db.conn()
        Finally
            'MsgBox("====>>>")
            Module1.M_QTY_LOT_ALL = 0
            combobox_line()

            load_data()
            Line_PD.Text = Module1.CODE_PD
            Line_Emp_cd.Text = main.show_code_id_user()
            btn_ok.Visible = False
            check_action = 2
            Panel3.Visible = False
            Panel8.Visible = False
            Panel9.Visible = False
        End Try
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Line_list_view.Items.Clear()
        sel_where = ComboBox1.SelectedItem.ToString()
        Module1.line = sel_where
        Dim status_combobox1 = 0
        Try
            Dim t As String = "select  pd.id, pd.PD , pd.Time_pd , DATEADD (hour , pd.Time_pd , GETDATE()) AS DateAdd from sys_time_pd pd where pd.PD LIKE '%" & Module1.select_pd & "%'"
            Dim command_t As SqlCommand = New SqlCommand(t, myConn)
            reader = command_t.ExecuteReader()
            Do While reader.Read = True
                Module1.time_scan = reader("DateAdd").ToString()
            Loop
            reader.Close()
            ' MsgBox(Module1.time_scan)
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.LVL AS LVL, sw.PICK_FLG as PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN CAST (GETDATE() AS DATE) AND '" & Module1.time_scan & "' AND AA.LINE_CD = '" & sel_where & "'  ORDER BY AA.wi1 ASC" 'ใช้ บริษัท and AA.PF=0' 
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN DATEADD( DAY, 1, CAST (GETDATE() AS DATE)) AND DATEADD( DAY, 4, CAST (GETDATE() AS DATE)) AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
            Dim str_date_now2 As String = "select GETDATE() as d "
            'MsgBox("00001")
            Dim cmd_str2 As SqlCommand = New SqlCommand(str_date_now2, myConn)
            'MsgBox("00002")
            reader = cmd_str2.ExecuteReader()
            ' MsgBox("00003")
            Dim date_now As String = "NOOOOO"
            Dim date_now_database As String = "NNNOO"
            Do While reader.Read = True
                'MsgBox("-----")
                date_now = reader.Item(0)
                'MsgBox(date_now)
                '  MsgBox("00004")
                Dim sp_date_now = date_now.Split(" ")
                '   MsgBox("00005")
                date_now_database = sp_date_now(0)
                'MsgBox(date_now_database)
            Loop
            reader.Close()



            Dim srt_get_day_work = "select * from EXP_WORK_DAY where work_date = '" & date_now_database & "' and enable ='1'"
            Dim cmd_work_day As SqlCommand = New SqlCommand(srt_get_day_work, myConn)
            'MsgBox("01")
            reader = cmd_work_day.ExecuteReader()
            Dim wd As String = "NO"
            'MsgBox("02")
            Dim date_now_get As String = "NODATA"
            'MsgBox("03")
            If reader.Read() Then
                'MsgBox("04")
                wd = reader("number_work").ToString()
                wd += 1
                'MsgBox("05")
                date_now_get = reader("work_date").ToString()
                'MsgBox("06")
                reader.Close()
                Dim wok_show_data = "select * from EXP_WORK_DAY where work_id = '" & wd & "' and enable ='1' "
                'MsgBox(wok_show_data)
                'MsgBox("07")
                Dim cmd_work_show As SqlCommand = New SqlCommand(wok_show_data, myConn)
                'MsgBox("08")
                reader = cmd_work_show.ExecuteReader()
                'MsgBox("09")
                If reader.Read() Then
                    date_now_get = reader("work_date").ToString()
                Else
                    MsgBox("ELSE NAJA")
                End If
                'MsgBox(date_now_get)
                reader.Close()
            Else
                reader.Close()
                date_now_get = date_now_database
            End If
            Module1.date_now_database = date_now_get
            Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY , sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & date_now_get & "' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
            'MsgBox("strCommand = " & strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            Dim num As Integer
            num = 0
            'Dim LI As New ListViewItem 'new obj ''
            arr_LVL = New ArrayList()
            arr_com_flg = New ArrayList()
            Dim totala_scan_qty As Double = 0.0
            Do While reader.Read = True
                totala_scan_qty = CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
                x = New ListViewItem(reader("item_cd").ToString)
                x.SubItems.Add(reader("wi1").ToString)
                x.SubItems.Add(totala_scan_qty)
                Line_list_view.Items.Add(x)
                status_combobox1 = 1
                If reader("LVL").ToString = "1" Then
                    Line_list_view.Items(num).BackColor = Color.FromArgb(222, 140, 236)
                    arr_com_flg.Add(reader("PF").ToString)
                    arr_LVL.Add(reader("LVL").ToString)
                    'Line_list_view.Font = New Font(Line_list_view.Font.Size, Line_list_view.Font.Size, FontStyle.Bold)
                    btn_ok.Visible = False
                Else
                    arr_com_flg.Add(reader("PF").ToString)
                    arr_LVL.Add(reader("LVL").ToString)
                End If
                If reader("PF").ToString = "1" Then
                    Line_list_view.Items(num).BackColor = Color.FromArgb(103, 255, 103)
                    btn_ok.Visible = False
                ElseIf reader("PF").ToString = "2" Then
                    Line_list_view.Items(num).BackColor = Color.Yellow
                    btn_ok.Visible = True
                End If
                arr_QTY.Add(totala_scan_qty)
                num = num + 1
            Loop
            reader.Close()
            'scan_pick.line_cd.Text = sel_where '
            'part_detail.line_cd.Text = sel_where '
        Catch ex As Exception
            reader.Close()
            MsgBox("ERROR FUNCTION load_data" & vbNewLine & ex.Message, 16, "Status")
        Finally
            'MsgBox("OK")
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Line_list_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Line_list_view_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Line_list_view.SelectedIndexChanged
        If IsNothing(Me.Line_list_view.FocusedItem) Then

        ElseIf Line_list_view.FocusedItem.Index >= 0 Then
            If Line_list_view.Items.Count > 0 Then
                Dim index As Integer = Line_list_view.FocusedItem.Index
                g_index = index
                'MsgBox(" index = " & index)
                'MsgBox(" arr_LVL = " & arr_LVL(index))
                'MsgBox(" arr_com_flg = " & arr_com_flg(index))
                If arr_LVL(index) <> "1" And (arr_com_flg(index) = "0" Or arr_com_flg(index) = "2") Then
                    btn_ok.Visible = True
                End If
                If arr_LVL(index) = "1" Or arr_com_flg(index) = "1" Then
                    btn_ok.Visible = False
                End If
            End If
        Else

            MessageBox.Show("An Error has halted thid process")
        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ml = 0
        Timer1.Enabled = True
        Panel8.Visible = True
        loader()
        Application.DoEvents()
        Line_list_view.Items.Clear()
        Module1.M_CHECK_LOT_COUNT_FW = New ArrayList()
        Dim page As Page_projects = New Page_projects()
        Timer1.Enabled = False
        Panel8.Visible = False
        page.PD.Show()
        Me.Close()
    End Sub
    Public Sub loader()
        Panel8.Visible = True
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ok.Click
        Timer1.Enabled = True
        loader()
        Application.DoEvents()
        Try
            query_detail_location()
            If Module1.check_query = 1 Then
                Dim p_scan As scan_location = New scan_location()
                p_scan.PD4 = Me
                Panel8.Visible = False
                Timer1.Enabled = False
                p_scan.Show()
                Me.Hide()
            End If
        Catch ex As Exception
            MsgBox("Please select Part")
        Finally
        End Try
    End Sub
    Public Function get_index()
        Return Line_list_view.FocusedItem.Index
    End Function

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Line_PD_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Line_PD.ParentChanged

    End Sub

    Private Sub Line_Emp_cd_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Line_Emp_cd.ParentChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Function given_code_line() As String
        Dim data = "LINE: " + sel_where
        Return data
    End Function

    Public Function get_pd() As String
        Return PD3.give_value_code_pd
    End Function

    Public Function query_detail_location() As String
        Try
            Dim code_line = sel_where
            Dim code_part = Line_list_view.FocusedItem.SubItems(0).Text
            Dim code_wi = Line_list_view.FocusedItem.SubItems(1).Text
            Module1.wi = code_wi

            Part_Selected = ""
            Part_Name = ""
            lo = ""
            QTY = ""
            wi = ""
            Lot_No = ""

            'Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE BETWEEN CAST(GETDATE() AS DATE) AND DATEADD(DAY, 4, CAST(GETDATE() AS DATE)) ORDER BY wi ASC" 'แบบเก่าที่ใช้ บริษัท
            ' Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' ORDER BY wi ASC" 'แบบใช้ที่บ้าน
            'MsgBox(strCommand)
            '  Dim strCommand As String = "SELECT sd.* FROM ( SELECT ss_f.item_cd, ss_f.com_flg, SUBSTRING (LOT_RECEIVE, 6, 8) AS date_lot FROM sup_frith_in_out ss_f WHERE ss_f.com_flg <> 1 ) F_I_O INNER JOIN sup_work_plan_supply_dev sd ON F_I_O.item_cd = sd.ITEM_CD WHERE sd.line_cd = '" & code_line & "' AND sd.item_cd = '" & code_part & "' AND sd.wi = '" & code_wi & "' AND ( ps_unit_numerator <> '' AND sd.location_part <> '' ) AND sd.pick_flg != 1 AND sd.WORK_ODR_DLV_DATE BETWEEN CAST (GETDATE() AS DATE) AND DATEADD( DAY, 4, CAST (GETDATE() AS DATE)) ORDER BY sd.wi, F_I_O.date_lot ASC" 'ทำการแก้ไขอยู่'

            Dim check_val As String = "SELECT count(id) as C_id from sup_frith_in_out where item_cd = '" & code_part & "' and com_flg <> 1"
            Dim c_check_val As SqlCommand = New SqlCommand(check_val, myConn)
            Dim C_val As Integer = 0

            reader = c_check_val.ExecuteReader()
            Do While reader.Read()
                C_val = reader("C_id").ToString
            Loop
            reader.Close()
            'MsgBox("C_val = " & C_val)
            If C_val > 0 Then
                'MsgBox("data PART TAG WEB POST")
                GET_DATA_WEB_POST(code_line, code_part, code_wi)
            Else
                'MsgBox("data PART TAG FW")
                Try
                    GET_DATA_FW(code_line, code_part, code_wi)
                Catch ex As Exception
                    MsgBox("ERROR " & vbNewLine & ex.Message, 16, "Status in")
                End Try

            End If

            Module1.check_query = 1
        Catch ex As Exception
            reader.Close()
            Module1.check_query = 0
            MsgBox("Please select Part" & vbNewLine & ex.Message, 16, "Status in")
        End Try
        M_Part_Selected = Part_Selected
        Return 0
    End Function

    Public Function GET_DATA_WEB_POST(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String)

        Dim strCommand As String = "SELECT sd.*,F_I_O.LT ,F_I_O.QTY_OF_LOT , F_I_O.PO  FROM ( SELECT ss_f.item_cd, ss_f.com_flg,ss_f.LOT_RECEIVE as LT , ss_f.qty as QTY_OF_LOT ,ss_f.PUCH_ODR_CD as PO  , SUBSTRING (LOT_RECEIVE, 6, 8) AS date_lot FROM sup_frith_in_out ss_f WHERE ss_f.com_flg <> 1 ) F_I_O INNER JOIN sup_work_plan_supply_dev sd ON F_I_O.item_cd = sd.ITEM_CD WHERE sd.line_cd = '" & code_line & "' AND sd.item_cd = '" & code_part & "' AND sd.wi = '" & code_wi & "' AND ( ps_unit_numerator <> '' AND sd.location_part <> '' ) AND sd.pick_flg != 1 AND sd.WORK_ODR_DLV_DATE  = '" & Module1.date_now_database & "' ORDER BY sd.wi, F_I_O.date_lot DESC" 'ใช้ที่บ้านแบบใหม่
        'Dim strCommand As String = "SELECT sd.*,F_I_O.LT ,F_I_O.QTY_OF_LOT , F_I_O.PO  FROM ( SELECT ss_f.item_cd, ss_f.com_flg,ss_f.LOT_RECEIVE as LT , ss_f.qty as QTY_OF_LOT ,ss_f.PUCH_ODR_CD as PO  , SUBSTRING (LOT_RECEIVE, 6, 8) AS date_lot FROM sup_frith_in_out ss_f WHERE ss_f.com_flg <> 1 ) F_I_O INNER JOIN sup_work_plan_supply_dev sd ON F_I_O.item_cd = sd.ITEM_CD WHERE sd.line_cd = '" & code_line & "' AND sd.item_cd = '" & code_part & "' AND sd.wi = '" & code_wi & "' AND ( ps_unit_numerator <> '' AND sd.location_part <> '' ) AND sd.pick_flg != 1 AND sd.WORK_ODR_DLV_DATE  BETWEEN '2020-07-07' AND '2020-07-07' ORDER BY sd.wi, F_I_O.date_lot DESC"
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        'MsgBox(strCommand)
        reader = command.ExecuteReader()
        'MsgBox(arr_QTY(g_index))
        Dim Model As String = Nothing
        Dim totala_scan_qty As Double = 0.0
        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
            Part_Selected = reader("ITEM_CD").ToString
            Part_Name = reader("ITEM_NAME").ToString
            lo = reader("LOCATION_PART").ToString
            QTY = totala_scan_qty
            wi = reader("WI").ToString
            Lot_No = reader("LT").ToString
            'MsgBox("data =" & reader("MODEL").ToString)
            Model = reader("MODEL").ToString
            If Model <> Nothing Or Model <> "data_null" Then
                Module1.M_Model = reader("MODEL").ToString
            Else
                Module1.M_Model = " - "
            End If
            Module1.M_WI_STOP_SCAN = wi
            Module1.M_LINE_CD = code_line
            Module1.check_QTY = totala_scan_qty
            Module1.past_name = Part_Name
            Module1.locations = lo
            Module1.M_LOT = reader("LT").ToString
        Loop
        reader.Close()

        get_data_qty_po_lot(code_line, code_part, code_wi)
        'MsgBox("LOT = " & Module1.M_LOT)
        reader.Close()
        Return 0
    End Function

    Public Sub GET_DATA_FW(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String)
        Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi ,MODEL , PICK_QTY  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE  = '" & Module1.date_now_database & "' ORDER BY wi ASC"
        ' Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi ,MODEL , PICK_QTY  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE BETWEEN '2020-06-24' AND '2020-06-24' ORDER BY wi ASC"

        Dim Model As String = Nothing
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        'MsgBox(strCommand)
        reader = command.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0

        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
            Part_Selected = reader("ITEM_CD").ToString
            Part_Name = reader("ITEM_NAME").ToString
            lo = reader("LOCATION_PART").ToString
            QTY = totala_scan_qty
            wi = reader("WI").ToString
            Model = reader("MODEL").ToString
            Module1.check_QTY = QTY
            Module1.past_name = Part_Name
            Module1.locations = lo
            Module1.M_LOT = "-"
            If Model <> Nothing Or Model <> "data_null" Then
                Module1.M_Model = reader("MODEL").ToString
            Else
                Module1.M_Model = " - "
            End If
            Module1.M_WI_STOP_SCAN = wi
            Module1.M_LINE_CD = code_line
        Loop
        ' MsgBox("LOT = " & Module1.M_LOT)
        reader.Close()
        Try
            Dim strCommand2 As String = "SELECT COUNT (fa_id) AS c, SUM (fa_total), SUM (fa_use), (SUM(fa_total) - SUM(fa_use)) AS data_total FROM sup_frith_in_out_fa WHERE fa_item_cd = '" & code_part & "' and fa_com_flg = '0'"
            'MsgBox("strCommand2 = " & strCommand2)
            Dim command2 As SqlCommand = New SqlCommand(strCommand2, myConn)
            reader = command2.ExecuteReader()
            Dim data_total As Integer = 0
            Dim total_qty As Integer = 0
            Do While reader.Read()
                If reader("c").ToString > 0 Then
                    Module1.M_QTY_LOT_ALL = reader("data_total").ToString
                End If
            Loop
            reader.Close()
            get_data_qty_po_lot_fa(code_line, code_part, code_wi, QTY, data_total)
        Catch ex As Exception
            reader.Close()
            MsgBox("ERROR Fail" & vbNewLine & ex.Message, 16, "Status in")
        End Try
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

    Public Function get_Part_Selected() As String
        Dim data = "Part_Selected : " + Part_Selected
        Return data
    End Function

    Public Function get_Part_Name() As String
        Dim data = "Part_Name : " + Part_Name
        Return data
    End Function

    Public Function get_Location() As String
        Dim data = "Location : " + lo
        Return data
    End Function

    Public Function get_QTY() As String
        Dim data = "QTY : " + QTY
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

    Private Sub Lb2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lb2.ParentChanged

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
            ComboBox1.Items.Add(reader.Item(0))
            num = num + 1
            status = 1
        Loop
        reader.Close()

        If status = 1 Then
            ComboBox1.SelectedIndex = 0
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
        Line_list_view.Items.Clear()
        sel_where = ComboBox1.SelectedItem.ToString()
        Module1.line = sel_where
        Dim status_combobox1 = 0
        Try
            Dim t As String = "select  pd.id, pd.PD , pd.Time_pd , DATEADD (hour , pd.Time_pd , GETDATE()) AS DateAdd from sys_time_pd pd where pd.PD LIKE '%" & Module1.select_pd & "%'"
            Dim command_t As SqlCommand = New SqlCommand(t, myConn)
            reader = command_t.ExecuteReader()
            Do While reader.Read = True
                Module1.time_scan = reader("DateAdd").ToString()
            Loop
            reader.Close()
            ' MsgBox(Module1.time_scan)
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.LVL AS LVL, sw.PICK_FLG as PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN CAST (GETDATE() AS DATE) AND '" & Module1.time_scan & "' AND AA.LINE_CD = '" & sel_where & "'  ORDER BY AA.wi1 ASC" 'ใช้ บริษัท and AA.PF=0' 
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN DATEADD( DAY, 1, CAST (GETDATE() AS DATE)) AND DATEADD( DAY, 4, CAST (GETDATE() AS DATE)) AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
            Dim str_date_now2 As String = "select GETDATE() as d "
            Dim cmd_str2 As SqlCommand = New SqlCommand(str_date_now2, myConn)
            'MsgBox("00002")
            reader = cmd_str2.ExecuteReader()
            ' MsgBox("00003")
            Dim date_now As String = "NOOOOO"
            Dim date_now_database As String = "NNNOO"
            Do While reader.Read = True
                'MsgBox("-----")
                date_now = reader.Item(0)
                'MsgBox(date_now)
                '  MsgBox("00004")
                Dim sp_date_now = date_now.Split(" ")
                '   MsgBox("00005")
                date_now_database = sp_date_now(0)
                'MsgBox(date_now_database)
            Loop
            reader.Close()

            Dim srt_get_day_work = "select * from EXP_WORK_DAY where work_date = '" & date_now_database & "' and enable ='1'"
            Dim cmd_work_day As SqlCommand = New SqlCommand(srt_get_day_work, myConn)

            'MsgBox("01")
            reader = cmd_work_day.ExecuteReader()
            Dim wd As String = "NO"
            'MsgBox("02")
            Dim date_now_get As String = "NODATA"
            'MsgBox("03")
            If reader.Read() Then
                'MsgBox("04")
                wd = reader("number_work").ToString()
                wd += 1
                'MsgBox("05")
                date_now_get = reader("work_date").ToString()
                'MsgBox("06")
                reader.Close()
                Dim wok_show_data = "select * from EXP_WORK_DAY where work_id = '" & wd & "' and enable ='1'"
                'MsgBox(wok_show_data)
                'MsgBox("07")
                Dim cmd_work_show As SqlCommand = New SqlCommand(wok_show_data, myConn)
                'MsgBox("08")
                reader = cmd_work_show.ExecuteReader()
                'MsgBox("09")
                If reader.Read() Then
                    date_now_get = reader("work_date").ToString()
                Else
                    MsgBox("ELSE NAJA")
                End If
                'MsgBox(date_now_get)
                reader.Close()
            Else
                reader.Close()
                date_now_get = date_now_database
            End If
            Module1.date_now_database = date_now_get
            Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY , sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & date_now_get & "' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
            'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
            'MsgBox("strCommand = " & strCommand)
            Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
            reader = command.ExecuteReader()
            Dim num As Integer
            num = 0
            'Dim LI As New ListViewItem 'new obj ''
            arr_LVL = New ArrayList()
            arr_com_flg = New ArrayList()
            Dim totala_scan_qty As Double = 0.0
            Do While reader.Read = True
                totala_scan_qty = CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
                x = New ListViewItem(reader("item_cd").ToString)
                x.SubItems.Add(reader("wi1").ToString)
                x.SubItems.Add(totala_scan_qty)
                Line_list_view.Items.Add(x)
                status_combobox1 = 1
                If reader("LVL").ToString = "1" Then
                    Line_list_view.Items(num).BackColor = Color.FromArgb(222, 140, 236)
                    arr_com_flg.Add(reader("PF").ToString)
                    arr_LVL.Add(reader("LVL").ToString)
                    'Line_list_view.Font = New Font(Line_list_view.Font.Size, Line_list_view.Font.Size, FontStyle.Bold)
                    btn_ok.Visible = False
                Else
                    arr_com_flg.Add(reader("PF").ToString)
                    arr_LVL.Add(reader("LVL").ToString)
                End If
                If reader("PF").ToString = "1" Then
                    Line_list_view.Items(num).BackColor = Color.FromArgb(103, 255, 103)
                    btn_ok.Visible = False
                ElseIf reader("PF").ToString = "2" Then
                    Line_list_view.Items(num).BackColor = Color.Yellow
                    btn_ok.Visible = True
                End If
                arr_QTY.Add(totala_scan_qty)
                num = num + 1
            Loop
            reader.Close()
            'scan_pick.line_cd.Text = sel_where '
            'part_detail.line_cd.Text = sel_where '
        Catch ex As Exception
            reader.Close()
            MsgBox("ERROR FUNCTION load_data" & vbNewLine & ex.Message, 16, "Status")
        Finally
            'MsgBox("OK")
        End Try
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        Panel9.Visible = True
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Panel9.Visible = False
    End Sub
    Public Function FIFO_FG()
        Return 0
    End Function
End Class