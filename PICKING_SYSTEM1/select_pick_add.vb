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
Public Class select_pick_add
    Public status_load_stock As String = "NO_DATA"
    Public myConn = "NOO"
    Dim reader As SqlDataReader
    Private Sub select_pick_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
            Dim connect_db = New connect()
            myConn = connect_db.conn()
            TextBox1.Text = "0"
            TextBox1.ForeColor = Color.Gray

            load_data()
        Finally
            TextBox2.Enabled = False
        End Try
    End Sub
    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Panel2.Visible = False
        Me.Close()
        main.Show()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        Dim get_date As String = DateTimePicker1.Format
        load_pd()
        load_line()
        load_wi()
        load_part()
        load_stock()
        'TextBox1.Focus()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        load_line()
        load_wi()
        load_part()
        load_stock()
        ' TextBox1.Focus()
    End Sub

    Private Sub Label8_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Sub load_pd()
        ComboBox1.Items.Clear()
        Dim get_pd As String = "select sec_name,dep_id from sys_department order by dep_id"
        Dim command As SqlCommand = New SqlCommand(get_pd, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read()
            ComboBox1.Items.Add(reader.Item(0))
        Loop
        reader.Close()
        ComboBox1.SelectedIndex = 0
    End Sub
    Public Sub load_line()
        ComboBox2.Items.Clear()
        Dim get_data_line As String = "SELECT DISTINCT LINE_CD from sup_work_plan_supply_dev where PD like '%" & ComboBox1.Text & "%' order by LINE_CD asc"
        Dim cmd_line As SqlCommand = New SqlCommand(get_data_line, myConn)
        reader = cmd_line.ExecuteReader()
        Dim num As Integer
        num = 0
        Dim status_line As Integer = 0
        Do While reader.Read()
            ComboBox2.Items.Add(reader.Item(0))
            num = num + 1
            status_line = 1
        Loop
        reader.Close()
        If status_line = 1 Then
            ComboBox2.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_wi()
        ComboBox3.Items.Clear()
        Module1.date_now_database = DateTimePicker1.Format
        Dim get_wi As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY , sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & DateTimePicker1.Text & "' AND AA.LINE_CD = '" & ComboBox2.Text & "' ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
        'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
        'MsgBox("get_wi = " & get_wi)
        Dim command_get_wi As SqlCommand = New SqlCommand(get_wi, myConn)
        reader = command_get_wi.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0
        Dim Status_wi As String = 0
        Do While reader.Read = True
            If reader.Item(2) = "1" Then
                Status_wi = 1
                ComboBox3.Items.Add(reader.Item(6))
            End If
        Loop
        reader.Close()
        If Status_wi = 1 Then
            ComboBox3.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_part()
        ComboBox4.Items.Clear()
        Module1.date_now_database = DateTimePicker1.Format
        Dim get_wi As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY  ,  sw.WORK_ODR_DLV_DATE AS d  , sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2,  sw.LOCATION_PART lo_part , sw.MODEL  MODEL, sw.ITEM_NAME ITEM_NAME , pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE = '" & DateTimePicker1.Text & "' AND AA.LINE_CD = '" & ComboBox2.Text & "' and AA.wi1 ='" & ComboBox3.Text & "'ORDER BY AA.wi1 ASC" 'แบบ list ออกมา ในวัน พน ในการpick แต่จะมีปัญหาคือ ถ้าถึงวันศุกร์ แล้วต้องจัดแผนวันจัน จะมองไม่เห็นข้อมูล จะมองเห็นแค่ วันเสาร์'
        'Dim strCommand As String = "SELECT AA.* FROM ( SELECT sw.PICK_QTY, sw.WORK_ODR_DLV_DATE AS d, sw.LVL AS LVL, sw.PICK_FLG AS PF, sw.item_cd AS item_cd, sw.LINE_CD, sw.wi AS wi1, pa.wi AS wi2, pa.del_flg, CASE WHEN (sw.wi = pa.wi) AND pa.del_flg = '1' THEN '9' ELSE '0' END AS FLG, sw.qty, CAST (sw.WORK_ODR_DLV_DATE AS DATE) AS DATE FROM sup_work_plan_supply_dev sw LEFT JOIN production_actual pa ON sw.WI = pa.WI ) AA WHERE AA.FLG <> '9' AND AA. DATE BETWEEN '2020-07-07' AND '2020-07-07' AND AA.LINE_CD = '" & sel_where & "' ORDER BY AA.wi1 ASC"
        'MsgBox("get_wi = " & get_wi)
        Dim command_get_wi As SqlCommand = New SqlCommand(get_wi, myConn)
        reader = command_get_wi.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0
        Dim Status_part As String = 0
        Do While reader.Read = True
            If reader.Item(2) = "2" Then
                ComboBox4.Items.Add(reader.Item(4))
                Status_part = 1
            End If
            Module1.A_LOCATION = reader("lo_part").ToString
            Module1.A_ITEM_NAME = reader("ITEM_NAME").ToString
            Module1.A_MODEL = reader("MODEL").ToString
        Loop
        reader.Close()
        If Status_part = 1 Then
            ComboBox4.SelectedIndex = 0
        End If
    End Sub
    Public Sub load_stock()
        Dim check_val As String = "SELECT count(id) as C_id from sup_frith_in_out where item_cd = '" & ComboBox4.Text & "' and com_flg <> 1"
        Dim c_check_val As SqlCommand = New SqlCommand(check_val, myConn)
        Dim C_val As Integer = 0
        reader = c_check_val.ExecuteReader()
        Do While reader.Read()
            C_val = reader("C_id").ToString
        Loop
        reader.Close()
        If C_val > 0 Then
            status_load_stock = "WEB_POST"
            GET_DATA_WEB_POST()
        Else
            status_load_stock = "FW"
            GET_DATA_FW()
        End If

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        load_wi()
        load_part()
        load_stock()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        load_part()
        load_stock()
        ' TextBox1.Focus()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        load_stock()
        ' TextBox1.Focus()
    End Sub
    Public Function GET_DATA_WEB_POST()
        Dim strCommand As String = "select sum(qty) as stock from sup_frith_in_out where item_cd='" & ComboBox4.Text & "' and com_flg = '0'" 'ใช้ที่บ้านแบบใหม่
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim Model As String = Nothing
        Dim totala_scan_qty As Double = 0.0
        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("stock").ToString))
            ' Part_Selected = reader("ITEM_CD").ToString
            'Part_Name = reader("ITEM_NAME").ToString
            QTY = totala_scan_qty
            ' wi = reader("WI").ToString
            'Lot_No = reader("LT").ToString
            ' Model = reader("MODEL").ToString
            'If Model <> Nothing Or Model <> "data_null" Then
            'Module1.M_Model = reader("MODEL").ToString
            'Else
            'Module1.M_Model = " - "
            'End If
            ' Module1.M_WI_STOP_SCAN = wi
            ' Module1.M_LINE_CD = code_line
            ' Module1.check_QTY = totala_scan_qty
            ' Module1.past_name = Part_Name
            'Module1.locations = lo
            ' Module1.M_LOT = reader("LT").ToString
        Loop
        reader.Close()

        TextBox2.Text = QTY
        Return 0
    End Function
    Public Sub GET_DATA_FW()
        Dim strCommand As String = "select sum(fa_total) -sum(fa_use) as stock from sup_frith_in_out_fa where fa_item_cd='" & ComboBox4.Text & "' and fa_com_flg = '0'"
        ' Dim strCommand As String = "SELECT item_cd, item_name , qty , location_part , wi ,MODEL , PICK_QTY  FROM sup_work_plan_supply_dev WHERE line_cd  = '" & code_line & "' AND item_cd = '" & code_part & "'AND wi  = '" & code_wi & "' AND (ps_unit_numerator <> '' AND location_part <> '') AND pick_flg != 1 AND WORK_ODR_DLV_DATE BETWEEN '2020-06-24' AND '2020-06-24' ORDER BY wi ASC"

        Dim Model As String = Nothing
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        'MsgBox(strCommand)
        reader = command.ExecuteReader()
        Dim totala_scan_qty As Double = 0.0

        Do While reader.Read()
            totala_scan_qty = CDbl(Val(reader("stock").ToString))
            '(Part_Selected = reader("ITEM_CD").ToString)
            '(Part_Name = reader("ITEM_NAME").ToString)
            '(lo = reader("LOCATION_PART").ToString)
            QTY = totala_scan_qty
            'wi = reader("WI").ToString
            'Model = reader("MODEL").ToString
            Module1.check_QTY = TextBox1.Text
            'Module1.past_name = Part_Name
            'Module1.locations = lo
            'Module1.M_LOT = "-"
            'If Model <> Nothing Or Model <> "data_null" Then
            ' Module1.M_Model = reader("MODEL").ToString
            ' Else
            ' Module1.M_Model = " - "
            'End If
            'Module1.M_WI_STOP_SCAN = wi
            'Module1.M_LINE_CD = code_line
        Loop
        ' MsgBox("LOT = " & Module1.M_LOT)
        reader.Close()
        TextBox2.Text = totala_scan_qty
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'part_detail_add.ListView2.Items.Clear()
        Dim part_de_ADD As part_detail_add = New part_detail_add()
        '        Panel2.Visible = False
        '       Timer1.Enabled = False
        '      ml = 0
        Module1.A_PD = ComboBox1.Text
        Module1.A_LINE = ComboBox2.Text
        Module1.A_WI = ComboBox3.Text
        Module1.A_PAST_NO = ComboBox4.Text
        Module1.A_PAST_QTY = TextBox1.Text
        Module1.check_QTY = TextBox1.Text
        Dim tb1 As Double = 0.0
        Dim tb2 As Double = 0.0
        tb1 = CDbl(Val(TextBox1.Text))
        tb2 = CDbl(Val(TextBox2.Text))
        If tb1 <= tb2 And tb1 <> 0 Then
            If status_load_stock = "WEB_POST" Then
                get_data_qty_po_lot(ComboBox2.Text, ComboBox4.Text, ComboBox3.Text)
            ElseIf status_load_stock = "FW" Then
                get_data_qty_po_lot_fa(ComboBox2.Text, ComboBox4.Text, ComboBox3.Text, TextBox1.Text)
            End If
            part_de_ADD.PD_ADD_PART = Me
            part_de_ADD.Show()
            Me.Hide()
        ElseIf tb1 <= "0" Or tb1 <= 0 Or tb1 <= 0.0 Then
            MsgBox("Please enter the QTY.")
        Else
            MsgBox("Not enough to pick add.")
        End If
    End Sub
    Public Function get_data_qty_po_lot(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String)
        Module1.M_CHECK_TYPE = "WEB_POST"
        Dim strCommand As String = "SELECT sd.*,F_I_O.LT ,F_I_O.QTY_OF_LOT , F_I_O.PO  FROM ( SELECT ss_f.item_cd, ss_f.com_flg,ss_f.LOT_RECEIVE as LT , ss_f.qty as QTY_OF_LOT ,ss_f.PUCH_ODR_CD as PO  , SUBSTRING (LOT_RECEIVE, 6, 8) AS date_lot FROM sup_frith_in_out ss_f WHERE ss_f.com_flg <> 1 ) F_I_O INNER JOIN sup_work_plan_supply_dev sd ON F_I_O.item_cd = sd.ITEM_CD WHERE sd.line_cd = '" & code_line & "' AND sd.item_cd = '" & code_part & "' AND sd.wi = '" & code_wi & "' AND ( ps_unit_numerator <> '' AND sd.location_part <> '' )  AND sd.WORK_ODR_DLV_DATE  = '" & DateTimePicker1.Text & "' ORDER BY sd.wi, F_I_O.date_lot ASC" 'ใช้ที่บ้านแบบใหม่
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim number As Integer = 1
        Dim RESULT_QTY As Integer = Module1.check_QTY
        Dim Lot_plus_qty As Integer = 0
        'MsgBox("Module1.check_QTY = " & Module1.check_QTY)
        'MsgBox("strCommand = " & strCommand)
        Dim lot_qty As String = "0"
        Dim N_lot_qty As Integer = 0
        Dim totala_scan_qty As Double = 0.0
        Do While reader.Read()
            totala_scan_qty = TextBox1.Text 'CDbl(Val(reader("qty").ToString)) - CDbl(Val(reader("PICK_QTY").ToString))
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
        Return 0
    End Function

    Public Function get_data_qty_po_lot_fa(ByVal code_line As String, ByVal code_part As String, ByVal code_wi As String, ByVal QTY As String)
        Module1.M_CHECK_TYPE = "FW"
        MsgBox(TextBox1.Text)
        Dim strCommand As String = "SELECT B.fa_lot AS LT, SUM (B.fa_total) AS fa_total, SUM (B.fa_use) AS fa_use FROM sup_frith_in_out_fa B WHERE B.fa_item_cd = '" & code_part & "' AND B.fa_com_flg = '0' GROUP BY B.fa_lot, B.fa_item_cd" 'แบบ ไม่ check wi FW
        'Dim strCommand As String = "SELECT B.fa_lot AS LT, SUM (B.fa_total) AS fa_total, SUM (B.fa_use) AS fa_use FROM sup_frith_in_out_fa B WHERE B.fa_item_cd = '" & code_part & "' AND B.fa_com_flg = '0' GROUP BY B.fa_lot, B.fa_item_cd" 'แบบ check wi FW
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Dim number As Integer = 1
        Dim RESULT_QTY As Integer = TextBox1.Text
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
            If CDbl(Val(TextBox1.Text)) <= CDbl(Val(Lot_plus_qty.ToString)) Then 'check ว่า QTY ที่ จะไป pick กับ QTY ของ LOT คือ ถ้า QTY ที่จะเอาน้อยกว่าหรือ= QTY ใน LOT ก็เอาไปได้เลบ'
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
            If CDbl(Val(TextBox1.Text)) > CDbl(Val(Lot_plus_qty.ToString)) Then
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
    Public Sub load_data()
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        Dim get_date As String = DateTimePicker1.Format
        load_pd()
        load_line()
        load_wi()
        load_part()
        load_stock()
    End Sub
    Private Sub Label11_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.ParentChanged

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        TextBox1.Text = ""
        TextBox1.ForeColor = Color.Black
    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If TextBox1.Text = "" Then
            TextBox1.Text = "0"
        End If

    End Sub
End Class