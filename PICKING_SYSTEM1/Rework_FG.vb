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
Public Class Rework_FG
    Dim reader As SqlDataReader
    Public myConn = "NOO"
    Public myconn_fa = "NO"
    Dim arr_qty_of_lot As ArrayList = New ArrayList()
    Dim arr_box As ArrayList = New ArrayList()
    Dim arr_lot As ArrayList = New ArrayList()
    Dim arr_remain_qty As ArrayList = New ArrayList()
    Dim F_Line_cd As ArrayList = New ArrayList()
    Dim F_delivery_date As ArrayList = New ArrayList()
    Dim arr_up_id As ArrayList = New ArrayList()
    Dim F_wi As ArrayList = New ArrayList()
    Dim F_DEL As ArrayList = New ArrayList()
    Dim F_menu As ArrayList = New ArrayList()
    Dim F_id_sup As ArrayList = New ArrayList()
    Dim F_item_cd As ArrayList = New ArrayList()
    Dim F_scan_qty As ArrayList = New ArrayList()
    Dim F_scan_lot As ArrayList = New ArrayList()
    Dim F_tag_typ As ArrayList = New ArrayList()
    Dim F_tag_readed As ArrayList = New ArrayList()
    Dim F_scan_emp As ArrayList = New ArrayList()
    Dim F_term_cd As ArrayList = New ArrayList()
    Dim F_updated_date As ArrayList = New ArrayList()
    Dim F_updated_by As ArrayList = New ArrayList()
    Dim F_updated_seq As ArrayList = New ArrayList()
    Dim F_com_flg As ArrayList = New ArrayList()
    Dim F_tag_remain_qty As ArrayList = New ArrayList()
    Dim F_Create_Date As ArrayList = New ArrayList()
    Dim F_Create_By As ArrayList = New ArrayList()
    Dim F_SLIP_CD As ArrayList = New ArrayList()
    Dim F_Box_control As ArrayList = New ArrayList()
    Dim count_arr_fw As Integer = 0
    Dim M_SLIP_CD As String = "NO_DATA"
    Public QR_CODE As String = ""
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel_confirm.Visible = True
    End Sub

    Private Sub TextBox1_Text_key_down(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub


    Private Sub scan_qr_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles scan_qr.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                Dim instruction = scan_qr.Text.Substring(7, 15)
                Dim arr_data = scan_qr.Text.Split(" ")
                inst.Text = instruction
                part_no.Text = arr_data(3)
                QTY.Text = arr_data(18)
        End Select
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Module1.RE_FG_SLIP_CD = inst.Text
        scan_qr.Text = ""
        inst.Text = ""
        part_no.Text = ""
        QTY.Text = ""

        Panel_select.Show()
    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub

    Private Sub Rework_FG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
            Dim connect_db = New connect()
            myConn = connect_db.conn()
            myconn_fa = connect_db.conn_fa()
        Finally
            Panel_confirm.Visible = False
            select_menu()
            scan_qr.Focus()
        End Try
    End Sub

    Private Sub scan_qr_KeyDown_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)

    End Sub
    Public Sub select_menu()
        Dim sql As String = "select * from category_rework where menu_rework_flg = '1'"
        Dim command As SqlCommand = New SqlCommand(sql, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read()
            cat_rework.Items.Add(reader("menu_rework_name").ToString())
        Loop
        reader.Close()
        cat_rework.SelectedIndex = 0
    End Sub

    Private Sub menu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cat_rework.SelectedIndexChanged

    End Sub

    Private Sub scan_qr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles scan_qr.TextChanged

    End Sub

    Private Sub btn_CANCEL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CANCEL.Click
        scan_qr.Text = ""
        inst.Text = ""
        part_no.Text = ""
        QTY.Text = ""
        scan_qr.Focus()
        Panel_confirm.Visible = False
    End Sub

    Private Sub btn_confirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_confirm.Click
        Try
            QR_CODE = scan_qr.Text
            Dim arr_QR_CODE = QR_CODE.Split(" ")
            Dim SLIP_CD = arr_QR_CODE(1)
            Dim count As Integer = 0
            Dim sql = "select * from sup_scan_pick_detail where SLIP_CD = '" & SLIP_CD & "' and com_flg = '1'"
            Dim sql_move = ""
            Dim command As SqlCommand = New SqlCommand(sql, myConn)
            reader = command.ExecuteReader()
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
                F_delivery_date.Add(reader.Item(19))
                F_id_sup.Add(reader.Item(0))
                F_menu.Add(reader.Item(16))
                F_DEL.Add(reader.Item(18))
                F_SLIP_CD.Add(reader.Item(19))
                F_Box_control.Add(reader.Item(20))
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
                Dim menu As String = F_menu(num)
                Dim DEL As String = F_DEL(num)
                Dim SLIP_CD_table As String = F_SLIP_CD(num)
                Dim delivery_date As String = F_delivery_date(num)
                Dim box_control As String = F_Box_control(num)
                M_SLIP_CD = SLIP_CD
                num += 1
                log_rework_fg(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, Line_cd, delivery_date, id_sup, menu, DEL, SLIP_CD_table)
                return_flg(tag_readed, scan_qty)
            Next
            Dim str_update_flg = "update sup_scan_pick_detail set  com_flg = '8' where SLIP_CD = '" & M_SLIP_CD & "'"
            Dim cmd_update_flg As SqlCommand = New SqlCommand(str_update_flg, myConn)
            reader = cmd_update_flg.ExecuteReader()
            reader.Close()
            'delete_supscan_pick_detail()
            MsgBox("SUCCESS")
            Panel_confirm.Visible = False
            scan_qr.Text = ""
            inst.Text = ""
            part_no.Text = ""
            QTY.Text = ""
            scan_qr.Focus()
        Catch ex As Exception
            MsgBox("STATUS===>FAIL" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub
    Public Sub log_rework_fg(ByVal count As String, ByVal F_wi As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal scan_lot As String, ByVal tag_typ As String, ByVal tag_readed As String, ByVal scan_emp As String, ByVal term_cd As String, ByVal updated_date As String, ByVal updated_by As String, ByVal updated_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal create_date As String, ByVal create_by As String, ByVal line_cd As String, ByVal delivery_date As String, ByVal id_sup As String, ByVal menu As String, ByVal DEL As String, ByVal SLIP_CD_TABLE As String)
        Try
            Dim arr_QR_CODE = QR_CODE.Split(" ")
            Dim SLIP_CD = arr_QR_CODE(1)
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim cat_rework_pk As String = CDbl(Val(cat_rework.SelectedIndex)) + 1
            Dim sql_move = "INSERT INTO log_rework_picking (id_sup ,  Rework_create_date , Rework_create_by , Rework_flg , Cat_rework) VALUES ('" & id_sup & "', '" & date_now & "', '" & Module1.A_USER_ID & "' , '1' , '" & cat_rework_pk & "')"
            Dim command2 As SqlCommand = New SqlCommand(sql_move, myConn)
            reader = command2.ExecuteReader()
            reader.Close()
        Catch ex As Exception
            MsgBox("LOST" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub
    Public Sub return_flg(ByVal tag_readed As String, ByVal scan_qty As String)
        Dim str_update_qr = "EXEC [dbo].[API_return_flg]  @scan_qty='" & scan_qty & "', @flg_status ='8' ,@read_qr = '" & tag_readed & "'"
        Dim cmd_update_qr As SqlCommand = New SqlCommand(str_update_qr, myconn_fa)
        reader = cmd_update_qr.ExecuteReader()
        reader.Close()
    End Sub
    Public Sub delete_supscan_pick_detail()
        Dim del = "delete from sup_scan_pick_detail where SLIP_CD = '" & M_SLIP_CD & "'"
        Dim command2 As SqlCommand = New SqlCommand(del, myConn)
        reader = command2.ExecuteReader()
        reader.Close()
    End Sub
    Private Sub Panel_confirm_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_confirm.GotFocus

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Panel_select.Hide()
        scan_qr.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Detail_Rework_fg = New Detail_rework_fg()
        Detail_Rework_fg.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Hide()
        main.Show()
    End Sub

    Private Sub Panel_select_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel_select.GotFocus

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class