﻿Imports System.Runtime.InteropServices
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

Public Class Detail_rework_fg
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
    Dim check_data As ArrayList = New ArrayList()
    Dim TAG_QR As ArrayList = New ArrayList()
    Dim TAG_ID As ArrayList = New ArrayList()
    Dim count_arr_fw As Integer = 0
    Public M_SLIP_CD As String = ""
    Public QR_CODE As String = ""
    Public count_box As Double = 0.0
    Public count_qty As Double = 0.0
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Detail_rework_fg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            'myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            'myConn.Open()
            Dim connect_db = New connect()
            myConn = connect_db.conn()
            myconn_fa = connect_db.conn_fa()
            show_box.Text = 0
            show_qty.Text = 0
            select_menu()
            scan_qr_fg.Focus()
        Finally
        End Try
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
    Private Sub select_lot_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub show_qty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles show_qty.TextChanged

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Rework_FG.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If TAG_ID(0) = Not Nothing Then
                MsgBox("NOT SCAN")
            End If
            Dim count As Integer = 0
            Dim index As Integer = 0
            For Each key_main In TAG_ID
                Dim sql = "select * from FA_TAG_FG where IND = '" & TAG_ID(index) & "'"
                Dim command As SqlCommand = New SqlCommand(sql, myconn_fa)
                reader = command.ExecuteReader()
                Do While reader.Read()
                    F_scan_qty.Add(reader.Item(2))
                    F_tag_readed.Add(reader.Item(1))
                Loop
                reader.Close()
                Dim num As Integer = 0
                For Each key In F_tag_readed
                    Dim scan_qty As String = F_scan_qty(num)
                    Dim tag_readed As String = F_tag_readed(num)
                    num += 1
                    Try
                        '                        log_rework_fg(count, wi, item_cd, scan_qty, scan_lot, tag_typ, tag_readed, scan_emp, term_cd, updated_date, updated_by, updated_seq, com_flg_table, tag_remain_qty, Create_date, Create_By, Line_cd, delivery_date, id_sup, Menu, DEL, SLIP_CD_table)
                        Dim sql_get_date = "SELECT GETDATE() as date_now"
                        Dim cmd_get_date As SqlCommand = New SqlCommand(sql_get_date, myConn)
                        reader = cmd_get_date.ExecuteReader()
                        Dim date_now As String = ""
                        If reader.Read() Then
                            date_now = reader("date_now").ToString()
                        End If
                        reader.Close()
                        Dim insert_log_rework = "EXEC dbo.Log_rework_fg @qr_read = '" & tag_readed & "' , @location = '" & cat_rework.Text & "'  , @date_time = '" & date_now & "', @re_use = '" & Module1.A_USER_ID & "'"
                        Dim cmd_get_insert_log As SqlCommand = New SqlCommand(insert_log_rework, myconn_fa)
                        reader = cmd_get_insert_log.ExecuteReader()
                        reader.Close()
                        return_flg(tag_readed, scan_qty)
                    Catch ex As Exception
                        MsgBox("MOVE FALL" & vbNewLine & ex.Message)
                    End Try
                    ' MsgBox("1234")
                Next
                index = index + 1
            Next

            'delete_supscan_pick_detail()
            MsgBox("SUCCESS")
            Rework_FG.Show()
            Me.Close()
        Catch ex As Exception
            MsgBox("STATUS===>FAIL" & vbNewLine & ex.Message, "FAILL")
        End Try
    End Sub
    Public Sub log_rework_fg(ByVal count As String, ByVal F_wi As String, ByVal F_item_cd As String, ByVal scan_qty As String, ByVal scan_lot As String, ByVal tag_typ As String, ByVal tag_readed As String, ByVal scan_emp As String, ByVal term_cd As String, ByVal updated_date As String, ByVal updated_by As String, ByVal updated_seq As String, ByVal com_flg_table As String, ByVal tag_remain_qty As String, ByVal create_date As String, ByVal create_by As String, ByVal line_cd As String, ByVal delivery_date As String, ByVal id_sup As String, ByVal menu As String, ByVal DEL As String, ByVal SLIP_CD_TABLE As String)
        Try
            Dim time As DateTime = DateTime.Now
            Dim format As String = "yyyy-MM-dd HH:mm:ss"
            Dim date_now = time.ToString(format)
            Dim cat_rework_pk As String = CDbl(Val(Rework_FG.cat_rework.SelectedIndex)) + 1
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

    Private Sub combox_type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub scan_qr_fg_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles scan_qr_fg.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Enter
                Dim status As String = "NO_DATA"
                If scan_qr_fg.Text.Length = "103" Then
                    ' Dim sql As String = "select count(id) as c_id  , scan_qty , id , com_flg from sup_scan_pick_detail where tag_readed = '" & scan_qr_fg.Text & "'  group by scan_qty , id , com_flg"
                    Dim sql As String = "select count(IND) as c_id  , TAG_QTY , IND , FLG_STATUS from FA_TAG_FG where READ_QR = '" & scan_qr_fg.Text & "'  group by TAG_QTY , IND , FLG_STATUS"
                    Dim command As SqlCommand = New SqlCommand(sql, myconn_fa)
                    reader = command.ExecuteReader()
                    If reader.Read() Then
                        'If reader("FLG_STATUS").ToString() = "2" Then
                        For Each key In TAG_QR
                            If key = scan_qr_fg.Text Then
                                status = "1"
                                reader.Close()
                                GoTo alert_loop
                            End If
                        Next
                        TAG_QR.Add(scan_qr_fg.Text)
                        TAG_ID.Add(reader("IND").ToString())
                        count_box += 1
                        show_box.Text = count_box
                        count_qty = CDbl(Val(count_qty)) + CDbl(Val(reader("TAG_QTY").ToString()))
                        show_qty.Text = count_qty
                        ' reader.Close()
                        'Else
                        '    MsgBox("TAG นี้ ไม่ได้ TRANFER มา")
                        'End If
                    Else
                        MsgBox("QR นี้ ไม่มีของในสต๊อก")
                    End If
                    'End If
                    reader.Close()
                Else
                    MsgBox("NO TAG FG")
                End If
alert_loop:
                If status = "1" Then
                    MsgBox("สแกนซ้ำ")
                End If
                scan_qr_fg.Text = ""
                scan_qr_fg.Focus()
        End Select
    End Sub

    Private Sub menu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class