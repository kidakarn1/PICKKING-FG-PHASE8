Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Imports PICKING_SYSTEM
Public Class Select_PD
    'Dim myConn As SqlConnection
    Public myConn = "NOO"
    Dim reader As SqlDataReader
    Public given_code_pd As String
    Dim dat As String = String.Empty
    Public count_id As Integer
    Public max_count As Integer
    Public index As Integer



    Private Sub Select_PD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim connect_db = New connect()
            myConn = connect_db.conn()
        Finally

            query_sys_department()
            emp_name.Text = main.show_empToanofrm()
            se_code_user.Text = main.show_code_id_user()
            Dim v = main.show_code_id_user()
            Module1.user_id = v.Substring(7)
            'Me.emp_cd.Focus()
        End Try
        'name3.Text = frmMain.passToanofrm
    End Sub


    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.ParentChanged
        Label1.Text = "DATE " + ": " + DateTime.Now.ToString("dd-MM-yyyy")
    End Sub

    Private Sub emp_name_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles emp_name.ParentChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub se_code_user_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles se_code_user.ParentChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btb_pd_back.Click
        main.ml = 0
        main.Timer1.Enabled = False
        main.Panel2.Visible = False
        Me.Close() 'จาก close'
        main.Show()

    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ComboBox1.Text.Substring(0, 2)
        Module1.select_pd = ComboBox1.Text.Substring(2)
        next_page()
    End Sub
    Public Function give_value_code_pd() As String
        ' Line.Line_PD.Text = "PD : " + sel_pd
        Dim data = "PD : " + given_code_pd
        Return data
    End Function
   
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Panel1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel1.GotFocus

    End Sub
    Public Sub query_sys_department()
        Dim Line As Select_Line = New Select_Line()
        Line.Line_Emp_cd.Text = main.code_id_user ' 'แสดงในหน้า Select_line

        If main.pd_user <> "" Or main.pd_user IsNot Nothing Then

        End If

        PD_hidden_dep_id.Text = main.pd_user - 1 'ส่ง PD user ไปหน้า select_pd'
        'a.ComboBox1.SelectedIndex = 0'
        Dim strCommand As String = "select sec_name,dep_id from sys_department where  enable='1' order by dep_id"
        Dim command As SqlCommand = New SqlCommand(strCommand, myConn)
        reader = command.ExecuteReader()
        Do While reader.Read()
            ComboBox1.Items.Add(reader.Item(0))
        Loop
        reader.Close()
        ComboBox1.SelectedIndex = main.pd_user - 1 'ให่ค่า PD'
    End Sub
    Public Sub next_page()
        '   Try
        Dim sel_pd As String = ""
        sel_pd = ComboBox1.SelectedItem.ToString()
        Dim subSelpd As String
        subSelpd = sel_pd.Substring(2, 4)
        given_code_pd = sel_pd
        Module1.CODE_PD = sel_pd
        Dim Line As Select_Line = New Select_Line()
        Try
            Module1.data_combo = subSelpd
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status combo")
        Finally
            'ComboBox1.SelectedIndex = 0 'ให่ค่า PD'
            Line.PD3 = Me
            Line.PD3 = Me
            Line.Show()
            Me.Hide()
        End Try
        'Catch ex As Exception
        'MsgBox("Please select production!!!" & vbNewLine & ex.Message, 16, "Status BTN$")
        'MsgBox("Please select production!!!")
        ' End Try
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub



End Class