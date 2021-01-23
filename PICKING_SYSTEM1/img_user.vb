Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Imports PICKING_SYSTEM
Imports System.Drawing.Text
Imports System.Drawing
Imports System.Drawing.Bitmap
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Drawing.Imaging
Public Class img_user
    Dim myBitmap As System.Drawing.Bitmap
    Dim myGraphics As Graphics
    Dim myDestination As Rectangle
    Dim myConn As SqlConnection
    Dim reader As SqlDataReader
    Public given_code_pd As String
    Dim dat As String = String.Empty
    Public count_id As Integer
    Public max_count As Integer
    Public index As Integer
    Private Sub test_img_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            myConn.Open()
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status naja")
        Finally
            ' MsgBox(main.show_empToanofrm())
            name_user.Text = Module1.Fullname
        End Try
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub
    Public Sub New()
        InitializeComponent()
        Try
            If Api.DownloadImage("http://192.168.82.23/member/photo/" & Module1.user_id & ".jpg") IsNot Nothing Then
                PictureBox2.Image = Api.DownloadImage("http://192.168.82.23/member/photo/" & Module1.user_id & ".jpg")
                Dim result As String = Api.update_data("http://192.168.161.102/exp_api3party/Api_cut_stock_web_post/get_data_service?ITEM_CD=%221234%22")
                MsgBox("status = " & result)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Label1_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles name_user.ParentChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class