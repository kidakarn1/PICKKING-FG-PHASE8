Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Public Class Home
    Public scan_terminal_id = "PICK001"
    Dim myConn As SqlConnection
    Dim path As String
    Public Str As String
    Public count_emp_id As Integer
    Dim imagefile As String
    Public pd_of_user As String
    Public passToanofrm As String
    Public empToanofrm As String
    Public code_id_user As String
    Public pd_user As String
    Public valSel As Array
    Public strData(,) As String
    Public miscData() As Object = {}

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Dim m As main = New main()
        m.Show()
        Me.Hide()
    End Sub

    Private Sub Label3_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.ParentChanged

    End Sub
End Class