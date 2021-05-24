Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Forms.Form
Imports System
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Linq
Imports Newtonsoft
Public Class connect
    Public myConn As SqlConnection
    Public myConn_FA As SqlConnection
    Public Function conn()
next_connected:
        Dim myConn As SqlConnection
        Try
            Dim IP As String = "192.168.161.101"
            If Api.check_net() = True Then
                '  myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
                myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
                myConn.Open()
                Return myConn
            Else
                'MsgBox("NET หลุด กด ENT")
                GoTo next_connected
            End If

        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status in")
            GoTo next_connected
        End Try
        Return "SORRY CONNECT FAILL"
    End Function
    Public Function conn_fa()
next_connected_fa:
        Dim myConn_FA As SqlConnection
        Try
            ' myConn_FA = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=FASYSTEMPH8;Integrated Security=False;User Id=sa;Password=p@sswd;")
            myConn_FA = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=FASYSTEMPH8;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            myConn_FA.Open()
            Return myConn_FA
        Catch ex As Exception
            MsgBox("Connect Database Fail Myconn FA" & vbNewLine & ex.Message, 16, "Status in ")
            GoTo next_connected_fa
        End Try
        Return "SORRY CONNECT FAILL"
    End Function
End Class