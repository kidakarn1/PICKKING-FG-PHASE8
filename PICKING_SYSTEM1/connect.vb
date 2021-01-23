Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration

Public Class connect
    Public myConn As SqlConnection
    Public myConn_FA As SqlConnection
    Public Function conn()
        Dim myConn As SqlConnection

        Try
            ' myConn = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
            myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            myConn.Open()
            Return myConn
        Catch ex As Exception
            MsgBox("Connect Database Fail" & vbNewLine & ex.Message, 16, "Status in")
        End Try
        Return "SORRY CONNECT FAILL"
    End Function
    Public Function conn_fa()
        Dim myConn_FA As SqlConnection
        Try
            ' myConn_FA = New SqlConnection("Data Source= 192.168.10.19\SQLEXPRESS2017,1433;Initial Catalog=FASYSTEMPH8;Integrated Security=False;User Id=sa;Password=p@sswd;")
            myConn_FA = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=FASYSTEMPH8;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
            myConn_FA.Open()
            Return myConn_FA
        Catch ex As Exception
            MsgBox("Connect Database Fail Myconn FA" & vbNewLine & ex.Message, 16, "Status in ")
        End Try
        Return "SORRY CONNECT FAILL"
    End Function
End Class