Imports System
Imports System.Net
Imports System.IO
Imports System.Windows.Forms.Form
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Module Api
    Public Function DownloadImage1(ByVal imageUrl As String) As System.Drawing.Image
        'Dim img = New Imaging
        Dim temp As System.Drawing.Image = My.Resources.ResourceManager.GetObject("Stock_location/ing")
        'Console.WriteLine(image.ToString());
        DownloadImage1 = temp
    End Function
    Public Function DownloadImage(ByVal _URL As String) As Image
        Dim _tmpImage As Image = Nothing

        Try
            ' Open a connection
            Dim _HttpWebRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create(_URL), System.Net.HttpWebRequest)

            _HttpWebRequest.AllowWriteStreamBuffering = True

            ' You can also specify additional header values like the user agent or the referer: (Optional)
            _HttpWebRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)"
            _HttpWebRequest.Referer = "http://www.google.com/"

            ' set timeout for 20 seconds (Optional)
            _HttpWebRequest.Timeout = 20000

            ' Request response:
            Dim _WebResponse As System.Net.WebResponse = _HttpWebRequest.GetResponse()

            ' Open data stream:
            Dim _WebStream As System.IO.Stream = _WebResponse.GetResponseStream()

            ' convert webstream to image
            _tmpImage = New System.Drawing.Bitmap(_WebStream)


            ' Cleanup
            _WebResponse.Close()
        Catch _Exception As Exception
            ' Error
            'Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
            Return Nothing
        End Try

        Return _tmpImage
    End Function

    Public Function update_data(ByVal _URL As String) As String
        Dim _tmpImage As Image = Nothing
        Dim re_data = "NO_DATA"
        Try
            Dim _HttpWebRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create(_URL), System.Net.HttpWebRequest)

            _HttpWebRequest.AllowWriteStreamBuffering = True
            _HttpWebRequest.Timeout = 20000

            Dim _WebResponse As System.Net.WebResponse = _HttpWebRequest.GetResponse()
            Using data As New StreamReader(_WebResponse.GetResponseStream())
                re_data = data.ReadToEnd
                'MsgBox("re_data" & re_data)
                'MsgBox(data.ReadToEnd)
                '  For Each key In data.ReadToEnd
                're_data &= key
                ' Next 'return to json '
            End Using
            _WebResponse.Close()
        Catch _Exception As Exception
            MsgBox("FALL WOW")
            Return Nothing
        End Try
        Return re_data
    End Function

    Public Function check_net()
        ' MsgBox("01")
        Dim _tmpImage As Image = Nothing
        Dim re_data = "NO_DATA"
        Try
            'MsgBox("02")
            ' Open a connection
            Dim _HttpWebRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create("http://192.168.82.23/member/photo/k0071.jpg"), System.Net.HttpWebRequest)
            'MsgBox("03")
            _HttpWebRequest.AllowWriteStreamBuffering = True
            'MsgBox("04")
            ' You can also specify additional header values like the user agent or the referer: (Optional)
            _HttpWebRequest.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)"
            _HttpWebRequest.Referer = "http://www.google.com/"
            'MsgBox("05")
            ' set timeout for 20 seconds (Optional)
            _HttpWebRequest.Timeout = 20000
            ' MsgBox("06")
            ' Request response:
            Dim _WebResponse As System.Net.WebResponse = _HttpWebRequest.GetResponse()

            ' Open data stream:
            Dim _WebStream As System.IO.Stream = _WebResponse.GetResponseStream()

            ' convert webstream to image
            _tmpImage = New System.Drawing.Bitmap(_WebStream)
            ' Cleanup
            _WebResponse.Close()
            Return True
        Catch _Exception As Exception
            ' Error
            'Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
            Return False
        End Try
        Return re_data
    End Function

    Public Function Get_order_fg(ByVal _URL As String) As String
        Dim _tmpImage As Image = Nothing
        Dim re_data = "NO_DATA"
        Try
            Dim _HttpWebRequest As System.Net.HttpWebRequest = CType(System.Net.HttpWebRequest.Create(_URL), System.Net.HttpWebRequest)

            _HttpWebRequest.AllowWriteStreamBuffering = True
            _HttpWebRequest.Timeout = 20000

            Dim _WebResponse As System.Net.WebResponse = _HttpWebRequest.GetResponse()
            Using data As New StreamReader(_WebResponse.GetResponseStream())
                re_data = data.ReadToEnd
                'MsgBox("re_data" & re_data)
                'MsgBox(data.ReadToEnd)
                '  For Each key In data.ReadToEnd
                're_data &= key
                ' Next 'return to json '
            End Using
            _WebResponse.Close()
        Catch _Exception As Exception
            MsgBox("FALL WOW")
            Return Nothing
        End Try
        Return re_data
    End Function
End Module