Public Class show_image_part

    Private Sub show_image_part_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Public Sub New()
        InitializeComponent()
        If Api.DownloadImage("http://192.168.161.102/picking_system/uploads/pic/" & Module1.M_Part_Selected & ".jpg") IsNot Nothing Then
            show_img_part.Image = Api.DownloadImage("http://192.168.161.102/picking_system/uploads/pic/" & Module1.M_Part_Selected & ".jpg")
        End If
        'show_img_part.Image = 
    End Sub

    Private Sub show_img_part_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles show_img_part.Click

    End Sub
End Class