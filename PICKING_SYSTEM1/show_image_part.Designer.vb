<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class show_image_part
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(show_image_part))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.show_img_part = New System.Windows.Forms.PictureBox
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.MintCream
        Me.Panel1.Controls.Add(Me.show_img_part)
        Me.Panel1.Controls.Add(Me.WebBrowser1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(202, 130)
        '
        'show_img_part
        '
        Me.show_img_part.Enabled = False
        Me.show_img_part.Image = CType(resources.GetObject("show_img_part.Image"), System.Drawing.Image)
        Me.show_img_part.Location = New System.Drawing.Point(13, 12)
        Me.show_img_part.Name = "show_img_part"
        Me.show_img_part.Size = New System.Drawing.Size(174, 112)
        Me.show_img_part.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(13, 12)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(174, 112)
        '
        'show_image_part
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.MintCream
        Me.ClientSize = New System.Drawing.Size(202, 130)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Enabled = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(40, 268)
        Me.Name = "show_image_part"
        Me.Text = "show_image_part"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents show_img_part As System.Windows.Forms.PictureBox
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
End Class
