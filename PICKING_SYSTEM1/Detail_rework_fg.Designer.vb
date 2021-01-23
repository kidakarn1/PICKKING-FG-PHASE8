<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Detail_rework_fg
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Detail_rework_fg))
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.show_qty = New System.Windows.Forms.TextBox
        Me.show_box = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.scan_qr_fg = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cat_rework = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Button2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Button2.Location = New System.Drawing.Point(15, 477)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(207, 90)
        Me.Button2.TabIndex = 32
        Me.Button2.Text = "BACK"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Yellow
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(251, 477)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(207, 90)
        Me.Button1.TabIndex = 33
        Me.Button1.Text = "MOVE"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel1.Location = New System.Drawing.Point(3, 89)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(473, 17)
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 36.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(3, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(332, 58)
        Me.Label4.Text = "REWORK FG"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(28, 299)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(91, 86)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(322, 6)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(76, 67)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(140, 335)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 50)
        Me.Label2.Text = "QTY :"
        '
        'show_qty
        '
        Me.show_qty.Enabled = False
        Me.show_qty.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Bold)
        Me.show_qty.Location = New System.Drawing.Point(270, 335)
        Me.show_qty.Name = "show_qty"
        Me.show_qty.Size = New System.Drawing.Size(206, 39)
        Me.show_qty.TabIndex = 45
        '
        'show_box
        '
        Me.show_box.Enabled = False
        Me.show_box.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Bold)
        Me.show_box.Location = New System.Drawing.Point(270, 257)
        Me.show_box.Name = "show_box"
        Me.show_box.Size = New System.Drawing.Size(206, 39)
        Me.show_box.TabIndex = 46
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(16, 257)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(248, 45)
        Me.Label3.Text = "NUMBER BOX :"
        '
        'scan_qr_fg
        '
        Me.scan_qr_fg.Font = New System.Drawing.Font("Tahoma", 28.0!, System.Drawing.FontStyle.Bold)
        Me.scan_qr_fg.Location = New System.Drawing.Point(3, 173)
        Me.scan_qr_fg.Name = "scan_qr_fg"
        Me.scan_qr_fg.Size = New System.Drawing.Size(471, 52)
        Me.scan_qr_fg.TabIndex = 58
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 28.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(106, 119)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(276, 51)
        Me.Label5.Text = "SCAN TAG FG"
        '
        'cat_rework
        '
        Me.cat_rework.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.cat_rework.Location = New System.Drawing.Point(182, 413)
        Me.cat_rework.Name = "cat_rework"
        Me.cat_rework.Size = New System.Drawing.Size(270, 42)
        Me.cat_rework.TabIndex = 67
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(47, 413)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(138, 36)
        Me.Label1.Text = "MENU:"
        '
        'Detail_rework_fg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.Ivory
        Me.ClientSize = New System.Drawing.Size(476, 590)
        Me.ControlBox = False
        Me.Controls.Add(Me.cat_rework)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.scan_qr_fg)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.show_box)
        Me.Controls.Add(Me.show_qty)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Font = New System.Drawing.Font("Arial", 16.0!, System.Drawing.FontStyle.Bold)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MinimizeBox = False
        Me.Name = "Detail_rework_fg"
        Me.Text = "Detail_rework_fg"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents show_qty As System.Windows.Forms.TextBox
    Friend WithEvents show_box As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents scan_qr_fg As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cat_rework As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
