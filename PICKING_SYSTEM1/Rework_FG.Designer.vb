<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Rework_FG
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Rework_FG))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel_select = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel_confirm = New System.Windows.Forms.Panel
        Me.btn_CANCEL = New System.Windows.Forms.Button
        Me.btn_confirm = New System.Windows.Forms.Button
        Me.QTY = New System.Windows.Forms.TextBox
        Me.part_no = New System.Windows.Forms.TextBox
        Me.scan_qr = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.inst = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cat_rework = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel_select.SuspendLayout()
        Me.Panel_confirm.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Panel1.Controls.Add(Me.Panel_select)
        Me.Panel1.Controls.Add(Me.Panel_confirm)
        Me.Panel1.Controls.Add(Me.QTY)
        Me.Panel1.Controls.Add(Me.part_no)
        Me.Panel1.Controls.Add(Me.scan_qr)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.inst)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.cat_rework)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(476, 598)
        '
        'Panel_select
        '
        Me.Panel_select.BackColor = System.Drawing.Color.White
        Me.Panel_select.Controls.Add(Me.Label1)
        Me.Panel_select.Controls.Add(Me.Panel2)
        Me.Panel_select.Controls.Add(Me.Button5)
        Me.Panel_select.Controls.Add(Me.Button4)
        Me.Panel_select.Controls.Add(Me.Button3)
        Me.Panel_select.Controls.Add(Me.PictureBox1)
        Me.Panel_select.Location = New System.Drawing.Point(0, 0)
        Me.Panel_select.Name = "Panel_select"
        Me.Panel_select.Size = New System.Drawing.Size(476, 590)
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 36.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(0, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(332, 58)
        Me.Label1.Text = "REWORK FG"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkOrange
        Me.Panel2.Location = New System.Drawing.Point(0, 81)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(488, 17)
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Red
        Me.Button5.Font = New System.Drawing.Font("Tahoma", 26.0!, System.Drawing.FontStyle.Bold)
        Me.Button5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Button5.Location = New System.Drawing.Point(8, 532)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(115, 47)
        Me.Button5.TabIndex = 2
        Me.Button5.Text = "BACK"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.OrangeRed
        Me.Button4.Font = New System.Drawing.Font("Tahoma", 36.0!, System.Drawing.FontStyle.Bold)
        Me.Button4.Location = New System.Drawing.Point(85, 328)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(325, 183)
        Me.Button4.TabIndex = 1
        Me.Button4.Text = "NOT FULL"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Yellow
        Me.Button3.Font = New System.Drawing.Font("Tahoma", 36.0!, System.Drawing.FontStyle.Bold)
        Me.Button3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button3.Location = New System.Drawing.Point(85, 109)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(325, 187)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "FULL"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 92)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(476, 495)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'Panel_confirm
        '
        Me.Panel_confirm.BackColor = System.Drawing.Color.SlateGray
        Me.Panel_confirm.Controls.Add(Me.btn_CANCEL)
        Me.Panel_confirm.Controls.Add(Me.btn_confirm)
        Me.Panel_confirm.Location = New System.Drawing.Point(26, 71)
        Me.Panel_confirm.Name = "Panel_confirm"
        Me.Panel_confirm.Size = New System.Drawing.Size(412, 443)
        '
        'btn_CANCEL
        '
        Me.btn_CANCEL.BackColor = System.Drawing.Color.Red
        Me.btn_CANCEL.Font = New System.Drawing.Font("Tahoma", 48.0!, System.Drawing.FontStyle.Bold)
        Me.btn_CANCEL.ForeColor = System.Drawing.Color.White
        Me.btn_CANCEL.Location = New System.Drawing.Point(32, 231)
        Me.btn_CANCEL.Name = "btn_CANCEL"
        Me.btn_CANCEL.Size = New System.Drawing.Size(352, 196)
        Me.btn_CANCEL.TabIndex = 1
        Me.btn_CANCEL.Text = "CANCEL"
        '
        'btn_confirm
        '
        Me.btn_confirm.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btn_confirm.Font = New System.Drawing.Font("Tahoma", 48.0!, System.Drawing.FontStyle.Bold)
        Me.btn_confirm.ForeColor = System.Drawing.Color.White
        Me.btn_confirm.Location = New System.Drawing.Point(32, 21)
        Me.btn_confirm.Name = "btn_confirm"
        Me.btn_confirm.Size = New System.Drawing.Size(352, 196)
        Me.btn_confirm.TabIndex = 0
        Me.btn_confirm.Text = "CONFIRM"
        '
        'QTY
        '
        Me.QTY.AcceptsReturn = True
        Me.QTY.Enabled = False
        Me.QTY.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.QTY.Location = New System.Drawing.Point(197, 250)
        Me.QTY.Name = "QTY"
        Me.QTY.Size = New System.Drawing.Size(270, 42)
        Me.QTY.TabIndex = 38
        '
        'part_no
        '
        Me.part_no.AcceptsReturn = True
        Me.part_no.Enabled = False
        Me.part_no.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.part_no.Location = New System.Drawing.Point(197, 180)
        Me.part_no.Name = "part_no"
        Me.part_no.Size = New System.Drawing.Size(270, 42)
        Me.part_no.TabIndex = 37
        '
        'scan_qr
        '
        Me.scan_qr.Font = New System.Drawing.Font("Tahoma", 28.0!, System.Drawing.FontStyle.Bold)
        Me.scan_qr.Location = New System.Drawing.Point(3, 13)
        Me.scan_qr.Name = "scan_qr"
        Me.scan_qr.Size = New System.Drawing.Size(464, 52)
        Me.scan_qr.TabIndex = 36
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Button2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Button2.Location = New System.Drawing.Point(12, 442)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(207, 90)
        Me.Button2.TabIndex = 31
        Me.Button2.Text = "BACK"
        '
        'inst
        '
        Me.inst.AcceptsReturn = True
        Me.inst.Enabled = False
        Me.inst.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.inst.Location = New System.Drawing.Point(197, 104)
        Me.inst.Name = "inst"
        Me.inst.Size = New System.Drawing.Size(270, 42)
        Me.inst.TabIndex = 22
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(6, 104)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(205, 45)
        Me.Label5.Text = "Instruction :"
        '
        'cat_rework
        '
        Me.cat_rework.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.cat_rework.Location = New System.Drawing.Point(197, 328)
        Me.cat_rework.Name = "cat_rework"
        Me.cat_rework.Size = New System.Drawing.Size(270, 42)
        Me.cat_rework.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(6, 332)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(138, 36)
        Me.Label4.Text = "MENU:"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 22.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(6, 250)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 36)
        Me.Label3.Text = "QTY :"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(6, 180)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(164, 36)
        Me.Label2.Text = "Part No:"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Chartreuse
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(248, 442)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(207, 90)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "OK"
        '
        'Rework_FG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.ClientSize = New System.Drawing.Size(476, 590)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Rework_FG"
        Me.Text = "Rework_FG"
        Me.Panel1.ResumeLayout(False)
        Me.Panel_select.ResumeLayout(False)
        Me.Panel_confirm.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cat_rework As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents inst As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents QTY As System.Windows.Forms.TextBox
    Friend WithEvents part_no As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents scan_qr As System.Windows.Forms.TextBox
    Friend WithEvents Panel_confirm As System.Windows.Forms.Panel
    Friend WithEvents btn_confirm As System.Windows.Forms.Button
    Friend WithEvents btn_CANCEL As System.Windows.Forms.Button
    Friend WithEvents Panel_select As System.Windows.Forms.Panel
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
End Class
