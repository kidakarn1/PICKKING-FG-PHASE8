<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class scan_location_fg
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(scan_location_fg))
        Me.btn_ok = New System.Windows.Forms.Button
        Me.lb_code_user = New System.Windows.Forms.Label
        Me.lb_code_pd = New System.Windows.Forms.Label
        Me.lb_code_line = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.PictureBox17 = New System.Windows.Forms.PictureBox
        Me.PictureBox11 = New System.Windows.Forms.PictureBox
        Me.PictureBox10 = New System.Windows.Forms.PictureBox
        Me.PictureBox12 = New System.Windows.Forms.PictureBox
        Me.PictureBox13 = New System.Windows.Forms.PictureBox
        Me.PictureBox14 = New System.Windows.Forms.PictureBox
        Me.PictureBox15 = New System.Windows.Forms.PictureBox
        Me.PictureBox16 = New System.Windows.Forms.PictureBox
        Me.STOCK_QTY = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.text_box_location = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Lot_No = New System.Windows.Forms.Label
        Me.QTY = New System.Windows.Forms.Label
        Me.Location = New System.Windows.Forms.Label
        Me.Part_Name = New System.Windows.Forms.Label
        Me.Part_Selected = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.fo = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_ok
        '
        Me.btn_ok.BackColor = System.Drawing.Color.Red
        Me.btn_ok.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.btn_ok.ForeColor = System.Drawing.Color.White
        Me.btn_ok.Location = New System.Drawing.Point(16, 470)
        Me.btn_ok.Name = "btn_ok"
        Me.btn_ok.Size = New System.Drawing.Size(82, 43)
        Me.btn_ok.TabIndex = 0
        Me.btn_ok.Text = "Back"
        '
        'lb_code_user
        '
        Me.lb_code_user.BackColor = System.Drawing.Color.MintCream
        Me.lb_code_user.Font = New System.Drawing.Font("Tahoma", 15.0!, System.Drawing.FontStyle.Bold)
        Me.lb_code_user.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lb_code_user.Location = New System.Drawing.Point(16, 23)
        Me.lb_code_user.Name = "lb_code_user"
        Me.lb_code_user.Size = New System.Drawing.Size(147, 25)
        Me.lb_code_user.Text = "Label1"
        '
        'lb_code_pd
        '
        Me.lb_code_pd.BackColor = System.Drawing.Color.MintCream
        Me.lb_code_pd.Font = New System.Drawing.Font("Tahoma", 15.0!, System.Drawing.FontStyle.Bold)
        Me.lb_code_pd.ForeColor = System.Drawing.Color.Blue
        Me.lb_code_pd.Location = New System.Drawing.Point(177, 23)
        Me.lb_code_pd.Name = "lb_code_pd"
        Me.lb_code_pd.Size = New System.Drawing.Size(120, 25)
        Me.lb_code_pd.Text = "Label3"
        '
        'lb_code_line
        '
        Me.lb_code_line.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lb_code_line.BackColor = System.Drawing.Color.MintCream
        Me.lb_code_line.Font = New System.Drawing.Font("Tahoma", 15.0!, System.Drawing.FontStyle.Bold)
        Me.lb_code_line.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lb_code_line.Location = New System.Drawing.Point(303, 23)
        Me.lb_code_line.Name = "lb_code_line"
        Me.lb_code_line.Size = New System.Drawing.Size(159, 63)
        Me.lb_code_line.Text = "Label2"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(476, 590)
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.MintCream
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.STOCK_QTY)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.text_box_location)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.btn_ok)
        Me.Panel1.Controls.Add(Me.Lot_No)
        Me.Panel1.Controls.Add(Me.QTY)
        Me.Panel1.Controls.Add(Me.Location)
        Me.Panel1.Controls.Add(Me.Part_Name)
        Me.Panel1.Controls.Add(Me.Part_Selected)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.lb_code_line)
        Me.Panel1.Controls.Add(Me.lb_code_pd)
        Me.Panel1.Controls.Add(Me.lb_code_user)
        Me.Panel1.Location = New System.Drawing.Point(8, 44)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(462, 544)
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(0, 63)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(469, 381)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.DarkSlateGray
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.PictureBox17)
        Me.Panel2.Controls.Add(Me.PictureBox11)
        Me.Panel2.Controls.Add(Me.PictureBox10)
        Me.Panel2.Controls.Add(Me.PictureBox12)
        Me.Panel2.Controls.Add(Me.PictureBox13)
        Me.Panel2.Controls.Add(Me.PictureBox14)
        Me.Panel2.Controls.Add(Me.PictureBox15)
        Me.Panel2.Controls.Add(Me.PictureBox16)
        Me.Panel2.Location = New System.Drawing.Point(83, 191)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(272, 292)
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(39, 248)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(222, 51)
        Me.Label10.Text = "กำลังโหลดข้อมูล"
        '
        'PictureBox17
        '
        Me.PictureBox17.Image = CType(resources.GetObject("PictureBox17.Image"), System.Drawing.Image)
        Me.PictureBox17.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox17.Name = "PictureBox17"
        Me.PictureBox17.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox17.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox11
        '
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox10
        '
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox12
        '
        Me.PictureBox12.Image = CType(resources.GetObject("PictureBox12.Image"), System.Drawing.Image)
        Me.PictureBox12.Location = New System.Drawing.Point(14, 12)
        Me.PictureBox12.Name = "PictureBox12"
        Me.PictureBox12.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox13
        '
        Me.PictureBox13.Image = CType(resources.GetObject("PictureBox13.Image"), System.Drawing.Image)
        Me.PictureBox13.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox13.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox14
        '
        Me.PictureBox14.Image = CType(resources.GetObject("PictureBox14.Image"), System.Drawing.Image)
        Me.PictureBox14.Location = New System.Drawing.Point(14, 12)
        Me.PictureBox14.Name = "PictureBox14"
        Me.PictureBox14.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox15
        '
        Me.PictureBox15.Image = CType(resources.GetObject("PictureBox15.Image"), System.Drawing.Image)
        Me.PictureBox15.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'PictureBox16
        '
        Me.PictureBox16.Image = CType(resources.GetObject("PictureBox16.Image"), System.Drawing.Image)
        Me.PictureBox16.Location = New System.Drawing.Point(15, 12)
        Me.PictureBox16.Name = "PictureBox16"
        Me.PictureBox16.Size = New System.Drawing.Size(246, 227)
        Me.PictureBox16.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'STOCK_QTY
        '
        Me.STOCK_QTY.BackColor = System.Drawing.Color.White
        Me.STOCK_QTY.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.STOCK_QTY.Location = New System.Drawing.Point(24, 275)
        Me.STOCK_QTY.Name = "STOCK_QTY"
        Me.STOCK_QTY.Size = New System.Drawing.Size(408, 46)
        Me.STOCK_QTY.Text = "STOCK QTY :"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Green
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(361, 470)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 43)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "OK"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(90, 447)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(253, 20)
        Me.Label3.Text = "กรุณา สแกน QR CODE ของพื้นที่"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'text_box_location
        '
        Me.text_box_location.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.text_box_location.Location = New System.Drawing.Point(60, 407)
        Me.text_box_location.Name = "text_box_location"
        Me.text_box_location.Size = New System.Drawing.Size(342, 35)
        Me.text_box_location.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 16.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(43, 371)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(375, 33)
        Me.Label2.Text = "Please scan QR code at location"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Lot_No
        '
        Me.Lot_No.BackColor = System.Drawing.Color.White
        Me.Lot_No.Font = New System.Drawing.Font("Tahoma", 15.0!, System.Drawing.FontStyle.Bold)
        Me.Lot_No.Location = New System.Drawing.Point(16, 516)
        Me.Lot_No.Name = "Lot_No"
        Me.Lot_No.Size = New System.Drawing.Size(416, 20)
        Me.Lot_No.Text = "Lot No."
        '
        'QTY
        '
        Me.QTY.BackColor = System.Drawing.Color.White
        Me.QTY.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.QTY.ForeColor = System.Drawing.Color.Fuchsia
        Me.QTY.Location = New System.Drawing.Point(23, 321)
        Me.QTY.Name = "QTY"
        Me.QTY.Size = New System.Drawing.Size(415, 33)
        Me.QTY.Text = "QTY :"
        '
        'Location
        '
        Me.Location.BackColor = System.Drawing.Color.White
        Me.Location.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Location.ForeColor = System.Drawing.Color.Red
        Me.Location.Location = New System.Drawing.Point(24, 223)
        Me.Location.Name = "Location"
        Me.Location.Size = New System.Drawing.Size(416, 27)
        Me.Location.Text = "Location :"
        '
        'Part_Name
        '
        Me.Part_Name.BackColor = System.Drawing.Color.White
        Me.Part_Name.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Part_Name.ForeColor = System.Drawing.Color.Green
        Me.Part_Name.Location = New System.Drawing.Point(24, 171)
        Me.Part_Name.Name = "Part_Name"
        Me.Part_Name.Size = New System.Drawing.Size(416, 51)
        Me.Part_Name.Text = "Part Name :"
        '
        'Part_Selected
        '
        Me.Part_Selected.BackColor = System.Drawing.Color.White
        Me.Part_Selected.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold)
        Me.Part_Selected.ForeColor = System.Drawing.Color.Blue
        Me.Part_Selected.Location = New System.Drawing.Point(24, 115)
        Me.Part_Selected.Name = "Part_Selected"
        Me.Part_Selected.Size = New System.Drawing.Size(416, 47)
        Me.Part_Selected.Text = "Part Selected :"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 20.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(177, 63)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 41)
        Me.Label1.Text = "Details"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(3, 79)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(456, 284)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'fo
        '
        Me.fo.Location = New System.Drawing.Point(125, 15)
        Me.fo.Name = "fo"
        Me.fo.Size = New System.Drawing.Size(140, 23)
        Me.fo.TabIndex = 27
        '
        'Timer1
        '
        '
        'scan_location_fg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(477, 591)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.fo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "scan_location_fg"
        Me.Text = "scan_location"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn_ok As System.Windows.Forms.Button
    Friend WithEvents lb_code_user As System.Windows.Forms.Label
    Friend WithEvents lb_code_pd As System.Windows.Forms.Label
    Friend WithEvents lb_code_line As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Location As System.Windows.Forms.Label
    Friend WithEvents Part_Name As System.Windows.Forms.Label
    Friend WithEvents Part_Selected As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Lot_No As System.Windows.Forms.Label
    Friend WithEvents QTY As System.Windows.Forms.Label
    Friend WithEvents text_box_location As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents STOCK_QTY As System.Windows.Forms.Label
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents fo As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents PictureBox17 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox12 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox13 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox14 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox16 As System.Windows.Forms.PictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
