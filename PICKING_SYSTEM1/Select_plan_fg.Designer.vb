<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class Select_plan_fg
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
        Me.components = New System.ComponentModel.Container
        Dim Button1 As System.Windows.Forms.Button
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Select_plan_fg))
        Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.LB_DAY = New System.Windows.Forms.Label
        Me.days = New System.Windows.Forms.ComboBox
        Me.p_show_confrim = New System.Windows.Forms.Panel
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.OK_CON = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.alert_check_part_no = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel8 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.PictureBox17 = New System.Windows.Forms.PictureBox
        Me.PictureBox11 = New System.Windows.Forms.PictureBox
        Me.PictureBox10 = New System.Windows.Forms.PictureBox
        Me.PictureBox12 = New System.Windows.Forms.PictureBox
        Me.PictureBox13 = New System.Windows.Forms.PictureBox
        Me.PictureBox14 = New System.Windows.Forms.PictureBox
        Me.PictureBox15 = New System.Windows.Forms.PictureBox
        Me.PictureBox16 = New System.Windows.Forms.PictureBox
        Me.btn_ok = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Line_list_view = New System.Windows.Forms.ListView
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.Line_Emp_cd = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.fo = New System.Windows.Forms.TextBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Button1 = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.p_show_confrim.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Button1.BackColor = System.Drawing.Color.Red
        resources.ApplyResources(Button1, "Button1")
        Button1.ForeColor = System.Drawing.Color.White
        Button1.Name = "Button1"
        AddHandler Button1.Click, AddressOf Me.Button1_Click
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Panel3)
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Name = "Panel1"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.MediumPurple
        Me.Panel2.Controls.Add(Me.LB_DAY)
        Me.Panel2.Controls.Add(Me.days)
        Me.Panel2.Controls.Add(Me.p_show_confrim)
        Me.Panel2.Controls.Add(Me.OK_CON)
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Controls.Add(Me.ComboBox2)
        Me.Panel2.Controls.Add(Me.ComboBox1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.alert_check_part_no)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Panel8)
        Me.Panel2.Controls.Add(Me.btn_ok)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Button1)
        Me.Panel2.Controls.Add(Me.Line_list_view)
        Me.Panel2.Controls.Add(Me.Line_Emp_cd)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.fo)
        resources.ApplyResources(Me.Panel2, "Panel2")
        Me.Panel2.Name = "Panel2"
        '
        'LB_DAY
        '
        Me.LB_DAY.BackColor = System.Drawing.Color.Aquamarine
        resources.ApplyResources(Me.LB_DAY, "LB_DAY")
        Me.LB_DAY.Name = "LB_DAY"
        '
        'days
        '
        resources.ApplyResources(Me.days, "days")
        Me.days.Name = "days"
        '
        'p_show_confrim
        '
        Me.p_show_confrim.BackColor = System.Drawing.Color.DarkCyan
        Me.p_show_confrim.Controls.Add(Me.Button4)
        Me.p_show_confrim.Controls.Add(Me.Button3)
        resources.ApplyResources(Me.p_show_confrim, "p_show_confrim")
        Me.p_show_confrim.Name = "p_show_confrim"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        resources.ApplyResources(Me.Button4, "Button4")
        Me.Button4.Name = "Button4"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button3.Name = "Button3"
        '
        'OK_CON
        '
        Me.OK_CON.BackColor = System.Drawing.Color.Yellow
        resources.ApplyResources(Me.OK_CON, "OK_CON")
        Me.OK_CON.ForeColor = System.Drawing.Color.Black
        Me.OK_CON.Name = "OK_CON"
        '
        'PictureBox1
        '
        resources.ApplyResources(Me.PictureBox1, "PictureBox1")
        Me.PictureBox1.Name = "PictureBox1"
        '
        'ComboBox2
        '
        resources.ApplyResources(Me.ComboBox2, "ComboBox2")
        Me.ComboBox2.Name = "ComboBox2"
        '
        'ComboBox1
        '
        resources.ApplyResources(Me.ComboBox1, "ComboBox1")
        Me.ComboBox1.Name = "ComboBox1"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Aquamarine
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Name = "Label2"
        '
        'alert_check_part_no
        '
        resources.ApplyResources(Me.alert_check_part_no, "alert_check_part_no")
        Me.alert_check_part_no.Name = "alert_check_part_no"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Aquamarine
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Name = "Label1"
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.DarkSlateGray
        Me.Panel8.Controls.Add(Me.Label10)
        Me.Panel8.Controls.Add(Me.PictureBox17)
        Me.Panel8.Controls.Add(Me.PictureBox11)
        Me.Panel8.Controls.Add(Me.PictureBox10)
        Me.Panel8.Controls.Add(Me.PictureBox12)
        Me.Panel8.Controls.Add(Me.PictureBox13)
        Me.Panel8.Controls.Add(Me.PictureBox14)
        Me.Panel8.Controls.Add(Me.PictureBox15)
        Me.Panel8.Controls.Add(Me.PictureBox16)
        resources.ApplyResources(Me.Panel8, "Panel8")
        Me.Panel8.Name = "Panel8"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Name = "Label10"
        '
        'PictureBox17
        '
        resources.ApplyResources(Me.PictureBox17, "PictureBox17")
        Me.PictureBox17.Name = "PictureBox17"
        '
        'PictureBox11
        '
        resources.ApplyResources(Me.PictureBox11, "PictureBox11")
        Me.PictureBox11.Name = "PictureBox11"
        '
        'PictureBox10
        '
        resources.ApplyResources(Me.PictureBox10, "PictureBox10")
        Me.PictureBox10.Name = "PictureBox10"
        '
        'PictureBox12
        '
        resources.ApplyResources(Me.PictureBox12, "PictureBox12")
        Me.PictureBox12.Name = "PictureBox12"
        '
        'PictureBox13
        '
        resources.ApplyResources(Me.PictureBox13, "PictureBox13")
        Me.PictureBox13.Name = "PictureBox13"
        '
        'PictureBox14
        '
        resources.ApplyResources(Me.PictureBox14, "PictureBox14")
        Me.PictureBox14.Name = "PictureBox14"
        '
        'PictureBox15
        '
        resources.ApplyResources(Me.PictureBox15, "PictureBox15")
        Me.PictureBox15.Name = "PictureBox15"
        '
        'PictureBox16
        '
        resources.ApplyResources(Me.PictureBox16, "PictureBox16")
        Me.PictureBox16.Name = "PictureBox16"
        '
        'btn_ok
        '
        Me.btn_ok.BackColor = System.Drawing.Color.Green
        resources.ApplyResources(Me.btn_ok, "btn_ok")
        Me.btn_ok.ForeColor = System.Drawing.Color.White
        Me.btn_ok.Name = "btn_ok"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.OrangeRed
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Name = "Label3"
        '
        'Line_list_view
        '
        Me.Line_list_view.BackColor = System.Drawing.Color.White
        Me.Line_list_view.Columns.Add(Me.ColumnHeader2)
        Me.Line_list_view.Columns.Add(Me.ColumnHeader3)
        Me.Line_list_view.Columns.Add(Me.ColumnHeader4)
        resources.ApplyResources(Me.Line_list_view, "Line_list_view")
        Me.Line_list_view.FullRowSelect = True
        ListViewItem1.Text = resources.GetString("Line_list_view.Items")
        Me.Line_list_view.Items.Add(ListViewItem1)
        Me.Line_list_view.Name = "Line_list_view"
        Me.Line_list_view.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        resources.ApplyResources(Me.ColumnHeader2, "ColumnHeader2")
        '
        'ColumnHeader3
        '
        resources.ApplyResources(Me.ColumnHeader3, "ColumnHeader3")
        '
        'ColumnHeader4
        '
        resources.ApplyResources(Me.ColumnHeader4, "ColumnHeader4")
        '
        'Line_Emp_cd
        '
        Me.Line_Emp_cd.BackColor = System.Drawing.Color.Aquamarine
        resources.ApplyResources(Me.Line_Emp_cd, "Line_Emp_cd")
        Me.Line_Emp_cd.ForeColor = System.Drawing.Color.Black
        Me.Line_Emp_cd.Name = "Line_Emp_cd"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Name = "Label8"
        '
        'fo
        '
        resources.ApplyResources(Me.fo, "fo")
        Me.fo.Name = "fo"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.DarkSlateGray
        Me.Panel3.Controls.Add(Me.Panel4)
        resources.ApplyResources(Me.Panel3, "Panel3")
        Me.Panel3.Name = "Panel3"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.White
        Me.Panel4.Controls.Add(Me.Button2)
        Me.Panel4.Controls.Add(Me.Label7)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.Panel7)
        Me.Panel4.Controls.Add(Me.Panel6)
        Me.Panel4.Controls.Add(Me.Panel5)
        Me.Panel4.Controls.Add(Me.Label4)
        resources.ApplyResources(Me.Panel4, "Panel4")
        Me.Panel4.Name = "Panel4"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Name = "Button2"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        resources.ApplyResources(Me.Panel7, "Panel7")
        Me.Panel7.Name = "Panel7"
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.Yellow
        resources.ApplyResources(Me.Panel6, "Panel6")
        Me.Panel6.Name = "Panel6"
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(140, Byte), Integer), CType(CType(236, Byte), Integer))
        resources.ApplyResources(Me.Panel5, "Panel5")
        Me.Panel5.Name = "Panel5"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.ForeColor = System.Drawing.Color.Red
        Me.Label4.Name = "Label4"
        '
        'Timer1
        '
        '
        'Select_plan_fg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.MediumPurple
        resources.ApplyResources(Me, "$this")
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Select_plan_fg"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.p_show_confrim.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Line_Emp_cd As System.Windows.Forms.Label
    Friend WithEvents Line_list_view As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btn_ok As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
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
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents alert_check_part_no As System.Windows.Forms.PictureBox
    Friend WithEvents fo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents OK_CON As System.Windows.Forms.Button
    Friend WithEvents p_show_confrim As System.Windows.Forms.Panel
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents LB_DAY As System.Windows.Forms.Label
    Friend WithEvents days As System.Windows.Forms.ComboBox
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
