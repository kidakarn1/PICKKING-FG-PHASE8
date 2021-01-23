Public Class cut_ng

    Private Sub Panel2_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel2.GotFocus

    End Sub

    Private Sub cut_ng_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("8")
        ComboBox1.Items.Add("10")
        ComboBox2.Items.Add("FG")
        ComboBox2.Items.Add("FW")
        ComboBox2.Items.Add("RM")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class