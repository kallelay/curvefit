Public Class Options

    Private Sub Options_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        AUTO_SCALE = sender.checked
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        SHOW_ESTIM = sender.checked
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        SHOW_MEAS = sender.checked
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        XAXIS_TITLE = sender.Text
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        YAXIS_TITLE = sender.Text
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TITLE = sender.Text
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        MEAS_TITLE = sender.Text
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        ESTIM_TITLE = sender.Text
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        LEGEND_VISIBLE = sender.checked
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub

    Private Sub Options_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CheckBox1.Checked = SHOW_MEAS
        CheckBox2.Checked = SHOW_ESTIM
        CheckBox3.Checked = AUTO_SCALE
        CheckBox4.Checked = LEGEND_VISIBLE

        TextBox3.Text = TITLE
        TextBox1.Text = XAXIS_TITLE
        TextBox2.Text = YAXIS_TITLE
        TextBox4.Text = MEAS_TITLE
        TextBox5.Text = ESTIM_TITLE

        Button1.Text = "legend font" & vbNewLine & PLOT_FONT.ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FontDialog1.Font = PLOT_FONT
        If FontDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PLOT_FONT = FontDialog1.Font
        End If
        Button1.Text = "legend font" & vbNewLine & PLOT_FONT.ToString
        Try : MainForm.BackgroundWorker1.RunWorkerAsync() : Catch ex As Exception : End Try
    End Sub
End Class