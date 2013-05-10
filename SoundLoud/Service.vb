Imports MuteIt.KeyboardLowLevelHook
Imports MuteIt.Functions
Public Class Service
    Private WithEvents m_Hook As New MuteIt.KeyboardLowLevelHook

    Private Sub ExitToolStripMenuItem_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click, ExitToolStripMenuItem1.Click
        NotifyIcon1.Dispose()
        End
    End Sub

    Private Sub AboutMuteItToolStripMenuItem_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutMuteItToolStripMenuItem.Click
        About.ShowDialog()
    End Sub

    Private Sub m_Hook_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles m_Hook.KeyDown
        ' When a key which is bound is pressed
        Dim son As New MuteIt.Sound
        Dim status_functions As New MuteIt.Functions
        If (e.Control AndAlso (Keys.Alt AndAlso (e.KeyCode = Keys.Right))) Then ' Increase
            Status.Show()
            son.IncreaseVolume()
            status_functions.ChangeBackground()
            status_functions.ResetTimer()
            status_functions.MakeAPop()
        ElseIf (e.Control AndAlso (Keys.Alt AndAlso (e.KeyCode = Keys.Left))) Then ' Decrease
            Status.Show()
            son.DecreaseVolume()
            status_functions.ChangeBackground()
            status_functions.ResetTimer()
            status_functions.MakeAPop()
        ElseIf (e.Control AndAlso (Keys.Alt AndAlso (e.KeyCode = Keys.Up))) Or (Keys.Control AndAlso (e.Alt AndAlso (e.KeyCode = Keys.Down))) Then ' Mute or restore
            Status.Show()
            son.Mute()
            status_functions.ChangeBackground()
            status_functions.ResetTimer()
            status_functions.MakeAPop()
        End If
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Help.ShowDialog()
    End Sub
End Class