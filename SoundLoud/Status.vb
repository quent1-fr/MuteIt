Imports CoreAudio
Public Class Status

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Opacity = 0.9

        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - 256) / 2
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - 256) / 2

        Dim son As New MuteIt.Sound
        Dim DevEnum As New MMDeviceEnumerator()
        Dim volume = son.GetVolume().ToString

        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()

        ' This code makes a beautiful fade out when the timer tick.
        Do Until Me.Opacity = 0
            Me.Opacity -= 0.01
            System.Threading.Thread.Sleep(10)
        Loop

        Me.Close()

    End Sub

End Class
