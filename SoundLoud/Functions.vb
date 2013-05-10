Imports CoreAudio
Public Class Functions
    ' Change the image displayed by Status.picturebox1 
    Public Function ChangeBackground()
        Dim son As New MuteIt.Sound
        Dim DevEnum As New MMDeviceEnumerator()
        Dim volume = son.GetVolume().ToString

        If volume = 0 Or DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = True Then
            Status.PictureBox1.Image = My.Resources.muted
        ElseIf volume > 0 AndAlso (volume < 34) Then
            Status.PictureBox1.Image = My.Resources.low
        ElseIf volume > 33 AndAlso (volume < 67) Then
            Status.PictureBox1.Image = My.Resources.medium
        ElseIf volume > 66 Then
            Status.PictureBox1.Image = My.Resources.high
        End If

        Status.Label1.Text = son.GetVolume & "%"

        Return True
    End Function

    Public Function ResetTimer()
        Status.Timer1.Stop()
        Status.Timer1.Start()

        Return True
    End Function
    ' Play a song when a bound key is pressed
    Public Function MakeAPop()
        ' Alternative sound
        'Dim Sound As New System.Media.SoundPlayer(Environment.GetEnvironmentVariable("WINDIR") & "\Media\Windows Information Bar.wav")
        Dim Sound As New System.Media.SoundPlayer(Environment.GetEnvironmentVariable("WINDIR") & "\Media\Windows Menu Command.wav")

        Sound.Load()
        Sound.Play()

        Return True
    End Function
End Class
