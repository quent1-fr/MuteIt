Imports CoreAudio

Public Class Sound
    Public Function GetVolume() As Integer
        Dim MasterMinimum As Integer = 0
        Dim DevEnum As New MMDeviceEnumerator()
        Dim device As MMDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia)
        Dim Vol As Integer = 0

        With device.AudioEndpointVolume
            Vol = CInt(.MasterVolumeLevelScalar * 100)
            If Vol < MasterMinimum Then
                Vol = MasterMinimum / 100.0F
            End If
        End With

        Return Vol
    End Function
    Public Function IncreaseVolume() As Integer
        Dim DevEnum As New MMDeviceEnumerator()
        DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.VolumeStepUp()

        Return GetVolume()
    End Function
    Public Function DecreaseVolume() As Integer
        Dim DevEnum As New MMDeviceEnumerator()
        DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.VolumeStepDown()

        Return GetVolume()
    End Function
    Public Function Mute() As Boolean
        Dim DevEnum As New MMDeviceEnumerator()
        If DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = True Then
            DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = False
            Return False
        ElseIf DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = False Then
            DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = True
            Return True
        Else
            Return False
        End If

    End Function
End Class
