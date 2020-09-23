<div align="center">

## Sound Card Information


</div>

### Description

This Code Show Your Sound Card information if you like my code please vote
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Syed Adeel Rizvi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syed-adeel-rizvi.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syed-adeel-rizvi-sound-card-information__1-36424/archive/master.zip)

### API Declarations

```
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Private Const MAXPNAMELEN = 32
Private Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type
```


### Source Code

```
Private Sub Command1_Click()
Dim x As WAVEOUTCAPS
waveOutGetDevCaps 0, x, Len(x)
Label1.Caption = "Sound Card - " & x.szPname
Label2.Caption = "Sound Formats - " & x.dwFormats
Label3.Caption = "Sound Support - " & x.dwSupport
Label4.Caption = "Sound DriverVersion - " & x.vDriverVersion
Label5.Caption = "Sound Channels - " & x.wChannels
Label6.Caption = "Sound Mid - " & x.wMid
Label7.Caption = "Sound Pid - " & x.wPid
End Sub
```

