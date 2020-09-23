VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sjoerd MP3 Cut off"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveXtra 
      Caption         =   "Save Xtra"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play file"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog c 
      Left            =   2040
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "MP3 Files (*.mp3)*.mp3"
      FilterIndex     =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save file"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin ComctlLib.Slider s 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      _Version        =   327682
      TickStyle       =   3
   End
   Begin ComctlLib.ProgressBar pBar 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin MCI.MMControl m 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "&Open file"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.Slider s1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      _Version        =   327682
      TickStyle       =   3
   End
   Begin WMPLibCtl.WindowsMediaPlayer wm 
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3836
      _cy             =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi, this code can split an MP3
'First click on Read File and select source MP3
'Then use the upper slider to set how many seconds the new MP3 must last
'from the beginning
'And clikc Save to create the new MP3

'Or use the upper slider to set the start of the new MP3 and use the
'lower slider to set the end of the new MP3
'And click Save Xtra to create the new MP3

'The 'WriteTag' function is from another coder (see clsMp3 for the Pscode link

'U can use this code for whatever you wan't, but please leave credits for me (Sjoerd) or
'when you use the class or the 'WriteTag' function leave credits for this coder
'Thanks and much fun!!!

'Greetings, Sjoerd

'Please vote for me and leave comments

Option Explicit
Private strData       As String
Private R             As Integer
Private intI          As Long
Private strG          As String
Private strO          As String
Private intO          As Long
Private intF          As Long
Private FileName      As String
Private FileName1     As String

Private Sub cmdPlay_Click()

    wm.URL = FileName

End Sub

Private Sub cmdRead_Click()

On Error GoTo errHandler

    cmdPlay.Enabled = False
    c.ShowOpen
    FileName = c.FileName
    R = FreeFile
    Me.Caption = "Openening file..."
    Open FileName For Binary As #1
    strData = Space$(LOF(R))
    intO = LOF(R)
    Me.Caption = "Reading data..."
    Get #1, , strData
    Close #1
    With m
        .FileName = FileName
        .Command = "Open"
        intF = .Length / 1000
        pBar.Max = .Length
        s.Max = .Length
    End With 'm
    s1.Max = s.Max
    m.Command = "Close"
    Me.Caption = "File read"
    cmdPlay.Enabled = True
    cmdSaveXtra.Enabled = True
    cmdSave.Enabled = True

Exit Sub

errHandler:

MsgBox "There's been an error, maybe the MP3 you defined doesn't exists", vbCritical, "Sjoerd MP3 Cut Of"
Kill FileName
End Sub

Private Sub cmdSave_Click()

    Me.Caption = "Saving..."
    intI = Round(intO / intF, 0)
    intI = intI * Round(pBar.Value / 1000, 0)
    strG = Left$(strData, intI - -5000)
    strO = Space$(Len(strG))
    strO = strG
    c.ShowSave
    FileName1 = c.FileName
    On Error Resume Next
    Kill FileName1
    Open FileName1 For Binary As #3
    Put #3, , strO
    Close #3
    Me.Caption = "Done"
    On Error GoTo 0

End Sub

Private Sub cmdSaveXtra_Click()

  Dim strDatas As String
  Dim Lenn     As Long
  Dim IDE3     As New clsMp3
  Dim Lenn1    As Long

    If s.Value >= s1.Value Then
        MsgBox "The value of the upper slider must be lower then the value of the lower slider", vbCritical, "Sjoerd MP3 Cut Off"
        Exit Sub
    End If
    c.ShowSave
    FileName1 = c.FileName
    IDE3.ReadMP3 FileName
    On Error Resume Next
    Kill FileName1
    If LenB(IDE3.Songname) = 0 Then
        IDE3.Songname = FileName1
    End If
    If LenB(IDE3.Artist) = 0 Then
        IDE3.Artist = "Sjoerd MP3 Cut Off"
    End If
    If LenB(IDE3.Album) = 0 Then
        IDE3.Album = "No album data"
    End If
    If LenB(IDE3.Year) = 0 Then
        IDE3.Year = Year(Now)
    End If
    If LenB(IDE3.Comment) = 0 Then
        IDE3.Comment = "MP3 splitted with Sjoerd MP3 Cut Off"
    End If
    If LenB(IDE3.Genre) = 0 Then
        IDE3.Genre = "No genre data"
    End If '
    WriteTag FileName1, IDE3.Songname, IDE3.Artist, IDE3.Album, IDE3.Year, IDE3.Comment, IDE3.Genre
    Me.Caption = "Saving file..."
    intI = Round(intO / intF, 0)
    Lenn = Round(s1.Value \ 1000, 0)
    Lenn1 = Round(s.Value \ 1000, 0)
    Lenn = Lenn - Lenn1
    Lenn = Lenn * intI
    strDatas = Mid$(strData, Lenn1 * intI, Lenn)
    Open FileName1 For Binary As #4
    Put #4, , strDatas
    Close #4
    Me.Caption = "Done"
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    c.Filter = "MP3 files (*.mp3) |*.mp3"
    c.FilterIndex = 1

End Sub

Private Sub s1_Change()

    Me.Caption = Round(s.Value / 1000, 0) & " seconds | " & Round(s1.Value / 1000, 0) & " seconds"

End Sub

Private Sub s_Change()

    pBar.Value = s.Value
    'IDE3 = 5.000
    '3 seconden = 50.000
    'dus 1 seconde = 1.000 / 3 = 333
    Me.Caption = Round(s.Value / 1000, 0) & " seconds | " & Round(s1.Value / 1000, 0) & " seconds"

End Sub

Private Function WriteTag(ByVal strFileName As String, _
                          ByVal Songname As String, _
                          ByVal Artist As String, _
                          ByVal Album As String, _
                          ByVal strYear As String, _
                          ByVal Comment As String, _
                          ByVal Genre As Integer) As Long

  
  Dim mp3File As Integer
  Dim sn      As String * 30
  Dim com     As String * 30
  Dim art     As String * 30
  Dim alb     As String * 30
  Dim yr      As String * 4

    Me.Tag = "TAG"
    sn = Songname
    com = Comment
    art = Artist
    alb = Album
    yr = strYear
    mp3File = FreeFile
    Open strFileName For Binary Access Write As #mp3File
    Seek #mp3File, FileLen(strFileName) - 127
    Put #mp3File, , Me.Tag
    Put #mp3File, , sn
    Put #mp3File, , art
    Put #mp3File, , alb
    Put #mp3File, , yr
    Put #mp3File, , com
    Close #mp3File

End Function

':)Roja's VB Code Fixer V1.1.78 (7-2-2004 18:19:25) 22 + 158 = 180 Lines Thanks Ulli for inspiration and lots of code.

