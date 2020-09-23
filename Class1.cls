VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The purpose of this class is to extract information from MP3s. It was written to do the dirty
'work in an ASP MP3 streamer.
'1. ReadMP3() - Reads ID3v1, ID3v2, bitrate, variable bitrate, freq, and duration of a specified
'MP3. If the MP3 has v1 and v2 tags, then v2 will override v1.
'2. SearchDir() - Give this function a dir, and a file type to look for,
'and it will return the subdirs (pipe delimited), and the files (pipe delimited)
'Hi, I've found this code on Planet-Source-Code.com and I wan't to give the original coder
'credits for his work.
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=51204&lngWId=1
'P.s.
'I only use the function ReadMP3 to Read the IDE3 tag from the MP3 file
'The rest of the code I've just left in it's place to credit the coder
'and give other coders the option to use this code.
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WIN32 Types, Constants, and Declares
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Private Const INVALID_HANDLE_VALUE       As Integer = -1
'<:-):WARNING: Unused Const INVALID_HANDLE_VALUE
'<:-)May be a prototype Const you have not yet implimented or left over from a deleted Control.
'<:-):UPDATED: Un-Typed Const with numeric value changed to As Integer
''Private Const DDL_DIRECTORY              As Long = &H10
'<:-):WARNING: Unused Const DDL_DIRECTORY
'<:-)May be a prototype Const you have not yet implimented or left over from a deleted Control.
'<:-):UPDATED: Un-Typed Const with Hex (&H) value changed to As Long
Private Const MAX_PATH                   As Integer = 260
'<:-):UPDATED: Un-Typed Const with numeric value changed to As Integer
Private Type FILETIME
    dwLowDateTime                          As Long
    dwHighDateTime                         As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                       As Long
    ftCreationTime                         As FILETIME
    ftLastAccessTime                       As FILETIME
    ftLastWriteTime                        As FILETIME
    nFileSizeHigh                          As Long
    nFileSizeLow                           As Long
    dwReserved0                            As Long
    dwReserved1                            As Long
    cFileName                              As String * MAX_PATH
    cAlternate                             As String * 14
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'application level variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Private GetMP3Info                       As MP3Info
'<:-):WARNING: Unused Variable GetMP3Info
'<:-)May be a prototype GetMP3Info you have not yet implimented or left over from a deleted Control.
Private Const ErrBase                    As Integer = 9994
'<:-):UPDATED: Un-Typed Const with numeric value changed to As Integer
Private Type MP3Info
    Songname                               As String
    Artist                                 As String
    Album                                  As String
    Year                                   As String
    Comment                                As String
    Genre                                  As String
    Track                                  As String
    VBR                                    As Boolean
End Type
Private CFm_Duration                          As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Frequency                         As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Mode                              As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_MpegLayer                         As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_MpegVersion                       As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_BitRate                           As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Songname                          As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Artist                            As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Album                             As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Year                              As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
'<:-):WARNING: Variables with same name as a VB Command/Property making code difficult to read.
'<:-):SUGGESTION: Change the variable name to (strYear).
'<:-)Do NOT use Search and Replace; you will have to check each occurance
'<:-)to determine whether it is the Command/Property or Variable being referenced.
Private CFm_Comment                           As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Genre                             As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Track                             As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_VBR                               As Boolean
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Files                             As String
'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
Private CFm_Subdirs                           As String

'<:-):WARNING: Public Variables are poor Class coding, This variable has been converted to a private and a Property added to use it
''Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'<:-):WARNING: Unused Declare FindFirstFile
'<:-)May be a prototype Declare you have not yet implimented or left over from a deleted Control.
''Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'<:-):WARNING: Unused Declare FindNextFile
'<:-)May be a prototype Declare you have not yet implimented or left over from a deleted Control.
''Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'<:-):WARNING: Unused Declare FindClose
'<:-)May be a prototype Declare you have not yet implimented or left over from a deleted Control.
Public Property Get Duration() As String

  '<:-) Property created to replace Public Variable in Class

    Duration = CFm_Duration

End Property

Public Property Let Duration(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Duration = PropVal

End Property

Public Property Get Frequency() As String

  '<:-) Property created to replace Public Variable in Class

    Frequency = CFm_Frequency

End Property

Public Property Let Frequency(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Frequency = PropVal

End Property

Public Property Get Mode() As String

  '<:-) Property created to replace Public Variable in Class

    Mode = CFm_Mode

End Property

Public Property Let Mode(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Mode = PropVal

End Property

Public Property Get MpegLayer() As String

  '<:-) Property created to replace Public Variable in Class

    MpegLayer = CFm_MpegLayer

End Property

Public Property Let MpegLayer(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_MpegLayer = PropVal

End Property

Public Property Get MpegVersion() As String

  '<:-) Property created to replace Public Variable in Class

    MpegVersion = CFm_MpegVersion

End Property

Public Property Let MpegVersion(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_MpegVersion = PropVal

End Property

Public Property Get BitRate() As String

  '<:-) Property created to replace Public Variable in Class

    BitRate = CFm_BitRate

End Property

Public Property Let BitRate(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_BitRate = PropVal

End Property

Public Property Get Songname() As String

  '<:-) Property created to replace Public Variable in Class

    Songname = CFm_Songname

End Property

Public Property Let Songname(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Songname = PropVal

End Property

Public Property Get Artist() As String

  '<:-) Property created to replace Public Variable in Class

    Artist = CFm_Artist

End Property

Public Property Let Artist(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Artist = PropVal

End Property

Public Property Get Album() As String

  '<:-) Property created to replace Public Variable in Class

    Album = CFm_Album

End Property

Public Property Let Album(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Album = PropVal

End Property

Public Property Get Year() As String

  '<:-) Property created to replace Public Variable in Class

    Year = CFm_Year

End Property

Public Property Let Year(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Year = PropVal

End Property

Public Property Get Comment() As String

  '<:-) Property created to replace Public Variable in Class

    Comment = CFm_Comment

End Property

Public Property Let Comment(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Comment = PropVal

End Property

Public Property Get Genre() As String

  '<:-) Property created to replace Public Variable in Class

    Genre = CFm_Genre

End Property

Public Property Let Genre(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Genre = PropVal

End Property

Public Property Get Track() As String

  '<:-) Property created to replace Public Variable in Class

    Track = CFm_Track

End Property

Public Property Let Track(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Track = PropVal

End Property

Public Property Get VBR() As Boolean

  '<:-) Property created to replace Public Variable in Class

    VBR = CFm_VBR

End Property

Public Property Let VBR(PropVal As Boolean)

  '<:-) Property created to replace Public Variable in Class

    CFm_VBR = PropVal

End Property

Public Property Get Files() As String

  '<:-) Property created to replace Public Variable in Class

    Files = CFm_Files

End Property

Public Property Let Files(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Files = PropVal

End Property

Public Property Get Subdirs() As String

  '<:-) Property created to replace Public Variable in Class

    Subdirs = CFm_Subdirs

End Property

Public Property Let Subdirs(PropVal As String)

  '<:-) Property created to replace Public Variable in Class

    CFm_Subdirs = PropVal

End Property

Private Function BinToDec(BinValue As String) As Long

  Dim I As Long

    On Error GoTo ErrHand
    BinToDec = 0
    For I = 1 To Len(BinValue)
        If Mid$(BinValue, I, 1) = 1 Then
            BinToDec = BinToDec + 2 ^ (Len(BinValue) - I)
        End If
    Next I

Exit Function

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Private Function ByteToBit(ByteArray As Variant) As String

  '<:-):WARNING: Untyped Parameters auto-Typed. May be not be correct.
  '<:-)PREVIOUS CODE : Private Function ByteToBit(ByteArray) As String
  
  Dim Z As Integer '<:-):SUGGESTION: Integer could be upgraded to Long.
  Dim I As Integer '<:-):SUGGESTION: Integer could be upgraded to Long.

    On Error GoTo ErrHand
    'convert 4*1 byte array to 4*8 bits'''''
    ByteToBit = ""
    For Z = 1 To 4
        For I = 7 To 0 Step -1
            If Int(ByteArray(Z) / (2 ^ I)) = 1 Then
                ByteToBit = ByteToBit & "1"
                ByteArray(Z) = ByteArray(Z) - (2 ^ I)
             ElseIf LenB(ByteToBit) Then 'NOT INT(BYTEARRAY(Z)...
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : ElseIf ByteToBit <> "" Then
                ByteToBit = ByteToBit & "0"
            End If
        Next I '<:-)For-Variable Inserted
    Next Z

Exit Function

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Private Function GetGenreFromID(ByVal strListGenre As Integer) As String

  '<:-):WARNING: Large procedure (261 lines of code)
  '<:-)It is recommended that you try to break it into smaller procedures

    On Error GoTo ErrHand
    If strListGenre = 0 Then
        GetGenreFromID = "Blues"
     ElseIf strListGenre = 1 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Classic Rock"
     ElseIf strListGenre = 2 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Country"
     ElseIf strListGenre = 3 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Dance"
     ElseIf strListGenre = 4 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Disco"
     ElseIf strListGenre = 5 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Funk"
     ElseIf strListGenre = 6 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Grunge"
     ElseIf strListGenre = 7 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Hip-Hop"
     ElseIf strListGenre = 8 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Jazz"
     ElseIf strListGenre = 9 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Metal"
     ElseIf strListGenre = 10 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "New Age"
     ElseIf strListGenre = 11 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Oldies"
     ElseIf strListGenre = 12 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Other"
     ElseIf strListGenre = 13 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Pop"
     ElseIf strListGenre = 14 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "R&B"
     ElseIf strListGenre = 15 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Rap"
     ElseIf strListGenre = 16 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Reggae"
     ElseIf strListGenre = 17 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Rock"
     ElseIf strListGenre = 18 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Techno"
     ElseIf strListGenre = 19 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Industrial"
     ElseIf strListGenre = 20 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Alternative"
     ElseIf strListGenre = 21 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Ska"
     ElseIf strListGenre = 22 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Death Metal"
     ElseIf strListGenre = 23 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Pranks"
     ElseIf strListGenre = 24 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Soundtrack"
     ElseIf strListGenre = 25 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Euro-Techno"
     ElseIf strListGenre = 26 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Ambient"
     ElseIf strListGenre = 27 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Trip-Hop"
     ElseIf strListGenre = 28 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Vocal"
     ElseIf strListGenre = 29 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Jazz+Funk"
     ElseIf strListGenre = 30 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Fusion"
     ElseIf strListGenre = 31 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Trance"
     ElseIf strListGenre = 32 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Classical"
     ElseIf strListGenre = 33 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Instrumental"
     ElseIf strListGenre = 34 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Acid"
     ElseIf strListGenre = 35 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "House"
     ElseIf strListGenre = 36 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Game"
     ElseIf strListGenre = 37 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Sound Clip"
     ElseIf strListGenre = 38 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Gospel"
     ElseIf strListGenre = 39 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Noise"
     ElseIf strListGenre = 40 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "AlternRock"
     ElseIf strListGenre = 41 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Bass"
     ElseIf strListGenre = 42 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Soul"
     ElseIf strListGenre = 43 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Punk"
     ElseIf strListGenre = 44 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Space"
     ElseIf strListGenre = 45 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Meditative"
     ElseIf strListGenre = 46 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Instrumental Pop"
     ElseIf strListGenre = 47 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Instrumental Rock"
     ElseIf strListGenre = 48 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Ethnic"
     ElseIf strListGenre = 49 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Gothic"
     ElseIf strListGenre = 50 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Darkwave"
     ElseIf strListGenre = 51 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Techno-Indistrial"
     ElseIf strListGenre = 52 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Electronic"
     ElseIf strListGenre = 53 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Pop-Folk"
     ElseIf strListGenre = 54 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Eurodance"
     ElseIf strListGenre = 55 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Dream"
     ElseIf strListGenre = 56 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Southern Rock"
     ElseIf strListGenre = 57 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Comedy"
     ElseIf strListGenre = 58 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Cult"
     ElseIf strListGenre = 59 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Gansta"
     ElseIf strListGenre = 60 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Top 40"
     ElseIf strListGenre = 61 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Christian Rap"
     ElseIf strListGenre = 62 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Pop/Funk"
     ElseIf strListGenre = 63 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Jungle"
     ElseIf strListGenre = 64 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Native American"
     ElseIf strListGenre = 65 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Cabaret"
     ElseIf strListGenre = 66 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "New Wave"
     ElseIf strListGenre = 67 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Psychaledic"
     ElseIf strListGenre = 68 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Rave"
     ElseIf strListGenre = 69 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Showtunes"
     ElseIf strListGenre = 70 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Trailer"
     ElseIf strListGenre = 71 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Lo-Fi"
     ElseIf strListGenre = 72 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Tribal"
     ElseIf strListGenre = 73 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Acid Punk"
     ElseIf strListGenre = 74 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Acid Jazz"
     ElseIf strListGenre = 75 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Polka"
     ElseIf strListGenre = 76 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Retro"
     ElseIf strListGenre = 77 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Musical"
     ElseIf strListGenre = 78 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Rock & Roll"
     ElseIf strListGenre = 79 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Hard Rock"
     ElseIf strListGenre = 80 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Folk"
     ElseIf strListGenre = 81 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Folk-Rock"
     ElseIf strListGenre = 82 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "National Folk"
     ElseIf strListGenre = 83 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Swing"
     ElseIf strListGenre = 84 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Fast Fusion"
     ElseIf strListGenre = 85 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Bebop"
     ElseIf strListGenre = 86 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Latin"
     ElseIf strListGenre = 87 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Revival"
     ElseIf strListGenre = 88 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Celtic"
     ElseIf strListGenre = 89 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Bluegrass"
     ElseIf strListGenre = 90 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Avantgarde"
     ElseIf strListGenre = 91 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Gothic Rock"
     ElseIf strListGenre = 92 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Progressive Rock"
     ElseIf strListGenre = 93 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Psychadelic Rock"
     ElseIf strListGenre = 94 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Symphonic Rock"
     ElseIf strListGenre = 95 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Slow Rock"
     ElseIf strListGenre = 96 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Big Band"
     ElseIf strListGenre = 97 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Chorus"
     ElseIf strListGenre = 98 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Easy Listening"
     ElseIf strListGenre = 99 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Acoustic"
     ElseIf strListGenre = 100 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Humour"
     ElseIf strListGenre = 101 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Speech"
     ElseIf strListGenre = 102 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Chanson"
     ElseIf strListGenre = 103 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Opera"
     ElseIf strListGenre = 104 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Chamber Music"
     ElseIf strListGenre = 105 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Sonata"
     ElseIf strListGenre = 106 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Symphony"
     ElseIf strListGenre = 107 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Booty Bass"
     ElseIf strListGenre = 108 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Primus"
     ElseIf strListGenre = 109 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Porn Groove"
     ElseIf strListGenre = 110 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Satire"
     ElseIf strListGenre = 111 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Slow Jam"
     ElseIf strListGenre = 112 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Club"
     ElseIf strListGenre = 113 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Tango"
     ElseIf strListGenre = 114 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Samba"
     ElseIf strListGenre = 115 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Folklore"
     ElseIf strListGenre = 116 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Ballad"
     ElseIf strListGenre = 117 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Power Ballad"
     ElseIf strListGenre = 118 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Rhythmic Soul"
     ElseIf strListGenre = 119 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Freestyle"
     ElseIf strListGenre = 120 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Duet"
     ElseIf strListGenre = 121 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Punk Rock"
     ElseIf strListGenre = 122 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Drum Solo"
     ElseIf strListGenre = 123 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "A Capella"
     ElseIf strListGenre = 124 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Euro-House"
     ElseIf strListGenre = 125 Then 'NOT STRLISTGENRE...
        GetGenreFromID = "Dance Hall"
     ElseIf strListGenre > 125 Then 'NOT STRLISTGENRE...
        GetGenreFromID = ""
        '<:-):SUGGESTION: 'Else' line is a good safety net at end of 'If..ElseIf..End If' structures
    End If

Exit Function

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Private Sub ReadHeader(ByVal strFileName As String)

  '<:-):WARNING: Large procedure (103 lines of code)
  '<:-)It is recommended that you try to break it into smaller procedures
  '<:-):WARNING: Function changed to Sub as nothing is returned via the Function Name.
  '<:-):WARNING: Poorly named Parameters 'FileName'
  '<:-):UPDATED: 'strFileName'
  
  Dim ByteArray(4)    As Byte
  Dim XingH           As String * 4
  Dim FIO             As Long
  Dim I               As Long
  Dim Z               As Long
  Dim X               As Byte
  Dim HeadStart       As Long
  Dim Frames          As Long
  Dim Bin             As String
  Dim Temp            As Variant
  Dim Brate           As Variant
  Dim Freq            As Variant
  Dim VersionLayer(3) As String
  Dim SMode(3)        As String

    On Error GoTo ErrHand
    'tables
    VersionLayer(0) = 0
    VersionLayer(1) = 3
    VersionLayer(2) = 2
    VersionLayer(3) = 1
    SMode(0) = "stereo"
    SMode(1) = "joint stereo"
    SMode(2) = "dual channel"
    SMode(3) = "single channel"
    FIO = FreeFile
    'read the header
    Open strFileName For Binary Access Read As FIO
    If LOF(FIO) < 256 Then
        Close FIO
        Exit Sub
    End If
    '''''start check startposition for header''''''''''''
    '''''if start position <>1 then id3v2 tag exists'''''
    For I = 1 To LOF(FIO)           'check the whole file for the header
        Get #FIO, I, X
        If X = 255 Then             'header always start with 255 followed by 250 or 251
            Get #FIO, I + 1, X
            If X > 249 Then
                If X < 252 Then
                    '<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
                    '<:-) 'If <condition1> Then
                    '<:-) If <condition2> Then '
                    '<:-) Make <condition1> most likely to fail.
                    HeadStart = I       'set header start position
                    Exit For
                End If
            End If
            '<:-)Short Circuit inserted this line
        End If
    Next I
    'no header start position was found
    If HeadStart = 0 Then
        Exit Sub
    End If
    '''end check start position for header'''''''''''''
    ''start check for XingHeader'''
    Get #FIO, HeadStart + 36, XingH
    If XingH = "Xing" Then
        CFm_VBR = True
        For Z = 1 To 4 '
            Get #1, HeadStart + 43 + Z, ByteArray(Z)  'get framelength to array
        Next Z
        Frames = BinToDec(ByteToBit(ByteArray))   'calculate # of frames
     Else 'NOT XINGH...
        CFm_VBR = False
    End If
    '''end check for XingHeader
    '''start extract the first 4 bytes (32 bits) to an array
    For Z = 1 To 4 '
        Get #FIO, HeadStart + Z - 1, ByteArray(Z)
    Next Z
    '''stop extract the first 4 bytes (32 bits) to an array
    Close FIO
    'header string
    Bin = ByteToBit(ByteArray)
    'get mpegversion from table
    CFm_MpegVersion = VersionLayer(BinToDec(Mid$(Bin, 12, 2)))
    'get layer from table
    CFm_MpegLayer = VersionLayer(BinToDec(Mid$(Bin, 14, 2)))
    'get mode from table
    CFm_Mode = SMode(BinToDec(Mid$(Bin, 25, 2)))
    'look for version to create right table
    Select Case CFm_MpegVersion
     Case 1
        'for version 1
        Freq = Array(44100, 48000, 32000)
     Case 2 Or 25
        'for version 2 or 2.5
        Freq = Array(22050, 24000, 16000)
     Case Else
        CFm_Frequency = 0
        Exit Sub
    End Select
    'look for frequency in table
    CFm_Frequency = Freq(BinToDec(Mid$(Bin, 21, 2)))
    If CFm_VBR Then
        '<:-)Pleonasm Removed
        'define to calculate correct bitrate
        Temp = Array(, 12, 144, 144)
        CFm_BitRate = (FileLen(strFileName) * CFm_Frequency) / (Int(Frames)) / 1000 / Temp(CFm_MpegLayer)
     Else 'CFM_VBR = FALSE/0
        'look for the right bitrate table
        Select Case Val(CFm_MpegVersion & CFm_MpegLayer)
         Case 11
            'Version 1, Layer 1
            Brate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
         Case 12
            'V1 L1
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
         Case 13
            'V1 L3
            Brate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
         Case 21 Or 251
            'V2 L1 and 'V2.5 L1
            Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
         Case 22 Or 252 Or 23 Or 253
            'V2 L2 and 'V2.5 L2 etc...
            Brate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
         Case Else
            'if variable bitrate
            CFm_BitRate = 1
            Exit Sub
        End Select
        CFm_BitRate = Brate(BinToDec(Mid$(Bin, 17, 4)))
    End If
    'if there is a decimal place, then parse it off
    If InStr(1, CFm_BitRate, ".") Then
        CFm_BitRate = Left$(CFm_BitRate, InStr(1, CFm_BitRate, ".") - 1)
    End If
    'calculate duration
    CFm_Duration = Int((FileLen(strFileName) * 8) / CFm_BitRate / 1000)
    'Duration = Duration \ 60 & "m " & Duration - (Duration \ 60) * 60 & "s"

Exit Sub

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3 (" & strFileName & ")"
    Close FIO

End Sub

Private Function ReadID3(ByVal strFileName As String) As Boolean

  '<:-):WARNING: Large procedure (225 lines of code)
  '<:-)It is recommended that you try to break it into smaller procedures
  '<:-):WARNING: Poorly named Parameters 'FileName'
  '<:-):UPDATED: 'strFileName'
  
  Dim I           As Integer                   'holds temp int bytes'<:-):SUGGESTION: Integer could be upgraded to Long.
  Dim B           As Byte                   'holds temp byte values
  Dim s           As String                   'holds temp strings

    '<:-):WARNING: Dims with same name as a Control make code difficult to read.
    '<:-)Do NOT use Search and Replace; you will have to check each occurance
    '<:-)to determine whether it is the Control or Variable being referenced.
  Dim R           As Long
  Dim X           As Long
  Dim Version     As Byte
  Dim TagString   As String           'holds the string read from the mp3
  Dim FieldSize   As Long
  Dim SizeOffset  As Long
  Dim FieldOffset As Long
    'Dim strStrip As String
    '<:-):WARNING: Unused Dim commented out
  Dim FFile       As Long
  Dim TempString  As String
  Dim v1ID        As String * 3
  Dim v1Title     As String * 30
  Dim v1Artist    As String * 30
  Dim v1Album     As String * 30
  Dim v1Year      As String * 4
  Dim v1Comment   As String * 30
  Dim v1Genre     As Byte
  Dim v2Title     As String
  Dim v2Artist    As String
  Dim v2Album     As String
  Dim v2Year      As String
  Dim v2Genre     As String
  Dim v2Track     As String
    On Error GoTo ErrHand
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private (Encapsulted) functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'fixed length vars for reading v1ID3v1 tags
    'Dim Genre As Byte
    'vars that hold v1ID3v2 attibutes
    FFile = FreeFile
    'Open the file so we can read it
    Open strFileName For Binary As FFile
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check for an ID3v1 tag
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Get FFile, LOF(FFile) - 127, v1ID
    If v1ID = "TAG" Then 'If "TAG" is present, then we have a valv1ID v1ID3v1 tag and will extract all available v1ID3v1 info from the file
        Get FFile, , v1Title    'Always limited to 30 characters
        Get FFile, , v1Artist   'Always limited to 30 characters
        Get FFile, , v1Album    'Always limited to 30 characters
        Get FFile, , v1Year     'Always limited to 4 characters
        Get FFile, , v1Comment  'Always limited to 30 characters
        Get FFile, , v1Genre    'Always limited to 1 byte (?)
        CFm_Songname = v1Title
        CFm_Artist = v1Artist
        CFm_Album = v1Album
        CFm_Year = v1Year
        CFm_Comment = v1Comment
        CFm_Genre = v1Genre
        'get rv1ID of " " on the right sv1IDe of the string
        Do Until Right$(CFm_Songname, 1) <> " "
            CFm_Songname = Left$(CFm_Songname, Len(CFm_Songname) - 1)
        Loop
        Do Until Right$(CFm_Artist, 1) <> " "
            CFm_Artist = Left$(CFm_Artist, Len(CFm_Artist) - 1)
        Loop
        Do Until Right$(CFm_Album, 1) <> " "
            CFm_Album = Left$(CFm_Album, Len(CFm_Album) - 1)
        Loop
        Do Until Right$(CFm_Year, 1) <> " "
            CFm_Year = Left$(CFm_Year, Len(CFm_Year) - 1)
        Loop
        Do Until Right$(CFm_Comment, 1) <> " "
            CFm_Comment = Left$(CFm_Comment, Len(CFm_Comment) - 1)
        Loop
        Do Until Right$(CFm_Genre, 1) <> " "
            CFm_Genre = Left$(CFm_Genre, Len(CFm_Genre) - 1)
        Loop
        'get rv1ID of null characters
        CFm_Songname = Replace(CFm_Songname, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Songname = Replace(Songname, vbNullChar, "")
        CFm_Artist = Replace(CFm_Artist, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Artist = Replace(Artist, vbNullChar, "")
        CFm_Album = Replace(CFm_Album, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Album = Replace(Album, vbNullChar, "")
        CFm_Year = Replace(CFm_Year, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Year = Replace(Year, vbNullChar, "")
        CFm_Comment = Replace(CFm_Comment, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Comment = Replace(Comment, vbNullChar, "")
        CFm_Genre = Replace(CFm_Genre, vbNullChar, vbNullString)
        '<:-):WARNING: Empty String assignment updated to use vbNullString
        '<:-)PREVIOUS CODE : Genre = Replace(Genre, vbNullChar, "")
        'comments tend to have lots of unprintable chars, so remove them
        If Len(CFm_Comment) > 0 Then
            For X = 0 To 47
                If InStr(CFm_Comment, Chr$(X)) Then
                    CFm_Comment = Replace(CFm_Comment, Chr$(X), vbNullString)
                    '<:-):WARNING: Empty String assignment updated to use vbNullString
                    '<:-)PREVIOUS CODE : Comment = Replace(Comment, Chr$(X), "")
                End If
            Next X '<:-)For-Variable Inserted
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Check for a Header for ID3v2 tag
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Get FFile, 2, B
    If (B < 250 Or B > 251) Then
        'We have an v1ID3v2 tag
        If B = 68 Then
            Get FFile, 3, B
            If B = 51 Then
                'get version
                Get FFile, 4, Version
                'get how long the v2 tag is
                Get FFile, 7, B
                R = Val(B) * 20917152
                Get FFile, 8, B
                R = R + Val(B) * 16384
                Get FFile, 9, B
                R = R + Val(B) * 128
                Get FFile, 10, B
                R = R + Val(B)
                'invalid length was calculated
                If R > LOF(FFile) Or R > 2147483647 Then
                    Exit Function
                End If
                'get the v2 tag according to the length calculated
                TagString = Space$(R)
                Get FFile, 11, TagString
            End If
        End If
    End If
    If LenB(TagString) Then       'there is an ID3v2 tag in TagString
        '<:-):WARNING: Empty String comparision updated to use LenB()
        '<:-)PREVIOUS CODE : If TagString <> "" Then     'there is an ID3v2 tag in TagString
        ' Determine if the v1ID3v2 tag is v1ID3v2.2 or v1ID3v2.3
        Select Case Version
         Case 2 'v1ID3v2.2
            'Set the fieldnames for version 2.0
            v2Title = "TT2"
            v2Artist = "TOA"
            v2Album = "TAL"
            v2Year = "TYE"
            v2Genre = "TCO"
            FieldOffset = 7
            SizeOffset = 5
            v2Track = "TRCK"
         Case 3 'v1ID3v2.3
            'Set the fieldnames for version 3.0
            v2Title = "TIT2"
            v2Artist = "TPE1"
            v2Album = "TALB"
            v2Year = "TYER"
            v2Genre = "TCON"
            v2Track = "TRCK"
            FieldOffset = 11
            SizeOffset = 7
         Case Else
            'We don't have a valv1ID v1ID3v2 tag, so bail
            Exit Function
        End Select
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract track title
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        I = InStr(TagString, v2Title)
        If I > 0 Then
            'read the title
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) Or (B And 64) Then
                    '<:-)Pleonasm Removed
                    GoTo ReadAlbum
                End If '<:-) Structure Expanded.
            End If
            TempString = Mid$(TagString, I + FieldOffset, FieldSize)
            If LenB(TempString) Then
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : If TempString <> "" Then
                CFm_Songname = TempString
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract album title
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
        I = InStr(TagString, v2Album)
        If I > 0 Then
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) = 128 Or (B And 64) = 64 Then
                    GoTo ReadArtist
                End If '<:-) Structure Expanded.
            End If
            TempString = Mid$(TagString, I + FieldOffset, FieldSize)
            If LenB(TempString) Then
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : If TempString <> "" Then
                CFm_Album = TempString
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract artist name
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
        I = InStr(TagString, v2Artist)
        If I > 0 Then
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) = 128 Or (B And 64) = 64 Then
                    GoTo ReadYear
                End If '<:-) Structure Expanded.
            End If
            TempString = Mid$(TagString, I + FieldOffset, FieldSize)
            If LenB(TempString) Then
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : If TempString <> "" Then
                CFm_Artist = TempString
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract year title
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadYear:
        I = InStr(TagString, v2Year)
        If I > 0 Then
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) = 128 Or (B And 64) = 64 Then
                    GoTo ReadGenre
                End If '<:-) Structure Expanded.
            End If
            TempString = Mid$(TagString, I + FieldOffset, FieldSize)
            If LenB(TempString) Then
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : If TempString <> "" Then
                CFm_Year = TempString
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract genre
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadGenre:
        I = InStr(TagString, v2Genre)
        If I > 0 Then
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) = 128 Or (B And 64) = 64 Then
                    GoTo Readv2Track
                End If '<:-) Structure Expanded.
            End If
            s = Mid$(TagString, I + FieldOffset, FieldSize)
            If Left$(s, 1) = "(" Then
                If InStr(s, ")") Then
                    'the tag is in the form of "(10)metal"
                    TempString = Right$(s, Len(s) - InStr(s, ")"))
                End If
             Else 'NOT LEFT$(S,...
                TempString = s
            End If
            CFm_Genre = TempString
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Extract track number
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Readv2Track:
        I = InStr(TagString, v2Track)
        If I > 0 Then
            FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
            If Version = 3 Then
                'check for compressed or encrypted field
                B = Asc(Mid$(TagString, I + 9))
                If (B And 128) = 128 Or (B And 64) = 64 Then
                    GoTo Done
                End If '<:-) Structure Expanded.
            End If
            TempString = Mid$(TagString, I + FieldOffset, FieldSize)
            If LenB(TempString) Then
                '<:-):WARNING: Empty String comparision updated to use LenB()
                '<:-)PREVIOUS CODE : If TempString <> "" Then
                CFm_Track = TempString
            End If
        End If
Done:
    End If
    Close FFile
    'if the genre tag has a # instead of a string, then decode it
    If IsNumeric(CFm_Genre) Then
        CFm_Genre = GetGenreFromID(Genre)
    End If

Exit Function

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3 (" & strFileName & ")"
    Close FFile

End Function

Public Function ReadMP3(ByVal varFileName As Variant) As Boolean

  '<:-):WARNING: Poorly named Parameters 'FileName'
  '<:-):UPDATED: 'varFileName'

    On Error GoTo ErrHand
    'clear out mp3 info
    CFm_BitRate = ""
    CFm_Frequency = ""
    CFm_Duration = ""
    CFm_Songname = ""
    CFm_Artist = ""
    CFm_Album = ""
    CFm_Year = ""
    CFm_Comment = ""
    CFm_Genre = ""
    CFm_Track = ""
    CFm_MpegVersion = ""
    CFm_MpegLayer = ""
    CFm_VBR = False
    CFm_Mode = ""
    'reads header (duration, freq, bitrate, version, layer are stored in the header)
    ReadHeader varFileName
    'reads the ID3 tag (artist, album, year, comment, genre, track are stored in ID3 tag)
    ReadID3 varFileName
    ReadMP3 = True

Exit Function

ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description

End Function

''
''Public Function SearchDir(ByVal RootDir As String, ByVal FileType As String) As Boolean
'''<:-):WARNING: Unused Function SearchDir
'''<:-)May be a prototype Function you have not yet implimented or left over from a deleted Control.
'''Dim Dir_Names() As String
'''<:-):WARNING: Unused Dim commented out
'''Dim Num_Dirs As Integer '<:-):SUGGESTION: Integer could be upgraded to Long.
'''<:-):WARNING: Unused Dim commented out
'''Dim I As Integer '<:-):SUGGESTION: Integer could be upgraded to Long.
'''<:-):WARNING: Unused Dim commented out
''Dim fName         As String
''Dim Search_Handle As Long
''Dim File_Data     As WIN32_FIND_DATA
'''Dim TempInt As Integer '<:-):SUGGESTION: Integer could be upgraded to Long.
'''<:-):WARNING: Unused Dim commented out
''On Error GoTo ErrHand
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''DoEvents
'''clear vars because they are passed by ref and could contain data
''CFm_Subdirs = ""
''CFm_Files = ""
''' Get the matching files in this directory.
''Search_Handle = FindFirstFile(RootDir & FileType, File_Data)
''If Search_Handle <> INVALID_HANDLE_VALUE Then
''Do
''fName = File_Data.cFileName
''fName = RootDir & Left$(fName, InStr(fName, Chr$(0)) - 1)
'''add the file to the string
''CFm_Files = CFm_Files & fName & "|"
'''find the next file
''If FindNextFile(Search_Handle, File_Data) = 0 Then
''Exit Do
''End If '<:-) Structure Expanded.
''Loop
''FindClose Search_Handle
''End If
''' Get the list of subdirectories.
''Search_Handle = FindFirstFile(RootDir & "*.*", File_Data)
''If Search_Handle <> INVALID_HANDLE_VALUE Then
''Do
''If File_Data.dwFileAttributes And DDL_DIRECTORY Then
''fName = File_Data.cFileName
''fName = Left$(fName, InStr(fName, Chr$(0)) - 1)
''If fName <> "." Then
''If fName <> ".." Then
'''<:-):WARNING: Short Curcuit: 'If <condition1> And <condition2> Then' expanded
'''<:-) 'If <condition1> Then
'''<:-) If <condition2> Then '
'''<:-) Make <condition1> most likely to fail.
''CFm_Subdirs = CFm_Subdirs & fName & "|"
''End If
''End If
'''<:-)Short Circuit inserted this line
''End If
''DoEvents
''If FindNextFile(Search_Handle, File_Data) = 0 Then
''Exit Do
''End If '<:-) Structure Expanded.
''Loop
''FindClose Search_Handle
''End If
''SearchDir = True
''Exit Function
''ErrHand:
''Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"
''End Function
''
':)Roja's VB Code Fixer V1.1.78 (7-2-2004 17:26:21) 100 + 1166 = 1266 Lines Thanks Ulli for inspiration and lots of code.
