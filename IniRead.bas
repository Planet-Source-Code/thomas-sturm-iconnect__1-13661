Attribute VB_Name = "IniRead"
'Variables stored in *.ini-file: Note that (only) FTPHostname is NOT optional!
Public str_language As String   '[UserPrefs]:language=(optional)default "english" - can be set to anything specified in "lang.ini"
Public str_Caption(5) As String 'THIS IS THE STRING-ARRAY HOLDING THE "TRANSLATIONS"
Public str_Close As String      '[UserPrefs]:ExitClose=(optional)comma separated list of programs to close before shutdown
Public AutoDown As Boolean      '[UserPrefs]:AutoShutdown=(optional)0 or not present=leave computer on; >0=shutdown computer on exit
Public bol_UseNL As Boolean     '[UserPrefs]:Netlaunch=(optional)0 or not present=use build-in dial up (not recommended); >0=use Netlaunch
Public NLpath As String         '[UserPrefs]:Netlaunch=(optional)path to NetLaunch-folder; if not present, Program MUST be located in NetLaunch-folder!
Public strRes As String         '[UserPrefs]:ScreenRes=(optional)screen resolution used for connection;"800,600" (no quotes, default) works best with Desktop On Call
                                            'read Help-file for more info about possible screen resolutions
Public Server As String         '[UserPrefs]:Server=(optional)path to an application to be launched after connect
Public str_port As String       '[UserPrefs]:LocalPort=(optional) port the server listens to; default=80
Public bol_Optimize As Boolean  '[UserPrefs]:Optimize=(optional)0 or not present=don´t optimize for slow connections; >0=turn wallpaper & font smoothing off

Public ISPname As String        '[Connection]:ISP=(optional)name of ISP to be used for connection; if not present, the system default one will be used
Public FTPHostname As String    '[Connection]:FTPHostname=name of the ftp-host to connect to. Leading "ftp://" will be ignored.
Public FTPUsername As String    '[Connection]:FTPUsername=(optional)username to logon with. If omitted, will be "anonymous" (for public ftp-account)
Public FTPPassword As String    '[Connection]:FTPPassword=(optional)password to logon with. If omitted, will be "someone@home.com" (public ftps require "e-mail" adress for anonymous users).
Public FTPTimeout As String     '[Connection]:FTPTimeout=(optional)max. time to connect to server; default 1min
Public RemoteFile As String     '[Connection]:RemoteFile=(optional)file to create on the host. If omitted, will be "index.htm" (overwriting existing ones!!!).
Public Filename As String       '??
Public MaxRetries As Integer    '[Connection]:MaxRetries=(optional)number of retries to connect to host. If omitted, default is 20.

Public Interval As Integer      '[Timer]:Interval=(optional)timer interval in ms. The more, the less accurate. Default is 300ms.
Public Timer(2) As String       '[Timer]:Timer1=(optional)countdown before dial-up. Format "h:m:s". Default 0:0:5 (5s).
                                '[Timer]:Timer2=(optional)countdown before shutdown. Default 0:5:0 (5min), meaning you have 5 minutes to logon to your computer.
                                '[Timer]:Timer3=(optional)"stay-alive-timer". Default 0:10:0 (10min), makes shure you´re still connected by popping the prg up.
'Variables used internally for timing and program state
Public Times(2, 2) As Integer   'Timer array, also tells you what the prg does at the moment.
Public TimerState As Byte       'range 0,1,2. Selects timer from timer-array: Times(TimerState,Timer)
Public ConnectRetries As Integer 'number of connection-attempts already done by the program.

Public BPP, OldBPP As Integer   'screen color depth
Public ScrWidth, ScrHeight, OldWidth, OldHeight As Single 'screen resolution
Public i, j As Integer             '"j" will only be used in "Parse_Times()"
Dim str_Inifile As String       'to enable reading from lang.ini as well

Public Sub Get_ini_Values()
'[UserPrefs]
str_language = Read_ini("Iconnect.ini", "UserPrefs", "language")
If str_language = "" Then
    str_language = "english"
End If

For i = 0 To 5
    str_Caption(i) = Read_ini("language.ini", str_language, CStr(i))
Next i

str_Close = Read_ini("Iconnect.ini", "UserPrefs", "ExitClose")

If Val(Read_ini("Iconnect.ini", "UserPrefs", "AutoShutdown")) > 0 Then
    AutoDown = True
Else
    AutoDown = False
End If

If Val(Read_ini("Iconnect.ini", "UserPrefs", "UseNetLaunch")) > 0 Then
    bol_UseNL = True
    If Read_ini("Iconnect.ini", "UserPrefs", "Netlaunch") <> "" Then
        NLpath = Read_ini("Iconnect.ini", "UserPrefs", "Netlaunch")
        If Right(NLpath, 1) = "\" Then
            NLpath = NLpath & "launch.exe"
        Else
            NLpath = NLpath & "\launch.exe"
        End If
    Else
    NLpath = App.Path & "\launch.exe"
    End If
Else
    bol_UseNL = False
End If

strRes = Read_ini("Iconnect.ini", "UserPrefs", "ScreenRes")
If strRes <> "" Then
    Call Parse_strRes
Else
    ScrWidth = OldWidth
    ScrHeight = OldHeight
    BPP = OldBPP
End If

Server = Read_ini("Iconnect.ini", "UserPrefs", "Server")

str_port = Read_ini("Iconnect.ini", "UserPrefs", "LocalPort")
If Val(str_port) = 0 Then
    str_port = ""
End If

If Val(Read_ini("Iconnect.ini", "UserPrefs", "Optimize")) > 0 Then
    bol_Optimize = True
Else
    bol_Optimize = False
End If


'[Connection]
ISPname = Read_ini("Iconnect.ini", "Connection", "ISP")
FTPHostname = Read_ini("Iconnect.ini", "Connection", "FTPHostname")
If Left$(FTPHostname, 6) = "ftp://" Then
    FTPHostname = Mid$(FTPHostname, 6)
End If
FTPUsername = Read_ini("Iconnect.ini", "Connection", "FTPUsername")
If FTPUsername = "" Then
    FTPUsername = "anonymous"
End If
FTPPassword = Read_ini("Iconnect.ini", "Connection", "FTPPassword")
If FTPPassword = "" Then
    FTPPassword = "someone@home.com"
End If
FTPTimeout = Read_ini("Iconnect.ini", "Connection", "FTPTimeout")
If FTPTimeout = "" Then
    FTPTimeout = "00:01:00"
End If
RemoteFile = Read_ini("Iconnect.ini", "Connection", "RemoteFile")
If RemoteFile = "" Then
    RemoteFile = "index.htm"
End If
MaxRetries = CInt(Val(Read_ini("Iconnect.ini", "Connection", "MaxRetries")))
If MaxRetries = 0 Then
    MaxRetries = 20
End If

'[Timer]
If CInt(Val(Read_ini("Iconnect.ini", "Timer", "Interval"))) = 0 Then
    Interval = 300  'frequency of timer-events in ms: the more, the more accurate the timing gets
    j = Write_ini("Timer", "Interval", "300")
Else
    Interval = CInt(Val(Read_ini("Iconnect.ini", "Timer", "Interval")))
End If
If Read_ini("Iconnect.ini", "Timer", "Timer1") <> "" Then
    Timer(0) = Read_ini("Iconnect.ini", "Timer", "Timer1")
Else
    Timer(0) = "0:0:5"
    j = Write_ini("Timer", "Timer1", "0:0:5")
End If
If Read_ini("Iconnect.ini", "Timer", "Timer2") <> "" Then
    Timer(1) = Read_ini("Iconnect.ini", "Timer", "Timer2")
Else
    Timer(1) = "0:5:0"
    j = Write_ini("Timer", "Timer2", "0:5:0")
End If
If Read_ini("Iconnect.ini", "Timer", "Timer3") <> "" Then
    Timer(2) = Read_ini("Iconnect.ini", "Timer", "Timer3")
Else
    Timer(2) = "0:10:0"
    j = Write_ini("Timer", "Timer3", "0:10:0")
End If
Call Parse_Times
End Sub
Function Read_ini(Filename, Section, KeyName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    Read_ini = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), App.Path & "\" & Filename))
End Function
Function Write_ini(Section As String, KeyName As String, NewString As String) As Integer
    Dim r
    r = WritePrivateProfileString(Section, KeyName, NewString, App.Path & "\eyecon.ini")
End Function
Private Sub Parse_strRes()
    Dim von, bis As Long
    von = InStr(1, strRes, ",", vbTextCompare)       'find first comma in string (if present)
    bis = InStr(von + 1, strRes, ",", vbTextCompare) 'is there a second comma in our string?
    If von > 0 Then             'if "," is present, everything to the left must be ScrWidth
        ScrWidth = CSng(Val(Left$(strRes, von)))
    Else
        ScrWidth = CSng(Val(strRes))                 'no ",": strRes must only contain ScrWidth
    End If
    
    Select Case ScrWidth
        Case 640
            ScrHeight = 480
        Case 800
            ScrHeight = 600
        Case 1024
            ScrHeight = 768
        Case 1280                  'there´s two possible ScrHeight-s here: 960 and 1024
            If bis > von + 4 Then  'simplest check: 4 (ore more) digits indicate 1024
                ScrHeight = 1024
            Else
                ScrHeight = 960
            End If
        Case 1600   'to connect with that resolution REALLY makes no sense via modem, but...
            ScrHeight = 1200
        Case Else   'if none of this makes sense, play safe.
            ScrWidth = 800
            ScrHeight = 600
    End Select
    
    If bis <= Len(strRes) - 1 Then 'we´re not done yet: there might be a value for BPP set...
        BPP = CInt(Val(Right$(strRes, 2)))  'check last 2chars in strRes (32?,16?)
        If BPP = 0 Then                     'if they include the last "," (8?)
            BPP = CInt(Val(Right$(strRes, 1))) 'only check the last character
        End If
        If BPP = 0 Then                     'no useful value given for bpp
            BPP = OldBPP                    'make sure it works...
        ElseIf BPP > 0 And BPP < 12 Then    'gracefully accept everything between 1 and 11 as "8"
            BPP = 8
        ElseIf BPP >= 12 And BPP < 24 Then  'and between 12 and 23 as "16"
            BPP = 16
        Else
            BPP = 32
        End If
    Else
        BPP = OldBPP
    End If
End Sub
Private Sub Parse_Times()
Dim von, bis As Long
For j = 0 To 2
    von = 1
    i = 0
    von = InStr(von, Timer(j), ":")
    Times(j, i) = (Val(Left$(Timer(0), von)))
    For i = 1 To 2
        von = InStr(von, Timer(j), ":")
        von = von + 1
        bis = InStr(von, Timer(j), ":")
        If bis = 0 Then
            bis = Len(Timer(j)) + 1
        End If
    Times(j, i) = Val(Mid(Timer(j), von, bis - von))
    Next i
Next j
j = 0
End Sub
