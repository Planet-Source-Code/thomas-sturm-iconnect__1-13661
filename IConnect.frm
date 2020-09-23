VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIConnect 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4455
   Icon            =   "IConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmd_Setup 
      Height          =   615
      Index           =   1
      Left            =   2280
      Picture         =   "IConnect.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmd_Help 
      Height          =   615
      Index           =   0
      Left            =   3360
      Picture         =   "IConnect.frx":0BDC
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdResetbutton 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sckWinsock2 
      Left            =   4080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckWinsock1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExitbutton 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Frame frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.Timer tmrTimer1 
         Left            =   1560
         Top             =   120
      End
      Begin VB.Label lblLabel2 
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblLabel1 
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblLabel4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblLabel3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmIConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'countdown-timer variables
Dim str_Countdown As String 'the countdown as it appears on the form
Dim dte_Now As Date         'used as a marker (ie. to display online-time in window title)
Dim bol_Stop As Boolean     'stops the countdown without disabling the timer event
Dim bol_Connected As Boolean    'prevent shutdown if program didn´t dial up itself
'variables for storing desktop/form parameters
Dim str_Wallpaper As String 'path to current windows wallpaper, gets disabled if bol_Optimize=true
Dim str_Pattern As String   'path to current background pattern, gets disabled if bol_Optimize=true
Dim str_MeCaption As String 'name of this little app
Dim lng_hwnd As Long        'window handle as used in FindWindowByTitle and GetCaptionStr
Dim dte_Online As Date      'indicate how long we are connected to the net

Dim str_WinDir As String    'may be used in later versions to check where the *.ini-file is
Dim str_COMport As String   'COM-port the modem is attached to
'ftp_sendfile()
Dim lng_x As Long           'used for "random number" port generation
Dim lng_y As Long
Dim str_ip As String        'our IP-adress
Dim str_ftpString As String 'response of ftp-server
Dim str_HTML As String      '"file" to be stored on the server
Dim lng_strpointer As Long  'cursor-like pointer into str_HTML (used in make_HTML and table_HTML)
Dim str_tempstr As String   'part of str_HTML before "cursor" lng_strpointer
Dim lng_dummy As Long       'usually used for return values we´re not keen on (Shell(blabla...))

Private Sub Form_Load()
    str_MeCaption = "Iconnect 0.9e"
    str_WinDir = Environ("Windir")
    'fill up devmode structure
    Do
        lng_dummy = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until lng_dummy = 0
    'display current screen res:
    'you could use 'OldWidth = DevM.dmPelsWidth 'OldHeight = DevM.dmPelsHeight 'OldBPP = DevM.dmBitsPerPel
    'but some graphic adapters stop EnumDisplaySettings one entry to early. This is why we use:
    OldWidth = GetDeviceCaps(hdc, HORZRES)
    OldHeight = GetDeviceCaps(hdc, VERTRES)
    OldBPP = GetDeviceCaps(hdc, BITSPIXEL)
    lblLabel3.Caption = OldWidth & " X " & OldHeight & " - " & OldBPP & " Bit"
    'display default ISP
    ISPname = Get_RegStringVal(HKEY_CURRENT_USER, "RemoteAccess", "InternetProfile")
    If ISPname = "" Then
        ISPname = Get_RegStringVal(HKEY_CURRENT_USER, "RemoteAccess", "Default")
    End If
    lblLabel4.Caption = ISPname
    
    Call Get_ini_Values 'Note: this MUST be called after EnumDisplaySettings!
    
    If bol_Optimize = True Then 'for slow connections, get current wallpaper/desktop pattern for later
        str_Wallpaper = Get_RegStringVal(HKEY_CURRENT_USER, "Control Panel\desktop", "Wallpaper")
        str_Pattern = Get_RegStringVal(HKEY_CURRENT_USER, "Control Panel\desktop", "Pattern")
    End If
        
    If IsNetConnectOnline Then
        bol_Connected = True    'program was launched manually while already online!
        bol_Optimize = False    'in this case, don´t bother deactivating wallpaper etc. even if [UserPrefs]Optimize=1
    Else
        Call CheckModem
    End If
    
    TimerState = 0              'we´re offline: first countdown ticking (before dialup)
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
    Call Set_Captions           'set label captions according to language.ini
    Call Set_Timer(TimerState, Interval)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '(TimerState=0):countdown before Dialup,(TimerState=1):countdown before connect,(TimerState=2):"keep alive" timer
    If TimerState > 0 And bol_Connected = False Then
        'disconnect using NetLaunch or wininet.dll
        If bol_UseNL = True Then
            lng_dummy = Shell(NLpath & " Trennen " & ISPname)
        Else
            Call HangUp
        End If
        Pause (1) 'wait until connection is closed, otherwise computer will not shut down
        
        'set screen resolution back to original values if we weren´t connected while program was launched
        ScrWidth = OldWidth
        ScrHeight = OldHeight
        BPP = OldBPP
        lng_dummy = ChangeRes()
        
        If bol_Optimize = True Then     'was our connection "optimized"? (see:dialup())
            If bol_fsmooth = True Then  'was font smoothing originally turned on? (then turn it back on)
                lng_dummy = SystemParametersInfo(SPI_SETFONTSMOOTHING, 1&, ByVal vbNullString, 0)
            End If                      'now set wallpaper/pattern back up again
            lng_dummy = SystemParametersInfo(SPI_SETDESKWALLPAPER, 1&, ByVal str_Wallpaper, 0)
            lng_dummy = SystemParametersInfo(SPI_SETDESKPATTERN, 1&, ByVal str_Pattern, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
        End If
        
        'shut down all programs in ExitClose-list (*.ini-file)
        lng_dummy = 0
        Do
            DoEvents
            lng_dummy = lng_dummy + 1
            lng_dummy = InStr(lng_dummy, str_Close, ",", vbTextCompare)
            If lng_dummy > 0 Then
                str_tempstr = Left(str_Close, lng_dummy - 1)
            Else
                'no (more) "," in str_close (=[UserPrefs]ExitClose)
                lng_hwnd = GetHWnd(str_tempstr)
                lng_dummy = SendMessage(lng_hwnd, WM_CLOSE, 0, 0)
                Exit Do
            End If
            'cut the first entry in str_close off (since it´s in str_tempstr now)
            str_Close = Mid$(str_Close, lng_dummy + 1)
            If str_tempstr <> "" Then
                lng_hwnd = GetHWnd(str_tempstr)
                lng_dummy = SendMessage(lng_hwnd, WM_CLOSE, 0, 0)
                ' for some apps that think they shouldn´t be closed...
                SendKeys "{ENTER}", True
            Else
                Exit Do
            End If
        Loop
        
        'shutdown computer if required (=[UserPrefs]AutoShutdown>0)
        If AutoDown = True Then
            lng_dummy = Shell(str_WinDir & "\rundll.exe user.exe,exitwindows", 1)
        End If
    End If
    
    Unload Me
    End
End Sub

Private Sub Set_Captions()
    frame1.Caption = str_Caption(0)
    lblLabel1.Caption = str_Caption(1)
    cmdExitbutton.Caption = str_Caption(2)
    cmdResetbutton.Caption = str_Caption(3)
End Sub

Private Sub Set_Timer(TimerState, Interval)
    tmrTimer1.Interval = Interval
    dte_Now = Time + TimeSerial(Times(TimerState, 0), Times(TimerState, 1), Times(TimerState, 2))
End Sub

Private Sub tmrTimer1_Timer()
    str_Countdown = CDate(Time - dte_Now)
    lblLabel2.Caption = str_Countdown

    'if we´re online and countdown gets close to the end, make shure one sees Me :-)
    If TimerState > 0 Then
        If str_Countdown <= "00:00:10" Then
            lResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2, Me.Width, Me.Height, FLAGS)
        End If
    End If

    If Time >= dte_Now Then             'str_Countdown never gets <0, that´s why.
        If TimerState = 0 Then
            Call dialup                 'this is where the fun starts...
        Else
            Call Form_Unload(0)
        End If
    End If
    
    'Note that the first timer-event makes the form visible!
    If dte_Online > "00:00:00" Then
        Me.Caption = str_MeCaption & " online: " & CDate(Time - dte_Online)
    Else
        Me.Caption = str_MeCaption
    End If
End Sub

Private Sub CheckModem()
    'get the COM-port the modem is attached to:
    str_COMport = Get_RegStringVal(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Modem\0000", "AttachedTo")
    If str_COMport = "" Then
        str_COMport = Get_RegStringVal(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Modem\0001", "AttachedTo")
    End If
    'try to open port and send something to it: if it gets send back, the modem is on
    If Not Init_Com(str_COMport, "9600,N,8,1") Then     'initialise COM-port failed: check registry!
        Unload Me
    End If
    If Write_Com("ATZ" & Chr(13)) <> 4 Then         'COM-port doesn´t send the 4 Bytes?!
        lng_dummy = Close_com                       'even when the modem´s off, you can still send via COM.
        Unload Me
    End If
    If Read_Com = "" Then                           'Modem is switched off
        lng_dummy = Close_com
        Unload Me
    End If
    lng_dummy = Close_com                           'always close COM-port, otherwise dial-up will fail
End Sub

Private Sub cmdExitbutton_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdResetbutton_click()
    'if you can click this one, you are successfully connected to your computer
    Call Set_Timer(2, 300)
    CloseWindow (Me.hwnd) 'make sure our little app doesn´t disturb until needed
End Sub

Private Sub cmd_Setup_Click(index As Integer)
    lng_dummy = Shell(str_WinDir & "\Notepad.exe " & App.Path & "\Iconnect.ini", vbNormalFocus)
    tmrTimer1.Enabled = False
    lng_dummy = ShellWait(lng_dummy)
    Call Get_ini_Values             'some settings only take effect after the program has been restartet
    tmrTimer1.Enabled = True
End Sub

Private Sub cmd_Help_Click(index As Integer)
    'write the appropriate language into the frameset "index.htm" (appearing after "mainframe" src...)
    Open App.Path & "\Docs\index.htm" For Input As #1
    Do While Not EOF(1)
       str_tempstr = str_tempstr & Input(1, #1)
    Loop
    Close #1
    lng_dummy = InStr(1, str_tempstr, "mainframe", vbTextCompare) + 15
    str_tempstr = Left$(str_tempstr, lng_dummy) & str_language & ".htm" & Mid$(str_tempstr, InStr(lng_dummy, str_tempstr, ">", vbTextCompare) - 1)
    Open App.Path & "\Docs\index.htm" For Output As #1
    Print #1, str_tempstr
    Close #1
    
    lng_dummy = Open_Browser("file:///" & App.Path & "\Docs\index.htm", Me.hwnd)
    tmrTimer1.Enabled = False
    lng_dummy = ShellWait(lng_dummy)
    tmrTimer1.Enabled = True
End Sub

Private Sub dialup()
    tmrTimer1.Enabled = False
    CloseWindow (Me.hwnd) 'make sure our little app doesn´t disturb until needed
    
    'read or make RemoteFile (="HTML-page" we want to store on the ftp-server)
    'note that the actual date, time and current ip get inserted in Make-HTML()
    If Dir(App.Path & "\" & RemoteFile) = "" Then
        str_HTML = "<HTML><BODY><FONT FACE='Arial, Helvetica'><P> local date: local time: </P><P><B>Current IP: </B><A href=http://></A></P></FONT></BODY></HTML>"
    Else
        Open RemoteFile For Input As #1
        Do While Not EOF(1)
            str_HTML = str_HTML & Input(1, #1)
        Loop
        Close #1
    End If
    
    If bol_Connected = False Then
        If bol_UseNL = True Then
            lng_dummy = Shell(NLpath & " " & ISPname)
        Else
            lng_dummy = InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0)
        End If
        If lng_dummy > 0 Then
            'something went wrong trying to dial-up:
            TimerState = 1      'do as if we were connected...
            Call Form_Unload(0) 'and quit!
        End If
        'ScrWidth, ScrHeight, BPP are declared in *.ini-file ([UserPrefs]ScreenRes)
        lng_dummy = ChangeRes()
        lblLabel3.Caption = DevM.dmPelsWidth & " X " & DevM.dmPelsHeight & " - " & DevM.dmBitsPerPel & " Bit"
    End If
    
    If bol_Optimize = True Then 'do we want to go all the way to get a fast connection?
        lng_dummy = SystemParametersInfo(SPI_GETFONTSMOOTHING, 0, lResult, 0) 'is font-smoothing turned on?
        bol_fsmooth = CBool(lResult)
        If bol_fsmooth = True Then 'if it is, then turn it off
            lng_dummy = SystemParametersInfo(SPI_SETFONTSMOOTHING, 0, ByVal vbNullString, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
        End If
        lng_dummy = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", 0) 'disable wallpaper and desktop pattern
        lng_dummy = SystemParametersInfo(SPI_SETDESKPATTERN, 0, "", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
    End If
      
    dte_Online = Time
    'this is the loop that waits for the connection
    Do While Not IsNetConnectOnline
        DoEvents
        'if dialing takes more than the time specified in timer(1) (default: 5´), shutdown system
        If Time > dte_Online + CDate(Timer(1)) Then
            TimerState = 1
            Call Form_Unload(0)
        End If
    Loop
    dte_Online = Time   'display the time we´re online in the caption of the form
    TimerState = 1      'we´re going online: prepare 2nd countdown...
    Call ftp_Sendfile
    
    lng_dummy = ShowWindow(Me.hwnd, 1)
    cmdExitbutton.Caption = str_Caption(4)
    frame1.Caption = str_Caption(5)
    cmdExitbutton.Default = False
    cmdExitbutton.Font.Bold = False
    cmdResetbutton.Visible = True
    cmdResetbutton.Default = True
    cmdResetbutton.Font.Bold = True
    cmdResetbutton.SetFocus
    
    
    If bol_Connected = False Then 'don´t launch external programs if we didn´t connect ourselves
        If Server = "" Then
            lng_dummy = Shell(Get_RegStringVal(HKEY_LOCAL_MACHINE, "SOFTWARE\IBM\Desktop On-Call\2.0", "install_path") & "\sessprop.exe")
        Else
            lng_dummy = Shell(Server)
        End If
    End If
    
    Call Set_Timer(TimerState, Interval) '2nd countdown (timer(1)) ticking...
    tmrTimer1.Enabled = True             'note that cmdExitbutton_Click() now really shuts down your system!
End Sub

Public Sub ftp_Sendfile()
    dte_Now = Time + CDate(FTPTimeout)
    'based upon code by Kristian Trenskow, found at http://www.planetsourcecode.com
    sckWinsock2.RemoteHost = FTPHostname        'from *.ini-file
    sckWinsock2.RemotePort = 21                 'ftp-server listen to port 21
    sckWinsock2.Connect
    Do Until sckWinsock2.State = sckConnected   'wait for command conection to be established
        DoEvents
        If Time >= dte_Now Then
            Call Form_Unload(0)
        End If
        'due to a bug in wininet.dll: it sometimes won´t recognise a connetion and tries to connect itself
        'since this happens in a loop, don´t wait for it (wininet.dll) to notice
        SendKeys "{ENTER}", False
    Loop
    sckWinsock2.SendData "USER " & FTPUsername & vbCrLf 'FTPUsername from IniRead
    str_ftpString = ""
    Do Until str_ftpString <> ""                'wait ´til server noticed
        DoEvents
        If Time >= dte_Now Then
            Call Form_Unload(0)
        End If
    Loop
    sckWinsock2.SendData "PASS " & FTPPassword & vbCrLf 'FTPPassword from IniRead
    str_ftpString = ""
    Do Until str_ftpString <> ""                'got everything ?
        DoEvents
        If Time >= dte_Now Then
            Call Form_Unload(0)
        End If
    Loop
    Randomize
    lng_x = Int(12 * Rnd + 5)      'find two random numbers to specify port the server connects to
    Randomize
    lng_y = Int(254 * Rnd + 1)
    
    str_ip = sckWinsock2.LocalIP   'Note: sckWinsock1.LocalIP may contain your LAN-IP!
    lblLabel4.Caption = str_ip     'i.e. the IP of your actual network-adapter
    Call Make_HTML
    
    Do Until InStr(str_ip, ".") = 0 'replace every "." in IP With a ","
        str_ip = Mid(str_ip, 1, InStr(str_ip, ".") - 1) & "," & Mid(str_ip, InStr(str_ip, ".") + 1)
    Loop
    sckWinsock2.SendData "PORT " & str_ip & "," & Trim(Str(lng_x)) & "," & Trim(Str(lng_y)) & vbCrLf 'tell the server With which IP he has to connect and with which port
    str_ftpString = ""
    Do Until str_ftpString <> ""         'wait until server responds
        DoEvents
        If Time >= dte_Now Then
            Call Form_Unload(0)
        End If
    Loop
    sckWinsock1.Close
    sckWinsock1.LocalPort = lng_x * 256 Or lng_y ' set port of second winsock-control to the port the server will connect to
    ' lng_x is the most-significant byte of the port number, lng_x is the least significant byte. To find the port, you have to move every
    ' bit 8 places to the right (or multiply with 256). Then compare every bit with the bits of lng_y, using OR
    sckWinsock1.Listen          'listen For the FTP-Server to connect
    sckWinsock2.SendData "STOR " & RemoteFile & vbCrLf 'Store a file, With RETR you can Get a file, with LIST you get a list of all file on the server, all this information is sent through the data-connection (to change directory use CWD)
    Pause (1)
    dte_Now = Time
    Do Until sckWinsock1.State = sckConnected   'wait until the FTP-Server connects
        DoEvents
        'reconnect if it takes more than some seconds (timer(0) defaults to 5s) OR connection fails (state=9)
        If Time > dte_Now + CDate(Timer(0)) Or sckWinsock1.State = 9 Then
            Call Retry_connect
        End If
    Loop
    Pause (1) 'wait a little bit, because the server needs a moment (don't know how, but it only works like that)
    If sckWinsock1.State = sckConnected Then
        sckWinsock1.SendData str_HTML 'send some data, the FTP-Server will store it in the file.
        'Send only ASCII data, If you send Binary you have to tell it the server before, use Type to Do this
    End If
    Pause (1)
    sckWinsock1.Close 'close data-connection
    Pause (1)
    sckWinsock2.Close 'you don't have to close the connection here, you also can transfer another file
End Sub

Private Sub Retry_connect()
    sckWinsock1.Close
    Pause (1)
    sckWinsock2.Close
    Pause (1)
    ConnectRetries = ConnectRetries + 1
    lblLabel4.Caption = "Re-Connect: " & ConnectRetries
    If ConnectRetries <= MaxRetries Then
        Call ftp_Sendfile
    Else
        Call cmdExitbutton_Click
    End If
End Sub

Private Sub Make_HTML()
    'in this sub, we check for the "key-words" <local date>, <local time> and <current-ip> to insert
    'the correct values after them. Our str_HTML should then read:
    'str_HTML = "<HTML><BODY><FONT FACE='Arial, Helvetica'><P>local date: " & Date & "  local time: " & Time & "</P><P><B>Current IP: </B><A href=http://" & str_ip & ">" & str_ip & "</A></P></FONT></BODY></HTML>"
    lng_strpointer = InStr(1, str_HTML, "local date:", vbTextCompare)
    If lng_strpointer > 0 Then
        lng_strpointer = lng_strpointer + 11                          'len("local date: ")=11
        str_tempstr = Left(str_HTML, lng_strpointer) & Date           'insert date after " "
        If IsDate(Mid(str_HTML, lng_strpointer + 1, 8)) Then          'if there already was a date
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer + 9) 'leave it out of the string
        Else
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer)   'but copy to str_HTML before " "
        End If
    End If
    lng_strpointer = InStr(1, str_HTML, "local time:", vbTextCompare)
    If lng_strpointer > 0 Then
        lng_strpointer = lng_strpointer + 11                          'len("local time: ")=11
        str_tempstr = Left(str_HTML, lng_strpointer) & Time           'insert to str_HMTL
        If IsDate(Mid(str_HTML, lng_strpointer + 1, 8)) Then          'if there already was a time
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer + 9) 'leave it out of the string
        Else
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer)   'but copy to str_HTML before " "
        End If
    End If
    lng_strpointer = InStr(1, str_HTML, "current ip:", vbTextCompare) 'search for first occurance of ip
    If lng_strpointer > 0 Then                                        'now find the next hyperlink after that
        lng_strpointer = InStr(lng_strpointer, str_HTML, "://")
        If lng_strpointer > 0 Then
            lng_strpointer = lng_strpointer + 2                       '(a href=:"http://...):len("//")=2
            If str_port = "" Then
                str_tempstr = Left$(str_HTML, lng_strpointer) & str_ip
            Else
                str_tempstr = Left$(str_HTML, lng_strpointer) & str_ip & ":" & str_port
            End If
            lng_dummy = 1
            If IsNumeric(Mid$(str_HTML, lng_strpointer + 1, 1)) Then  'was there already an IP?
                'of course len(str_ip) delivers the length of the currentIP, but we need the previous one!
                Call getlen_ip
            End If
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer + lng_dummy)   'ip inserted in link
            lng_strpointer = InStr(lng_strpointer + Len(str_ip), str_HTML, ">")
            str_tempstr = Left$(str_HTML, lng_strpointer) & str_ip
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer + lng_dummy)
        End If
    End If
    
    str_tempstr = ""
    Call table_HTML
    If str_tempstr <> "" Then
        'we found a table in str_HTML: means, we loaded it from a file
        Call store_HTML
    End If
    
End Sub

Private Sub getlen_ip()
    Do
        'check every character in str_HTML after lng_strpointer: if it´s ">" then IP is one char shorter!
        If Mid$(str_HTML, lng_strpointer + lng_dummy, 1) = ">" Then
            lng_dummy = lng_dummy - 1
            Exit Do
        Else
            'if its not ">" countinue counting
            lng_dummy = lng_dummy + 1
        End If
    Loop
End Sub
Private Sub table_HTML()
    'will fill up a table (if there is any) on the HTML-page from top to bottom with date/time/ip
    lng_strpointer = InStr(lng_strpointer, str_HTML, "<tr>", vbTextCompare) 'find beginning of first (next) table row
    If lng_strpointer > 0 Then
        lng_strpointer = InStr(lng_strpointer, str_HTML, "<td>", vbTextCompare) + 4 'first table cell after that
        If Mid$(str_HTML, lng_strpointer, 1) = "<" Then 'is empty? (then assume all other cells of that row are, too)
            str_tempstr = Left(str_HTML, lng_strpointer - 1) & Date                      'insert date
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer)
            lng_strpointer = InStr(lng_strpointer, str_HTML, "<td>", vbTextCompare) + 4   'find next cell
            str_tempstr = Left(str_HTML, lng_strpointer - 1) & Time                      'insert time
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer)
            lng_strpointer = InStr(lng_strpointer, str_HTML, "<td>", vbTextCompare) + 4   'find next cell
            str_tempstr = Left(str_HTML, lng_strpointer - 1) & str_ip                    'insert ip
            str_HTML = str_tempstr & Mid$(str_HTML, lng_strpointer)
        Else
            ' this row is not empty: look for the next one
            Call table_HTML
        End If
    Else
        'lng_strpointer=0: there aren´t any rows left in our table
        Exit Sub
    End If
End Sub

Private Sub store_HTML()
    Open RemoteFile For Output As #1
    Print #1, str_HTML
    Close #1
End Sub
Private Sub Pause(Seconds)
    dte_Now = Time
    Do Until Time > dte_Now + Seconds / 86400
        DoEvents
    Loop
End Sub

Private Sub sckWinsock1_ConnectionRequest(ByVal requestID As Long)
    sckWinsock1.Close
    sckWinsock1.Accept requestID
End Sub

Private Sub sckWinsock1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    sckWinsock1.GetData data
    Debug.Print data
    sckWinsock1.Close 'you have to close the connection after the Server had send you data, he will establish it again, when he sends more
    sckWinsock1.Listen
End Sub

Private Sub sckWinsock2_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    sckWinsock2.GetData data
    Debug.Print data
    str_ftpString = data 'store data
End Sub

