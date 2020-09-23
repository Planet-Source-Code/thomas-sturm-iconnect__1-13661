Attribute VB_Name = "ChRes"
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    
'constants for GetDeviceCaps
Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10
'constants for EnumDisplaySettings
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
'constants for ChangeDisplaySettings
Public Const CDS_FORCE As Long = &H80000000


Public Type typDEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public DevM As typDEVMODE
Public Function ChangeRes() As Integer
    'ScrWidth As Single, ScrHeight As Single, BPP As Integer
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = ScrWidth
    DevM.dmPelsHeight = ScrHeight
    DevM.dmBitsPerPel = BPP
    'note that if CDS_FORCE is not specified, Windows will NEVER change the BPP!
    'btw. I found with my graphics adapter, the colors will screw up when switching from
    '32 to 16 Bit (esp. grey) but hey, still better than rebooting :-)
    Call ChangeDisplaySettings(DevM, CDS_FORCE)
    'we could check for return values here: they would be
    'DISP_CHANGE_SUCCESSFUL = 0 'DISP_CHANGE_RESTART = 1 'DISP_CHANGE_FAILED = -1 'DISP_CHANGE_BADMODE = -2
    'DISP_CHANGE_NOTUPDATED = -3 'DISP_CHANGE_BADFLAGS = -4 'DISP_CHANGE_BADPARAM = -5
    'but since no one should be there to notice, why bother?
End Function
