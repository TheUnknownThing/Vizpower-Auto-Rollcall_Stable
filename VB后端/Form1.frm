VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ºÏ≤‚÷–°≠°≠"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   1620
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Label Label1 
      Caption         =   "ºÏ≤‚«©µΩ÷–"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim clName As String
Dim jg As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Type mType
fhwnd As Long
fText As String * 255
fRect As RECT
pHwnd As Long
pText As String * 255
End Type
Private Sub mGetAllWindow(m_Type() As mType)
Dim Wndback As Long
Dim I As Long
Do
ReDim Preserve m_Type(I)
DoEvents

m_Type(I).fhwnd = FindWindowEx(0, Wndback, vbNullString, vbNullString)
If m_Type(I).fhwnd = 0 Then
Exit Sub
Else
GetWindowText m_Type(I).fhwnd, m_Type(I).fText, 255

GetWindowRect m_Type(I).fhwnd, m_Type(I).fRect

m_Type(I).pHwnd = GetParent(m_Type(I).fhwnd)

GetWindowText m_Type(I).pHwnd, m_Type(I).pText, 255
End If
Wndback = m_Type(I).fhwnd
I = I + 1
Loop
End Sub
Sub PlayWavFile(strFileName As String, PlayCount As Long, JianGe As Long)
If Len(Dir(strFileName)) = 0 Then Exit Sub
If PlayCount = 0 Then Exit Sub
If JianGe < 1000 Then JianGe = 1000
DoEvents
sndPlaySound strFileName, 16 + 1
Sleep JianGe
Call PlayWavFile(strFileName, PlayCount - 1, JianGe)
End Sub

Private Sub Form_Load()
Dim cType() As mType
mGetAllWindow cType()
Dim lpClassName As String
Dim lhwnd As Long
Dim I As Long
Dim isVisible As Long
Dim a As RECT
Dim bili As Double
Dim wxblong As Long, wxblength As Long
Dim presswxbX As Long, presswxbY As Long
Dim prs As Long
For I = LBound(cType) To UBound(cType)
lpClassName = Space(6)
GetClassName cType(I).fhwnd, lpClassName, 255
If lpClassName = "#32770" Then
If IsWindowVisible(cType(I).fhwnd) = 1 Then
GetWindowRect cType(I).fhwnd, a
wxblong = a.Right - a.Left
wxblength = a.Bottom - a.Top
bili = wxblong / wxblength
If bili > 0.9 Then
If bili < 0.95 Then
PlayWavFile "leile.wav", 1, 1000
presswxbX = (a.Right + a.Left) / 2
presswxbY = a.Bottom - wxblength / 8
AutoPressMouse presswxbX, presswxbY
End If
End If
End If
End If
Next
End
End Sub

Private Sub AutoPressMouse(x As Long, y As Long)
SetCursorPos x, y
mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub


