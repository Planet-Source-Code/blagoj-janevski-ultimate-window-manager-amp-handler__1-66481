VERSION 5.00
Begin VB.Form frminfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Informer"
   ClientHeight    =   3585
   ClientLeft      =   5925
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5850
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtprocess 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox txtparentclass 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox txtparent 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtclass 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton cmdsetcaption 
      Caption         =   "Set"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtcaption 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   4920
      Top             =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frminfo.frx":0000
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Handle:"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblwndhandle 
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbly 
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblx 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Press F11 to hide window manager."
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Process:"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Top parent class:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Top parent caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Class:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Window caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////'
' This code were explicitly developed for PSC(Planet Source Code) Users,
' as Open Source Project. This code are property of their author.
' The code is provided "as is" WITHOUT any warranty.

' You may use any of this code in you're own application(s).

' Code by Blagoj Janevski
' Please vote for me on planet-source-code.com
' e-mail: blagoj_bl@yahoo.com for comments,help or anything else.
' (c) XbXan 2006
'//////////////////////////////////////////////////////////////////////////////'

Option Explicit



Private Sub cmdsetcaption_Click()
SetWindowText lblwndhandle.Caption, Trim(txtcaption.Text)
SetWindowText GetTopLevelWindow(CLng(lblwndhandle.Caption)), Trim(txtparent.Text)
End Sub


Private Sub Command1_Click()

Dim li As ListItem
Dim j As Long

If frmhide.lstview.ListItems.Count = 0 Then Exit Sub



    For j = 1 To (frmhide.lstview.ListItems.Count)
        Set li = frmhide.lstview.ListItems.Item(j)
        If CStr(Trim(lblwndhandle.Caption)) = li.ListSubItems.Item(5).Text Then
            frmhide.lstview.SetFocus
            li.Selected = True
            li.EnsureVisible
            Exit For
        End If
    Next
    frmhide.Show
    frmhide.lstview.SetFocus
    Set li = Nothing
End Sub



Private Sub Form_Load()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when the left button is clicked then enable the timer
If Button = 1 Then Timer2.Enabled = True

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change the cursor when draging
If Button = 1 Then
        frminfo.MousePointer = 99 'custom
        frminfo.MouseIcon = Image1.Picture
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when the left button is clicked then disable the timer
If Button = 1 Then
    Timer2.Enabled = False
    frminfo.MousePointer = 0
    frminfo.MouseIcon = Nothing
End If

End Sub

Private Sub Timer2_Timer()
Dim pt As POINTAPI
Dim hwnd As Long
Dim lclass As Long
Dim clsname As String * 100
Dim wndcaption As String * 256
Dim wndparentclass As String * 100
Dim wndparentcaption As String * 256
Dim lenwnd As Long
Dim tpid As Long
Dim procname As String

'get cursor position
GetCursorPos pt
'get the window
hwnd = WindowFromPoint(pt.X, pt.Y)
lblwndhandle.Caption = hwnd
'get the window caption
lenwnd = GetWindowTextLength(hwnd)
GetWindowText hwnd, wndcaption, 256
wndcaption = Mid(wndcaption, 1, lenwnd)
'get the window class
lclass = GetClassName(hwnd, clsname, 100)
clsname = Mid(clsname, 1, lclass)
'Get top most parent window
hwnd = GetTopLevelWindow(hwnd)
'Get top most parent window caption
lenwnd = GetWindowTextLength(hwnd)
GetWindowText hwnd, wndparentcaption, 256
wndparentcaption = Mid(wndparentcaption, 1, lenwnd)
'Get top most parent window class
lclass = GetClassName(hwnd, wndparentclass, 100)
wndparentclass = Mid(wndparentclass, 1, lclass)
GetWindowThreadProcessId hwnd, tpid
procname = GetProcName(tpid)
lblx.Caption = "X: " & pt.X
lbly.Caption = "Y: " & pt.Y
txtcaption.Text = Trim(wndcaption)
txtclass.Text = Trim(clsname)
txtparent.Text = Trim(wndparentcaption)
txtparentclass.Text = Trim(wndparentclass)
txtprocess.Text = procname


End Sub

