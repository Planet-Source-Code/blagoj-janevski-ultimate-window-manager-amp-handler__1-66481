VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhide 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Manager"
   ClientHeight    =   6480
   ClientLeft      =   1290
   ClientTop       =   705
   ClientWidth     =   9630
   Icon            =   "frmhide.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9630
   Begin MSComctlLib.ListView lstview 
      Height          =   3735
      Left            =   4080
      TabIndex        =   20
      Top             =   2520
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Window Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Window Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Top level parent caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Top level parent class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Process"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Handle"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton command1 
      Caption         =   "Window informer"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CheckBox chktopmost 
      Caption         =   "Set top most window"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "If this is selected pressing F12 will set the window on which the mouse is on or the foreground window to top most"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkmetopmost 
      Caption         =   "Make me top most window"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CheckBox chkfore 
      Caption         =   "Foreground window"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdrefreshwnd 
      Caption         =   "Re&fresh"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkall 
      Caption         =   "Hide/Show &all windows in the system"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   3015
   End
   Begin VB.ListBox lstprocess 
      Height          =   1815
      ItemData        =   "frmhide.frx":0ECA
      Left            =   6480
      List            =   "frmhide.frx":0ECC
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdhelp 
      Caption         =   "H&elp"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please select..."
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
      Begin VB.OptionButton optshow 
         Caption         =   "&Show"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opthide 
         Caption         =   "&Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtprocname 
      Height          =   285
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5520
      Top             =   1200
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   7440
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Label lblinfownd 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Window list:"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   7560
      X2              =   7680
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   7560
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Label Label4 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Process list:"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Select the process from the list or type the process name:"
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Press CTRL-ALT-H to hide or unhide all windows of specified process, default IEXPLORE.EXE"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnutray 
      Caption         =   "wefedf2"
      Visible         =   0   'False
      Begin VB.Menu mnuabout 
         Caption         =   "Powered by Window Manager"
         Enabled         =   0   'False
      End
      Begin VB.Menu wef 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnulst 
      Caption         =   "sdfwe"
      Visible         =   0   'False
      Begin VB.Menu mnulstsetcaption 
         Caption         =   "Set caption"
      End
      Begin VB.Menu mnulstshow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnulsthide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnlsttop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnulstnotop 
         Caption         =   "Not on top"
      End
      Begin VB.Menu mnulstdestroy 
         Caption         =   "Destroy"
      End
   End
End
Attribute VB_Name = "frmhide"
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

Private Sub chkfore_Click()
If chkfore.Value = vbChecked Then Label4.Caption = "F7 button to minimize foreground window to tray." & vbNewLine & "F8 to hide the foreground window " & vbNewLine & "F9 to show all the windows hidden with F8 button." & vbNewLine & "F10 to hide all windows in the array." & vbNewLine & "F11 to hide/show me." & vbNewLine & "F12 to set/unset fore ground window to top." Else Label4.Caption = "F7 button to minimize any window to tray." & vbNewLine & "F8 to hide the window on which the cursor is." & vbNewLine & "F9 to show all the windows hidden with F8 button." & vbNewLine & "F10 to hide all windows in the array." & vbNewLine & "F11 to hide/show me." & vbNewLine & "F12 to set/unset any window to top."


End Sub

Private Sub chkmetopmost_Click()
If chkmetopmost.Value = vbChecked Then
    SetTopMostWindow frmhide.hwnd, True
Else
    SetTopMostWindow frmhide.hwnd, False
End If

End Sub

Private Sub cmdhelp_Click()
MsgBox "Be carefull when selecting show all the windows in the system." & vbNewLine & "You will have to log off or restart the computer, if you select that option. If you don't believe me try!" & vbCrLf & "Made exclusivly for PSC users. blagoj_bl@yahoo.com", vbInformation + vbApplicationModal, "Help & About"
End Sub

Private Sub cmdok_Click()
    'setup the values
    If opthide.Value = True Then
        iehidden = 0    'hide window
        opthide.Value = False
        optshow.Value = True
    Else
        iehidden = 5    'show window
        opthide.Value = True
        optshow.Value = False
    End If
    
    'default is iexplore.exe
    If Trim(txtprocname.Text) = "" Then txtprocname.Text = "iexplore.exe"
    'check chkall value
    If chkall.Value = vbChecked Then
        pname = "ALL"
        EnumWindows AddressOf EnumWindowsProc, 5
    Else
        pname = ""
        GetProcessesPids Trim(LCase(txtprocname.Text)), procpids
        'all application instances
        'e.g multiple internet explorer windows
        Dim i As Integer
        i = 1
        While procpids(i) <> -1
            pid = procpids(i)
            EnumWindows AddressOf EnumWindowsProc, 5
            i = i + 1
        Wend
    End If
   'bring to top our form
   BringWindowToTop (frmhide.hwnd)



End Sub

Private Sub cmdrefresh_Click()
lstprocess.Clear
GetProcessList frmhide.lstprocess
End Sub

Private Sub cmdrefreshwnd_Click()
lstview.ListItems.Clear
If chkall.Value = vbUnchecked Then
    pname = ""
    GetProcessesPids Trim(LCase(txtprocname.Text)), procpids
    'all application instances
    'e.g multiple internet explorer windows
    Dim i As Integer
    i = 1
    While procpids(i) <> -1
        pid = procpids(i)
        GetWindowList frmhide.lstview, lblinfownd
        i = i + 1
    Wend
Else
    pname = "ALL"
    GetWindowList frmhide.lstview, lblinfownd
End If
End Sub

Private Sub Command1_Click()
frminfo.Show
End Sub

Private Sub Form_Load()
'check for previous instances
If App.PrevInstance Then
    Dim wnd As Long
    wnd = FindWindow("ThunderRT6FormDC", frmhide.Caption)
    ShowWindow wnd, 5
    BringWindowToTop wnd
    End
End If
'hide from task list
App.Title = "" 'or you can use app.taskvisible=false
k = 1
Label4.Caption = "F7 button to minimize any window to tray." & vbNewLine & "F8 to hide the window on which the cursor is." & vbNewLine & "F9 to show all the windows hidden with F8 button." & vbNewLine & "F10 to hide all windows in the array." & vbNewLine & "F11 to hide/show me." & vbNewLine & "F12 to set/unset any window to top."
GetProcessList Me.lstprocess
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Tag = "" Then Exit Sub
Dim msg As Long
msg = X / Screen.TwipsPerPixelX
Select Case msg
    Case WM_RBUTTONDOWN
        PopupMenu mnutray
    Case WM_LBUTTONDBLCLK
        mnuopen_Click
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'show any window that was hidden in the tray
Dim twnd As Long
twnd = CLng(Mid(Me.Tag, 1, (InStr(1, Me.Tag, "#", vbTextCompare) - 1)))
ShowWindow twnd, 5
ShowAllWindows
End
End Sub
Private Sub lstprocess_Click()
txtprocname.Text = lstprocess.List(lstprocess.ListIndex)
cmdrefreshwnd_Click
End Sub


Private Sub lstview_BeforeLabelEdit(Cancel As Integer)
'disable changes
Cancel = 1
End Sub



Private Sub lstview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnulst
End Sub

Private Sub mnlsttop_Click()
Dim li As ListItem

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
SetTopMostWindow CLng(li.ListSubItems(5).Text), True


End Sub

Private Sub mnuclose_Click()
On Error Resume Next
Dim twnd As Long
Dim pos As Long
Dim traydata As NOTIFYICONDATA

Dim ticon As Long
'find the position of the separator #
pos = InStr(1, Me.Tag, "#", vbTextCompare)
'get the window handle and icon
twnd = CLng(Mid(Me.Tag, 1, pos - 1))
ticon = CLng(Mid(Me.Tag, InStr(pos + 1, Me.Tag, "#", vbTextCompare) + 1))

closewindow twnd
'form's window handle
traydata.cbSize = Len(traydata)
traydata.hwnd = CLng(Mid(Me.Tag, pos + 1, InStr(pos + 1, Me.Tag, "#", vbTextCompare) - pos - 1))
traydata.hIcon = ticon
traydata.uFlags = NIF_ICON

'Shell_NotifyIcon with NIM_DELETE
'doesn't work anyone can tell me why?
Shell_NotifyIcon NIM_DELETE, traydata

DestroyIcon (ticon)
DestroyWindow (traydata.hwnd)
UpdateTrayWindow
End Sub

Private Sub mnulstdestroy_Click()
Dim li As ListItem

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
closewindow CLng(li.ListSubItems(5).Text)
lstview.ListItems.Remove li.Index



End Sub

Private Sub mnulsthide_Click()
Dim li As ListItem

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
ShowWindow CLng(li.ListSubItems(5).Text), 0


End Sub

Private Sub mnulstnotop_Click()
Dim li As ListItem

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
SetTopMostWindow CLng(li.ListSubItems(5).Text), False

End Sub

Private Sub mnulstsetcaption_Click()
Dim li As ListItem
Dim inbox As String

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
inbox = InputBox("Write the new caption", "Set caption")
If inbox = "" Then Set li = Nothing: Exit Sub
SetWindowText CLng(li.ListSubItems(5).Text), inbox
li.Text = inbox
End Sub

Private Sub mnulstshow_Click()
Dim li As ListItem

Set li = lstview.SelectedItem
If li Is Nothing Then Exit Sub
ShowWindow CLng(li.ListSubItems(5).Text), 5

End Sub

Private Sub mnuopen_Click()
On Error Resume Next
Dim twnd As Long
Dim pos As Long
Dim traydata As NOTIFYICONDATA

Dim ticon As Long
'find the position of the separator #
pos = InStr(1, Me.Tag, "#", vbTextCompare)
'get the window handle and icon
twnd = CLng(Mid(Me.Tag, 1, pos - 1))
ticon = CLng(Mid(Me.Tag, InStr(pos + 1, Me.Tag, "#", vbTextCompare) + 1))

ShowWindow twnd, 5

'form's window handle
traydata.cbSize = Len(traydata)
traydata.hwnd = CLng(Mid(Me.Tag, pos + 1, InStr(pos + 1, Me.Tag, "#", vbTextCompare) - pos - 1))
traydata.hIcon = ticon
traydata.uFlags = NIF_ICON

'Shell_NotifyIcon with NIM_DELETE
'doesn't work anyone can tell me why?
Shell_NotifyIcon NIM_DELETE, traydata

DestroyIcon (ticon)
DestroyWindow (traydata.hwnd)
UpdateTrayWindow

End Sub



Private Sub Timer1_Timer()
'check whether is the right window, else disable the timer
If Me.Tag <> "" Then Me.Timer1.Enabled = False

'CTRL-ALT-H to hide/show all windows of specified process
If GetAsyncKeyState(17) < 0 And GetAsyncKeyState(18) < 0 And GetAsyncKeyState(72) < 0 Then
    Timer1.Enabled = False
    Call cmdok_Click
    Timer1.Enabled = True
End If

Dim pt As POINTAPI
Dim i As Long, tmp As Long, fhwnd As Long
Dim traydata As NOTIFYICONDATA
Dim wndtext As String * 256
Dim wndlen As String
Dim clslen As Long
Dim clsname As String * 260
Dim clsinfo As WNDCLASS
Dim tpid As Long, hproc As Long
Dim n As Long
Dim icon As Long

'F7 button minimize any window to tray
If GetAsyncKeyState(118) < 0 Then
    
    Timer1.Enabled = False
    If chkfore.Value = vbUnchecked Then
        GetCursorPos pt
        tmp = WindowFromPoint(pt.X, pt.Y)
    
        'is there a window or not
        If tmp = 0 Then
            Timer1.Enabled = True
            Exit Sub
        End If
        fhwnd = GetTopLevelWindow(tmp)
    Else
        fhwnd = GetForegroundWindow
        If fhwnd = 0 Then
            Timer1.Enabled = True
            Exit Sub
        End If
    End If
    
    Dim hform As Form
    Set hform = New frmhide
    
    'setup attributes for the tray icon
    'extract the icon of the exe
    GetWindowThreadProcessId fhwnd, tpid
    icon = ExtractIcon(App.hInstance, GetProcessFullPath(tpid), 0)
    If icon = 0 Then icon = frmhide.icon
    traydata.hIcon = icon
    traydata.cbSize = Len(traydata)
    traydata.uID = vbNull
    
    'send the messages to our window
    traydata.hwnd = hform.hwnd
    'we need to have the handle of the window that is in tray
    hform.Tag = CStr(fhwnd) & "#" & CStr(hform.hwnd) & "#" & CStr(icon)
    traydata.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
    'To know what window is when user clicks on the
    'icon we setup the message identifier to the handle
    'of the window
    traydata.uCallbackMessage = WM_MOUSEMOVE
    wndlen = GetWindowText(fhwnd, wndtext, 256)
    traydata.szTip = Mid(wndtext, 1, wndlen) & vbNullChar
    'add to tray menu
    Shell_NotifyIcon NIM_ADD, traydata
    ShowWindow fhwnd, 0
    Sleep 200
    Timer1.Enabled = True
End If
'F8 button, hide the window on which the mouse is on
If GetAsyncKeyState(119) < 0 Then
    Timer1.Enabled = False
    If chkfore.Value = vbUnchecked Then
    
        GetCursorPos pt
        tmp = WindowFromPoint(pt.X, pt.Y)
    
        'is there a window or not
        If tmp = 0 Then
            Timer1.Enabled = True
            Exit Sub
        End If
    Else
        tmp = GetForegroundWindow
    End If
    'we need this checking because when
    'the mouse is over the desktop
    'pressing F8 multiple times,e.g. holding it
    'will cause subscript out of range
    'because it will only add the window handle to
    'the array thwnd and won't hide the window,
    'it will hide the icons only
    If WindowIsIn(GetTopLevelWindow(tmp)) Then
        'hide the window
        ShowWindow GetTopLevelWindow(tmp), 0
        Sleep 200
        Timer1.Enabled = True
        Exit Sub
    End If
    'self protection, remove it if you want
    If GetTopLevelWindow(tmp) = frmhide.hwnd Then
        Timer1.Enabled = True
        Exit Sub
    End If

    
    'Get the top parent window
    thwnd(k) = GetTopLevelWindow(tmp)
    
    'hide window
    ShowWindow thwnd(k), 0
    k = k + 1
    'if any window was closed
    ClearClosedWindows
    Sleep 200
    Timer1.Enabled = True
End If

'F9 button, show all windows that are hidden with F8 button
If GetAsyncKeyState(120) < 0 Then ShowAllWindows

'F10 button hide all windows in the array
If GetAsyncKeyState(121) < 0 Then
    i = 1
    While thwnd(i) <> Empty
        ShowWindow thwnd(i), 0
        i = i + 1
    Wend
End If

'F11 button, hide/show frmhide
If GetAsyncKeyState(122) < 0 Then
    If frmhide.Visible = True Then
        frmhide.Hide
        Sleep 150
    Else
        frmhide.Show
        Sleep 150
    End If
End If

'F12 button,set/unset any window to top
If GetAsyncKeyState(123) < 0 Then
    If chkfore.Value = vbUnchecked Then
    
        GetCursorPos pt
        tmp = WindowFromPoint(pt.X, pt.Y)
    
        'is there a window or not
        If tmp = 0 Then
            Timer1.Enabled = True
            Exit Sub
        End If
        tmp = GetTopLevelWindow(tmp)
    Else
        tmp = GetForegroundWindow
    End If
    
    If chktopmost.Value = vbChecked Then SetTopMostWindow tmp, True Else SetTopMostWindow tmp, False
End If
End Sub



Private Sub ClearClosedWindows()
Dim i As Long, d As Long
i = 1

While thwnd(i) <> Empty
    If IsWindow(thwnd(i)) = 0 Then
        'the last window is not valid
        If thwnd(i + 1) = Empty Then
            thwnd(i) = Empty
            'window handle was destroyed
            'decrease the counter k
            k = k - 1
        Else
            'next window
            d = i + 1
            While thwnd(d) <> Empty
                'there is a window handle here
                'replace the previous not valid
                'window handle with this one
                thwnd(d - 1) = thwnd(d)
                thwnd(d) = Empty
                d = d + 1
            Wend
            'decrease the counter k
            k = k - 1
        End If
    End If
'next window
i = i + 1
Wend
End Sub
Private Function WindowIsIn(ByVal fhwnd As Long) As Boolean
Dim i As Long
i = 1
WindowIsIn = False

While thwnd(i) <> Empty
    If thwnd(i) = fhwnd Then WindowIsIn = True: Exit Function
    i = i + 1
Wend

End Function

Private Sub Timer2_Timer()
MsgBox Me.Tag
End Sub

Private Sub txtprocname_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub
'shows all windows in the array
Private Sub ShowAllWindows()
Dim i As Integer
    i = 1
    While thwnd(i) <> Empty
        ShowWindow GetTopLevelWindow(thwnd(i)), 5
        i = i + 1
    Wend
End Sub
