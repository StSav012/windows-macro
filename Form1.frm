VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Windows Macro"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2775
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "000000 x 000000"
      ToolTipText     =   "Hit Ctrl+Enter to add a Move event"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":030A
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   2160
      Picture         =   "Form1.frx":031F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   1560
      Picture         =   "Form1.frx":0629
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Add from File"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      Picture         =   "Form1.frx":0933
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Stop"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":0C3D
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Start"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Cursor Position"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type OPENFILENAME
    lStructSize As Long          'The size of this struct (Use the Len function)
    hwndOwner As Long            'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    flags As Long                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const OFN_ALLOWMULTISELECT = &H200 ' The File Name list box allows multiple selections. If you also set the OFN_EXPLORER flag, the dialog box uses the Explorer-style user interface; otherwise, it uses the old-style user interface.
Private Const OFN_CREATEPROMPT = &H2000      ' If the user specifies a file that does not exist, this flag causes the dialog box to prompt the user for permission to create the file. If the user chooses to create the file, the dialog box closes and the function returns the specified name; otherwise, the dialog box remains open. If you use this flag with the OFN_ALLOWMULTISELECT flag, the dialog box allows the user to specify only one nonexistent file.
Private Const OFN_DONTADDTORECENT = &H2000000 ' Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the SHGetSpecialFolderLocation function with the CSIDL_RECENT flag.
Private Const OFN_ENABLEHOOK = &H20       ' Enables the hook function specified in the lpfnHook member.
Private Const OFN_ENABLEINCLUDENOTIFY = &H400000   ' Causes the dialog box to send CDN_INCLUDEITEM notification messages to your OFNHookProc hook procedure when the user opens a folder. The dialog box sends a notification for each item in the newly opened folder. These messages enable you to control which items the dialog box displays in the folder's item list.
Private Const OFN_ENABLESIZING = &H800000   ' Enables the Explorer-style dialog box to be resized using either the mouse or the keyboard. By default, the Explorer-style Open and Save As dialog boxes allow the dialog box to be resized regardless of whether this flag is set. This flag is necessary only if you provide a hook procedure or custom template. The old-style dialog box does not permit resizing.
Private Const OFN_ENABLETEMPLATE = &H40       ' The lpTemplateName member is a pointer to the name of a dialog template resource in the module identified by the hInstance member. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
Private Const OFN_ENABLETEMPLATEHANDLE = &H80       ' The hInstance member identifies a data block that contains a preloaded dialog box template. The system ignores lpTemplateName if this flag is specified. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
Private Const OFN_EXPLORER = &H80000    ' Indicates that any customizations made to the Open or Save As dialog box use the Explorer-style customization methods. For more information, see Explorer-Style Hook Procedures and Explorer-Style Custom Templates.
Private Const OFN_EXTENSIONDIFFERENT = &H400      ' The user typed a file name extension that differs from the extension specified by lpstrDefExt. The function does not use this flag if lpstrDefExt is NULL.
Private Const OFN_FILEMUSTEXIST = &H1000     ' The user can type only names of existing files in the File Name entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used. This flag can be used in an Open dialog box. It cannot be used with a Save As dialog box.
Private Const OFN_FORCESHOWHIDDEN = &H10000000 ' Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown.
Private Const OFN_HIDEREADONLY = &H4        ' Hides the Read Only check box.
Private Const OFN_LONGNAMES = &H200000   ' For old-style dialog boxes, this flag causes the dialog box to use long file names. If this flag is not specified, or if the OFN_ALLOWMULTISELECT flag is also set, old-style dialog boxes use short file names (8.3 format) for file names with spaces. Explorer-style dialog boxes ignore this flag and always display long file names.
Private Const OFN_NODEREFERENCELINKS = &H100000   ' Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut.
Private Const OFN_NOLONGNAMES = &H40000    ' For old-style dialog boxes, this flag causes the dialog box to use short file names (8.3 format). Explorer-style dialog boxes ignore this flag and always display long file names.
Private Const OFN_NONETWORKBUTTON = &H20000    ' Hides and disables the Network button.
Private Const OFN_NOREADONLYRETURN = &H8000     ' The returned file does not have the Read Only check box selected and is not in a write-protected directory.
Private Const OFN_NOTESTFILECREATE = &H10000    ' The file is not created before the dialog box is closed. This flag should be specified if the application saves the file on a create-nonmodify network share. When an application specifies this flag, the library does not check for write protection, a full disk, an open drive door, or network protection. Applications using this flag must perform file operations carefully, because a file cannot be reopened once it is closed.
Private Const OFN_NOVALIDATE = &H100      ' The common dialog boxes allow invalid characters in the returned file name. Typically, the calling application uses a hook procedure that checks the file name by using the FILEOKSTRING message. If the text box in the edit control is empty or contains nothing but spaces, the lists of files and directories are updated. If the text box in the edit control contains anything else, nFileOffset and nFileExtension are set to values generated by parsing the text. No default extension is added to the text, nor is text copied to the buffer specified by lpstrFileTitle. If the value specified by nFileOffset is less than zero, the file name is invalid. Otherwise, the file name is valid, and nFileExtension and nFileOffset can be used as if the OFN_NOVALIDATE flag had not been specified.
Private Const OFN_OVERWRITEPROMPT = &H2        ' Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
Private Const OFN_PATHMUSTEXIST = &H800      ' The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the File Name entry field, the dialog box function displays a warning in a message box.
Private Const OFN_READONLY = &H1      ' Causes the Read Only check box to be selected initially when the dialog box is created. This flag indicates the state of the Read Only check box when the dialog box is closed.
Private Const OFN_SHAREAWARE = &H4000  ' Specifies that if a call to the OpenFile function fails because of a network sharing violation, the error is ignored and the dialog box returns the selected file name. If this flag is not set, the dialog box notifies your hook procedure when a network sharing violation occurs for the file name specified by the user. If you set the OFN_EXPLORER flag, the dialog box sends the CDN_SHAREVIOLATION message to the hook procedure. If you do not set OFN_EXPLORER, the dialog box sends the SHAREVISTRING registered message to the hook procedure.
Private Const OFN_SHOWHELP = &H10      ' Causes the dialog box to display the Help button.

Private Sub Command1_Click()
    On Error Resume Next
    Dim n As Long, n_ As Long, i As Long
    Dim c As String, l As String
    Dim p1 As Long, p2 As Long
    n_ = 1
    Command1.Enabled = False
    Command3.Enabled = True
    Text1.Locked = True
    If Right(Text1.Text, 2) <> vbCrLf Then Text1.Text = Text1.Text + vbCrLf
    Do
        n = InStr(n_, Text1.Text, vbCrLf)
        ' Check whether Scroll Lock is OFF
        If GetKeyState(vbKeyScrollLock) = 0 Then
            If n > 0 Then
                l = Trim(Mid(Text1.Text, n_, n - n_))
            Else
                l = Trim(Mid(Text1.Text, n_))
            End If
            If l <> "" And l <> vbCrLf And Left(l, 1) <> "#" Then
                n_ = InStr(1, l, " ")
                If n_ > 0 Then
                    c = LCase(Left(l, n_ - 1))
                    l = Mid(l, n_ + 1)
                Else
                    c = LCase(l)
                    l = ""
                End If
                Select Case c
                    Case "beep", "sound"
                        n_ = InStr(1, l, " ")
                        If n_ = 0 Then
                            Interaction.Beep
                        Else
                            p1 = CLng(Left(l, n_ - 1))
                            p2 = CLng(Mid(l, n_ + 1))
                            Beep p1, p2
                        End If
                    Case "mv", "move", "moveto"
                        n_ = InStr(1, l, " ")
                        If n_ = 0 Then
                            MsgBox "Too few parameters passed to" & vbCrLf & c, vbExclamation
                        Else
                            p1 = CLng(Left(l, n_ - 1))
                            p2 = CLng(Mid(l, n_ + 1))
                            SetCursorPos p1, p2
                        End If
                    Case "sleep", "wait", "delay"
                        If l = "" Then
                            MsgBox "Too few parameters passed to" & vbCrLf & c, vbExclamation
                        Else
                            n_ = InStr(1, l, " ")
                            If n_ = 0 Then
                                p1 = CLng(l)
                                p2 = 1
                            Else
                                p1 = CLng(Left(l, n_ - 1))
                                p2 = CLng(Mid(l, n_ + 1))
                            End If
                            For i = 1 To p1 \ 100
                                SleepEx 100, p2
                                DoEvents
                                If Not Command3.Enabled Then Exit For
                            Next i
                            If Command3.Enabled Then SleepEx p1 Mod 100, p2
                        End If
                    Case "click"
                        If l = "" Then
                            p1 = 1
                        Else
                            p1 = CLng(l)
                        End If
                        Select Case p1
                            Case 1
                                mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
                                DoEvents
                                mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                            Case 2
                                mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
                                DoEvents
                                mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
                            Case 3
                                mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
                                DoEvents
                                mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
                            Case Else
                                MsgBox "Incorrect parameter value in" & vbCrLf & c, vbExclamation
                        End Select
                    Case "echo", "type"
                        SendKeys l
                    Case "press", "presskey"
                        If l = "" Then
                            MsgBox "Too few parameters passed to" & vbCrLf & c, vbExclamation
                        Else
                            l = LCase(l)
                            If Asc(l) > 96 And Asc(l) < 123 Then
                                Select Case l
                                    Case "backspace", "bksp"
                                        p1 = vbKeyBack
                                    Case "tab"
                                        p1 = vbKeyTab
                                    Case "enter", "return"
                                        p1 = vbKeyReturn
                                    Case "shift"
                                        p1 = vbKeyShift
                                    Case "ctrl", "control"
                                        p1 = vbKeyControl
                                    Case "alt", "menu"
                                        p1 = vbKeyMenu
                                    Case "esc", "escape"
                                        p1 = vbKeyEscape
                                    Case "space", " "
                                        p1 = vbKeySpace
                                    Case "page up", "pgup"
                                        p1 = vbKeyPageUp
                                    Case "page down", "pgdn"
                                        p1 = vbKeyPageDown
                                    Case "end"
                                        p1 = vbKeyEnd
                                    Case "home"
                                        p1 = vbKeyHome
                                    Case "left", "left arrow", "arrow left"
                                        p1 = vbKeyLeft
                                    Case "up", "up arrow", "arrow up"
                                        p1 = vbKeyUp
                                    Case "right", "right arrow", "arrow right"
                                        p1 = vbKeyRight
                                    Case "down", "down arrow", "arrow down"
                                        p1 = vbKeyDown
                                    Case "print screen", "prt scr", "prt sc", "snapshot", "screenshot", "sshot", "shot"
                                        p1 = vbKeySnapshot
                                    Case "delete", "del"
                                        p1 = vbKeyDelete
                                    Case Else
                                        MsgBox "Invalid argument value in" & vbCrLf & c, vbExclamation
                                        p1 = 0
                                End Select
                            Else
                                p1 = CLng(l)
                            End If
                            keybd_event p1, MapVirtualKey(p1, 0), KEYEVENTF_EXTENDEDKEY, 0
                            keybd_event p1, MapVirtualKey(p1, 0), KEYEVENTF_KEYUP + KEYEVENTF_EXTENDEDKEY, 0
                        End If
                    Case "keydown", "keydn"
                        If l = "" Then
                            MsgBox "Too few parameters passed to" & vbCrLf & c, vbExclamation
                        Else
                            l = LCase(l)
                            If Asc(l) > 96 And Asc(l) < 123 Then
                                Select Case l
                                    Case "backspace", "bksp"
                                        p1 = vbKeyBack
                                    Case "tab"
                                        p1 = vbKeyTab
                                    Case "enter", "return"
                                        p1 = vbKeyReturn
                                    Case "shift"
                                        p1 = vbKeyShift
                                    Case "ctrl", "control"
                                        p1 = vbKeyControl
                                    Case "alt", "menu"
                                        p1 = vbKeyMenu
                                    Case "esc", "escape"
                                        p1 = vbKeyEscape
                                    Case "space", " "
                                        p1 = vbKeySpace
                                    Case "page up", "pgup"
                                        p1 = vbKeyPageUp
                                    Case "page down", "pgdn"
                                        p1 = vbKeyPageDown
                                    Case "end"
                                        p1 = vbKeyEnd
                                    Case "home"
                                        p1 = vbKeyHome
                                    Case "left", "left arrow", "arrow left"
                                        p1 = vbKeyLeft
                                    Case "up", "up arrow", "arrow up"
                                        p1 = vbKeyUp
                                    Case "right", "right arrow", "arrow right"
                                        p1 = vbKeyRight
                                    Case "down", "down arrow", "arrow down"
                                        p1 = vbKeyDown
                                    Case "print screen", "prt scr", "prt sc", "snapshot", "screenshot", "sshot", "shot"
                                        p1 = vbKeySnapshot
                                    Case "delete", "del"
                                        p1 = vbKeyDelete
                                    Case Else
                                        MsgBox "Invalid argument value in" & vbCrLf & c, vbExclamation
                                        p1 = 0
                                End Select
                            Else
                                p1 = CLng(l)
                            End If
                            keybd_event p1, MapVirtualKey(p1, 0), KEYEVENTF_EXTENDEDKEY, 0
                        End If
                    Case "keyup"
                        If l = "" Then
                            MsgBox "Too few parameters passed to" & vbCrLf & c, vbExclamation
                        Else
                            l = LCase(l)
                            If Asc(l) > 96 And Asc(l) < 123 Then
                                Select Case l
                                    Case "backspace", "bksp"
                                        p1 = vbKeyBack
                                    Case "tab"
                                        p1 = vbKeyTab
                                    Case "enter", "return"
                                        p1 = vbKeyReturn
                                    Case "shift"
                                        p1 = vbKeyShift
                                    Case "ctrl", "control"
                                        p1 = vbKeyControl
                                    Case "alt", "menu"
                                        p1 = vbKeyMenu
                                    Case "esc", "escape"
                                        p1 = vbKeyEscape
                                    Case "space", " "
                                        p1 = vbKeySpace
                                    Case "page up", "pgup"
                                        p1 = vbKeyPageUp
                                    Case "page down", "pgdn"
                                        p1 = vbKeyPageDown
                                    Case "end"
                                        p1 = vbKeyEnd
                                    Case "home"
                                        p1 = vbKeyHome
                                    Case "left", "left arrow", "arrow left"
                                        p1 = vbKeyLeft
                                    Case "up", "up arrow", "arrow up"
                                        p1 = vbKeyUp
                                    Case "right", "right arrow", "arrow right"
                                        p1 = vbKeyRight
                                    Case "down", "down arrow", "arrow down"
                                        p1 = vbKeyDown
                                    Case "print screen", "prt scr", "prt sc", "snapshot", "screenshot", "sshot", "shot"
                                        p1 = vbKeySnapshot
                                    Case "delete", "del"
                                        p1 = vbKeyDelete
                                    Case Else
                                        MsgBox "Invalid argument value in" & vbCrLf & c, vbExclamation
                                        p1 = 0
                                End Select
                            Else
                                p1 = CLng(l)
                            End If
                            keybd_event p1, MapVirtualKey(p1, 0), KEYEVENTF_KEYUP + KEYEVENTF_EXTENDEDKEY, 0
                        End If
                    Case "alert"
                        MsgBox l, vbSystemModal + vbExclamation
                    Case "info"
                        MsgBox l, vbSystemModal + vbInformation
                    Case "launch", "start", "run", "cmd"
                        Shell l, vbNormalFocus
                    Case "end"
                        n = 0
                    Case Else
                        MsgBox "Unknown command:" & vbCrLf & c, vbSystemModal + vbExclamation
                End Select
            End If
            n_ = n + 2
        Else
            SleepEx 20, 1
        End If
        DoEvents
    Loop While n > 0 And Command3.Enabled
    Text1.Locked = False
    Command1.Enabled = True
    Command3.Enabled = False
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
End Sub

Private Sub Command4_Click()
    Dim o As OPENFILENAME
    o.hwndOwner = hWnd
    'o.hInstance = 0
    o.lpstrFilter = "All Files" & vbNullChar & "*"
    o.lpstrCustomFilter = vbNullString
    o.nMaxCustFilter = 0
    o.nFilterIndex = 0
    o.lpstrFile = String(1024, 0)
    o.nMaxFile = 1023
    o.lpstrFileTitle = String(1024, 0)
    o.nMaxFileTitle = 1023
    o.lpstrInitialDir = vbNullString
    o.lpstrTitle = vbNullString
    o.flags = OFN_DONTADDTORECENT + OFN_ENABLESIZING + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_NONETWORKBUTTON + OFN_HIDEREADONLY
    'o.nFileOffset
    'o.nFileExtension
    o.lpstrDefExt = vbNullString
    o.lCustData = 0
    'o.lpfnHook
    'o.lpTemplateName
    o.lStructSize = Len(o)
    Dim s As String, s_ As String, p As Long
    If GetOpenFileName(o) <> 0 Then
        Open o.lpstrFile For Input As #1
        s_ = vbNullString
        While Not EOF(1)
            Line Input #1, s
            s_ = s_ & s
            If Not EOF(1) Then s_ = s_ & vbCrLf
        Wend
        Close
        p = Text1.SelStart
        Text1.Text = Left(Text1.Text, Text1.SelStart) & s_ & Mid(Text1.Text, Text1.SelStart + Text1.SelLength + 1)
        Text1.SelStart = p + Len(s_)
    End If
End Sub

Private Sub Command5_Click()
    Dim s As OPENFILENAME
    s.hwndOwner = hWnd
    's.hInstance = 0
    s.lpstrFilter = "All Files" & vbNullChar & "*"
    s.lpstrCustomFilter = vbNullString
    s.nMaxCustFilter = 0
    s.nFilterIndex = 0
    s.lpstrFile = String(1024, 0)
    s.nMaxFile = 1023
    s.lpstrFileTitle = String(1024, 0)
    s.nMaxFileTitle = 1023
    s.lpstrInitialDir = vbNullString
    s.lpstrTitle = vbNullString
    s.flags = OFN_DONTADDTORECENT + OFN_ENABLESIZING + OFN_EXPLORER + OFN_NONETWORKBUTTON + OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
    's.nFileOffset
    's.nFileExtension
    s.lpstrDefExt = vbNullString
    s.lCustData = 0
    's.lpfnHook
    's.lpTemplateName
    s.lStructSize = Len(s)
    If GetSaveFileName(s) <> 0 Then
        Open s.lpstrFile For Output As #1
        Print #1, Text1.Text
        Close
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim p As Long, s As String
    If KeyCode = vbKeyReturn And Shift = 2 Then
        Form1.SetFocus
        p = Text1.SelStart
        s = "Move " & Replace(Text2.Text, " x", vbNullString)
        If Text1 <> ActiveControl Then s = s & vbCrLf
        Text1.Text = Left(Text1.Text, Text1.SelStart) & _
                     s & _
                     Mid(Text1.Text, Text1.SelStart + Text1.SelLength + 1)
        Text1.SelStart = p + Len(s)
        KeyCode = 0
        Text1.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
    If ScaleWidth < 2775 Then Width = Width + 2775 - ScaleWidth
    If ScaleHeight < 3135 Then Height = Height + 3135 - Height
    Label1.Top = ScaleHeight - 120 - Label1.Height
    Text2.Top = Label1.Top
    Text2.Left = ScaleWidth - 120 - Text2.Width
    Text1.Width = ScaleWidth - 240
    Text1.Height = Label1.Top - 240 - Command1.Height - Command1.Top
    Command5.Left = ScaleWidth - 120 - Command5.Width
    Command4.Left = Command5.Left - 120 - Command4.Width
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        KeyCode = 0
    End If
End Sub

Private Sub Timer1_Timer()
    Dim p As POINTAPI
    GetCursorPos p
    Static p_ As POINTAPI
    If p.x = p_.x And p.y = p_.y Then Exit Sub
    p_ = p
    Text2.Text = p.x & " x " & p.y
End Sub

