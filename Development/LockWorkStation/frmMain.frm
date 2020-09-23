VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock WorkStation"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2610
      ScaleHeight     =   570
      ScaleWidth      =   1290
      TabIndex        =   5
      Top             =   765
      Width           =   1320
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cancel"
         Height          =   465
         Left            =   45
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.PictureBox cmdLockIt2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1260
      ScaleHeight     =   570
      ScaleWidth      =   1290
      TabIndex        =   3
      Top             =   765
      Width           =   1320
      Begin VB.CommandButton cmdLockIt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lock It"
         Height          =   465
         Left            =   45
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Lock the Windows NT 4.0 Workstation"
      Top             =   135
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   270
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   45
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   645
      Left            =   -225
      TabIndex        =   1
      Top             =   -45
      Width           =   5190
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'**********************************************************************************
'
'   This program will lock the workstation of a Windows NT 4
'   machine.  However, for this to work there are 3 things that
'   must be done.  I've already coded this to happen automatically
'   but just in case you want to do it yourself.
'
'       1. You will need to add a key to the registry.  Open the
'       registry and go to:
'       HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\
'
'       Add a string key "ScreenSaverGracePeriod" and set it's
'       default value to 0.  The value you enter here is how
'       long it takes for the OS to lock after a screensaver has
'       been started.  You can enter any number that you want but
'       I use the number 0 so that it locks the screen immediately.
'
'       2. Make sure you have a screen saver selected in the control
'       panel and the "Password Proctected" feature is turned on.
'
'       3. This step is not always neccessary.  If you have done the
'       steps above and it did not work, then you may need to restart
'       your computer so the registry changes are recongnized.
'
'
'       How it works:

'       Under WINNT 4 there is not an API call to lock the workstation.
'       The program Winlogon actually locks it with internal code.  There
'       is a defalut time that WINNT 4 waits to actually lock the screen
'       after a screen saver has been started.  But you can override this
'       by entering the registry key above and setting it's wait time to
'       what ever you want.  In the code below, it's just starting the
'       currently selected screen saver in which calls winlogon in which
'       looks at the registry setting for wait time then locks the screen.
'       If there is no key "ScreenSaverGracePeriod" in the registry then
'       the default wait time is used.
'
'
'**********************************************************************************
'**********************************************************************************


'*******************
'API's used
'*******************
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'*******************
'Constants used
'*******************
Const WM_SYSCOMMAND As Long = &H112&
Const SC_SCREENSAVE As Long = &HF140&

Public mbActiveLock As Boolean



Private Sub cmdCancel_Click()
    '*******************
    'Quit application
    '*******************
    Unload Me
    
End Sub

Private Sub cmdLockIt_Click()
    Dim hWnd As Long
    Dim nRet As Long
    
    'Here is where all the work is done. 2 lines of code.
    
    '*******************
    'Get Desktop handle
    '*******************
    hWnd = GetDesktopWindow()
    
    '*******************
    'Start screensaver
    '*******************
    nRet = SendMessage(hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
    
End Sub


Private Sub Form_Load()
    Dim sRet As String
    Dim sTime As String
    
    'Make sure the 'ScreenSaverGracePeriod' key exists in registry.
    sRet = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "ScreenSaverGracePeriod")
    
    'if it doesn't exists then create it.
    If sRet = "Not Found" Then
        WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "ScreenSaverGracePeriod", ValString, "0"
    End If

    Call CheckScreenSaver
    
End Sub


Public Function FileExists(strPath As String) As Integer
    FileExists = Not (Dir(strPath) = "")
    
End Function


Private Sub CheckScreenSaver()
    '****************************************
    'Here we check to see if a screen saver
    'is set. If not then set it to the blank
    'screen saver and turn on pw protected.
    '****************************************
    Dim sRet As String
    Dim nFileExists As Integer
    
    
    On Error GoTo ErrorHandler
    
    'Make sure there is a screen saver turned in by checking the registry.
    sRet = ReadRegistry(HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE")
    
    'If "Not Found" then there is not a screen saver set, so set one.
    If sRet = "Not Found" Then
    
        'check if screen saver file exists. Is installed by default by WINNT 4. It should be here.
        '-1 means file exists, 0 means does not exists.
        nFileExists = FileExists("c:\winnt\system32\scrnsave.scr")
        
        If nFileExists = -1 Then
            'set screen saver to the registry
            WriteRegistry HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", ValString, "C:\WINNT\System32\scrnsave.scr"
        End If
    End If

    'turn on screen saver password protection if not on and save orig state in mbActiveLock
    mbActiveLock = ReadRegistry(HKEY_CURRENT_USER, "Control Panel\Desktop\", "ScreenSaverIsSecure")
    If mbActiveLock = False Then
        WriteRegistry HKEY_CURRENT_USER, "Control Panel\Desktop\", "ScreenSaverIsSecure", ValString, "1"
    End If
    
Exit Sub
ErrorHandler:
    MsgBox "Error checking for screen saver.", vbCritical, "Application Error"
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'change screen saver protection back to original state
    If mbActiveLock = False Then
        WriteRegistry HKEY_CURRENT_USER, "Control Panel\Desktop\", "ScreenSaverIsSecure", ValString, "0"
    End If
    WriteRegistry HKEY_CURRENT_USER, "Control Panel\Desktop\", "ScreenSaverIsSecure", ValString, "1"
End Sub
