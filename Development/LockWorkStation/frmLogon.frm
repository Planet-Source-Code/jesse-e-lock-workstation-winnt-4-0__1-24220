VERSION 5.00
Begin VB.Form frmLogon 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock WorkStation"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1935
      ScaleHeight     =   480
      ScaleWidth      =   3135
      TabIndex        =   10
      Top             =   1485
      Width           =   3165
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1395
         TabIndex        =   11
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   12
         Top             =   90
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   315
      Picture         =   "frmLogon.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   855
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1935
      ScaleHeight     =   480
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   900
      Width           =   3165
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1395
         TabIndex        =   7
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label lblUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "User Name"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   225
         TabIndex        =   8
         Top             =   90
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3780
      ScaleHeight     =   570
      ScaleWidth      =   1290
      TabIndex        =   4
      Top             =   2565
      Width           =   1320
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cancel"
         Height          =   465
         Left            =   45
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.PictureBox cmdLockIt2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2430
      ScaleHeight     =   570
      ScaleWidth      =   1290
      TabIndex        =   2
      Top             =   2565
      Width           =   1320
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ok"
         Height          =   465
         Left            =   45
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Text            =   "Please enter your User name and Password"
      Top             =   135
      Width           =   5145
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   645
      Left            =   -90
      TabIndex        =   0
      Top             =   -45
      Width           =   5505
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

