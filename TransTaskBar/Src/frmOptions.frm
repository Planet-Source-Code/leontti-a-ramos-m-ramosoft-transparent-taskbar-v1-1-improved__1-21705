VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTaskBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   60
      Picture         =   "frmOptions.frx":27A2
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   276
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   330
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtLevel 
      Height          =   315
      Left            =   4920
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "100"
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CheckBox chkAutoload 
      Caption         =   "Autoload"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4920
      TabIndex        =   4
      Top             =   1260
      Width           =   1200
   End
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   60
      Picture         =   "frmOptions.frx":3636
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   276
      TabIndex        =   7
      Top             =   60
      Width           =   4140
   End
   Begin MSComctlLib.Slider Level 
      Height          =   2175
      Left            =   4260
      TabIndex        =   0
      Top             =   -60
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   3836
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "frmOptions.frx":4339
      Orientation     =   1
      LargeChange     =   51
      SmallChange     =   3
      Min             =   55
      Max             =   255
      SelStart        =   55
      TickStyle       =   2
      TickFrequency   =   25
      Value           =   55
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Start(bLevel As Byte)
    txtLevel = bLevel
    Level.Value = bLevel
    ShowTransparency picTaskBar, picTest, bLevel
    chkAutoload.Value = IIf(IsAutoStart, 1, 0)
    Me.Show
End Sub


Private Sub cmdAbout_Click()
    frmAbout.Start
End Sub


Private Sub cmdAccept_Click()
    If chkAutoload Then
        SetAutoStart Level.Value
    Else
        RemoveAutoStart
    End If
    SaveSetting App.EXEName, "Settings", "TransparencyLevel", Level.Value
    MakeTaskbarTransparent Level.Value
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Level_Change()
    txtLevel = Level.Value
    ShowTransparency picTaskBar, picTest, Level.Value
End Sub

Private Sub txtLevel_GotFocus()
    With txtLevel
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtLevel_KeyPress(KeyAscii As Integer)
    If (((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtLevel_Validate(Cancel As Boolean)
    Dim LnValue As Integer
    
    LnValue = Val(txtLevel)
    If (LnValue > 255) Then LnValue = 255
    If (LnValue < 55) Then LnValue = 55
    Level.Value = LnValue
    txtLevel = Level.Value
End Sub


