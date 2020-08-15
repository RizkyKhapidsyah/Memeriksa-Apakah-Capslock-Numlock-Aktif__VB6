VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Apakah Capslock/Numlock Aktif"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Num Lock "
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Caps lock"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Contoh ini akan memeriksa apakah Caps Lock sedang 'aktif.
Tmp1 = GetKeyState(vbKeyCapital)
  If Tmp1 = 1 Then
     MsgBox "CapsLock sedang aktif", vbInformation, _
            "Aktif"
  Else
     MsgBox "CapsLock tidak aktif!", vbCritical, _
            "Tidak Aktif"
  End If
End Sub

Private Sub Command2_Click()
'Contoh ini akan memeriksa apakah Num Lock sedang 'aktif.
Tmp2 = GetKeyState(vbKeyNumlock)
  If Tmp2 = 1 Then
     MsgBox "NumLock sedang aktif", vbInformation, _
            "Aktif"
  Else
     MsgBox "NumLock tidak aktif!", vbCritical, _
            "Tidak Aktif"
  End If
End Sub


