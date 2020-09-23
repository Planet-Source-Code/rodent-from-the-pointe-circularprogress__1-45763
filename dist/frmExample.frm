VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "CircularProgress Example"
   ClientHeight    =   4770
   ClientLeft      =   3120
   ClientTop       =   2070
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   Begin prjExample.CircularProgress cp 
      Height          =   3285
      Left            =   765
      TabIndex        =   0
      Top             =   743
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   5794
      BackColor       =   16777215
      OutlineColor    =   4194304
      Value           =   63
      RemainingColor  =   16744576
      Percent         =   0.63
      ProgressColor   =   8388608
      RemainingFillType=   2
      ProgressFillType=   6
   End
   Begin VB.Timer tmrChange 
      Interval        =   100
      Left            =   15
      Tag             =   "0"
      Top             =   15
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrChange_Timer()
Dim tVal As Long

Randomize Timer

tVal = CInt(tmrChange.Tag) Mod 101
tmrChange.Tag = CStr(CInt(tmrChange.Tag) + 1)
If CInt(tmrChange.Tag) >= 101 Then tmrChange.Tag = CStr(0)

cp.Value = tVal


End Sub
