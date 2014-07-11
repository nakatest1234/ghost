VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'ŒÅ’è(ŽÀü)
   Caption         =   "ghost"
   ClientHeight    =   2415
   ClientLeft      =   1260
   ClientTop       =   1845
   ClientWidth     =   4770
   Height          =   2820
   KeyPreview      =   -1  'True
   Left            =   1200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Ëß¸¾Ù
   ScaleWidth      =   318
   Top             =   1500
   Width           =   4890
   Begin VB.PictureBox Mm 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '‚È‚µ
      Height          =   732
      Left            =   1800
      ScaleHeight     =   49
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.PictureBox Main_p 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  '‚È‚µ
      Height          =   852
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   120
      Width           =   612
   End
   Begin VB.PictureBox Mk_p 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '‚È‚µ
      Height          =   732
      Left            =   840
      ScaleHeight     =   49
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Or KeyCode = vbKeyEscape Then End_flg = True
If Game_flg Then
  If KeyCode = vbKeyF3 Then
    Pause = 1 - Pause
    If Pause Then
      Form1.Caption = "ghost -PAUSE"
      Printf "PAUSE", 240, 200, 20, vbYellow
    Else
      Form1.Caption = "ghost"
    End If
  End If
  If KeyCode = vbKeyShift And M.flg = 0 And BomNum > 0 Then M.flg = 65: BomNum = BomNum - 1
'  If KeyCode = vbKeyEscape And Z > 30 Then Me_Dei
'  If KeyCode = vbKeyF9 Then
'    Game_flg = False
'    Command_Data(0) = 0
'    Ini_Data
'    Title
'  End If
  Exit Sub
End If
Input_Command (KeyCode)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Owari
End Sub
Private Sub Form_Resize()
Touch_Form
End Sub
