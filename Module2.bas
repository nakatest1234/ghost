Attribute VB_Name = "Module2"
Option Explicit
Private Chara_Data
Private Shot_Data
Private TShot_Data
Function Exp_Data(s As String) '圧縮展開（といってもたいした圧縮ではない）
Dim i As Integer, j As Integer, temp As String
i = Len(s) * 0.5
For j = 0 To i - 1
  temp = temp + String(Val("&h" + (Mid(s, 2 * j + 2, 1))) + 1, Mid(s, 2 * j + 1, 1))
Next j
Exp_Data = temp
End Function

Sub Mk_Bmp() '絵を描くよ（ダイレクトに）
Dim h As Byte, i As Byte, j As Integer
Set_Data
Set_Pic Form1.Mk_p, 0, 0, 16 * 9, 32, False, vbWhite
For i = 0 To 6
  Data_Draw Chara_Data, 16 * i, 0, 16, 16, QBColor(Val("&H" + Mid("AEC9F0", i + 1, 1)))
Next i
For h = 0 To 1
  For i = h To 14 + h Step 2
    For j = 0 To 15
      If Form1.Mk_p.Point(80 + j, i) = 0 And ((j + h) Mod 2) Then Form1.Mk_p.PSet (80 + j, i), vbWhite
Next j, i, h
Data_Draw Shot_Data, 16 * 7, 0, 6, 11, vbGreen
Data_Draw Shot_Data, 16 * 7 + 6, 0, 6, 11, vbBlack
Data_Draw TShot_Data, 16 * 8, 0, 5, 5, vbCyan  'vbYellowは見にくかった(ボスと同色)
Data_Draw TShot_Data, 16 * 8 + 5, 0, 5, 5, vbBlack
For i = 0 To 4 'eye
  Form1.Mk_p.Line (16 * i + 10, 6)-Step(0, 2), vbBlack
  Form1.Mk_p.Line (16 * i + 12, 6)-Step(0, 2), vbBlack
Next i
'hand
Form1.Mk_p.PSet (16 * 4 + 7, 10), vbBlack: Form1.Mk_p.PSet (16 * 4 + 8, 12), vbBlack: Form1.Mk_p.PSet (16 * 4 + 7, 13), vbBlack
Form1.Mk_p.Line (16 * 4 + 6, 11)-Step(0, 3), vbBlack
'bandana
Form1.Mk_p.Line (5, 3)-Step(9, 1), QBColor(12), BF
Form1.Mk_p.PSet (5, 3), vbWhite: Form1.Mk_p.PSet (5 + 9, 3), vbWhite
'gantai
Form1.Mk_p.Line (16 * 1 + 5, 11)-Step(9, -9), vbBlack
Form1.Mk_p.PSet (16 * 1 + 4, 11), vbBlack: Form1.Mk_p.PSet (16 * 1 + 9, 6), vbBlack
'gurasan
Form1.Mk_p.Line (16 * 2 + 5, 6)-Step(10, 0), vbBlack: Form1.Mk_p.Line (16 * 2 + 8, 7)-Step(6, 0), vbBlack
Form1.Mk_p.PSet (16 * 2 + 9, 8), vbBlack: Form1.Mk_p.PSet (16 * 2 + 12, 8), vbBlack
'mouse
Form1.Mk_p.Line (16 * 3 + 8, 9)-Step(5, 3), QBColor(13), BF
Form1.Mk_p.PSet (16 * 3 + 8, 9), QBColor(9): Form1.Mk_p.PSet (16 * 3 + 9, 9), QBColor(9)
Form1.Mk_p.PSet (16 * 3 + 13, 9), QBColor(9): Form1.Mk_p.PSet (16 * 3 + 8, 10), QBColor(9)
'Back
For i = 0 To 3
  Form1.Mk_p.Line (16 * 7 + 8 * i, 16)-Step(7, 15), QBColor(10 + i), BF
Next i
'Return
For i = 0 To 6
  j = StretchBlt(Form1.Mk_p.hdc, 16 * (i + 1) - 1, 16, -16, 16, Form1.Mk_p.hdc, 16 * i, 0, 16, 16, SRCCOPY)
Next i
Form1.Mk_p.Refresh
End Sub


Sub Set_Data() 'データをいれる
Dim i As String, j As String, k As String
i = "0F" + "071303" + "061502" + "051701" _
  + "041900" + "041900" + "041900" + "041900" _
  + "031A00" + "031A00" + "031A00" + "031A00" _
  + "021B00" + "001C01" + "021902" + "061403"
j = "1103" + "001201" + "001300" + "0113" + "0113" _
  + "0113" + "0113" + "0113" + "001300" + "001201" + "1103"
'展開しましょう
Chara_Data = Exp_Data(i)
Shot_Data = Exp_Data(j)
TShot_Data = "00100" + "01110" + "11111" + "01110" + "00100"
End Sub

Sub Data_Draw(ByVal s As String, X As Integer, Y As Integer, w As Integer, h As Integer, col As Long)
'なぜかSub Data_Drwa(s as string,....)とできなかった
'引数が多いデータ描画
Dim i As Integer
For i = 1 To w * h
  If Mid(s, i, 1) <> "0" Then Form1.Mk_p.PSet (X + (i - 1) Mod w, Y + Int((i - 1) / w)), col
Next i
End Sub


