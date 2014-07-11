Attribute VB_Name = "Module1"
Option Explicit
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const MaxT = 10 '最大出現敵数
Public Const TST = 2 '弾の速さ
Public Const NtoR = 0.0175 'ラジアン値変換定数(3.14159/180)
'状態定数
Public Const SS_NO = 0 'なし
Public Const SS_BE = 1 'ある
Public Const SS_BM = 2 '爆発
Public Const SS_SG = 3 '無敵
'ステージ状態定数
Public Const ST_START = 0
Public Const ST_NORMAL = 1
Public Const ST_BOSS = 2
Public Const ST_END = 3
'ゴースト構造体
Type Gohst
  X As Integer
  Y As Integer
  Vx As Integer
  Vy As Integer
  flg As Byte
  N As Byte
  Status As Byte
  Shot As Byte
  Life As Byte
  Hit As Boolean
End Type
'各変数配列
Public M As Gohst '自機
Public T(MaxT + 1) As Gohst '敵機（配列11はﾎﾞｽ用）
Public TD(99) As Gohst '敵弾
Public MD(5) As Gohst '自弾
Public Backs(25) As Byte '背景配列
Public HN(4) As String 'ハイスコア所持者ネーム
Public HS(4) As Long 'ハイスコア
Public Command_Data(9) As Integer 'キー入力履歴
'Public Sn(360) As Single, Cs(360) As Single 'サインコサイン
Public Bn(18) As Integer 'ボーナス
'
Public NowD As Byte '現在敵弾出現数
Public MaxD As Byte '最大敵弾出現数
Public Hidden_flg As Byte '隠しコマンド成立フラグ
Public Extend As Long '１Up
Public Hit_Num As Long 'ヒット数
Public Hit_Pow As Byte 'ヒットメーター
Public Z As Integer 'ステージ進行量
Public St As Byte 'ステージ数
Public Sc As Long 'スコア
Public Nsc As Byte 'ストックスコア
Public Start_Path As String '起動パス
Public Src_X As Byte '画面比率
Public Stage As Byte 'ステージ状態変数
Public Sp As Integer 'スピード(エラーにしないならByte型の方にしたい)
Public Next_T As Byte '次の敵の出現カウント
Public Pause As Byte 'ポーズフラグ
Public Msp As Byte '自機のスピード
Public BomNum As Byte 'ボム数
Public Game_flg As Boolean 'ゲーム中か？
Public End_flg As Boolean
Function Adp_Comp(ByVal s As String) As Boolean
Dim i As Byte, sComp As String * 1
For i = 0 To Len(s) - 1
  sComp = Mid(s, i + 1, 1)
  If Not (StrComp(sComp, Chr(47), 0) > 0 And StrComp(sComp, Chr(91), 0) < 0) Then
    Adp_Comp = False
    Exit Function
  End If
Next i
Adp_Comp = True
End Function
Function Check_Command(ByVal s As String) As Boolean
'入力のチェック
Dim i As Byte, temp As Byte
Check_Command = False
temp = Len(s) / 2
For i = 0 To temp - 1
  If Command_Data(9 - i) <> Val(Mid(s, 2 * i + 1, 2)) Then Exit Function
Next i
Check_Command = True
End Function
Function Cs(ByVal r As Integer) As Single
Cs = Cos(r * NtoR)
End Function
Sub Draw_Back() '画面下の描画
Dim i As Byte, ret As Long
If St > 3 Then Exit Sub
For i = 0 To 24
  Backs(i) = Backs(i + 1)
Next i
Backs(25) = Int(9 * Rnd)
For i = 0 To 24
  ret = BitBlt(Form1.Mk_p.hdc, i * 8, 118 + Backs(i), 8, 16, Form1.Mm.hdc, 16 * 7 + St * 8, 16, SRCCOPY)
Next i
End Sub
Sub Draw_Bs() 'ボスの描画
Dim i As Byte, ret As Long
If Stage <> ST_BOSS Then Exit Sub
ret = StretchBlt(Form1.Mk_p.hdc, T(11).X - 16, T(11).Y - 16, 32, 32, Form1.Mm.hdc, 96, 16, 16, 16, SRCERASE)
ret = StretchBlt(Form1.Mk_p.hdc, T(11).X - 16, T(11).Y - 16, 32, 32, Form1.Mm.hdc, 16 * St, 16, 16, 16, SRCINVERT)
If T(11).Hit Then '５ボスはヒット時に光らない（泣）
  ret = StretchBlt(Form1.Mk_p.hdc, T(11).X - 16, T(11).Y - 16, 32, 32, Form1.Mm.hdc, 80, 16, 16, 16, SRCERASE)
  ret = StretchBlt(Form1.Mk_p.hdc, T(11).X - 16, T(11).Y - 16, 32, 32, Form1.Mm.hdc, 16 * St, 16, 16, 16, SRCINVERT)
  T(11).Hit = False
End If
If T(11).Status = SS_BM Then
  For i = 0 To 3
    Form1.Mk_p.Circle (T(11).X + Int(32 * Rnd) * Sgn(Rnd - 0.5), T(11).Y + Int(32 * Rnd) * Sgn(Rnd - 0.5)), 10 + Int(22 * Rnd), vbRed
  Next i
End If
End Sub
Sub Etc() 'その他の計算等
Dim i As Byte, ret As Long
If M.Life > 1 Then
  For i = 2 To M.Life
    ret = BitBlt(Form1.Main_p.hdc, 5 + 16 * (i - 2), 10, 16, 16, Form1.Mm.hdc, 96, 0, SRCERASE)
    ret = BitBlt(Form1.Main_p.hdc, 5 + 16 * (i - 2), 10, 16, 16, Form1.Mm.hdc, 64, 0, SRCINVERT)
  Next i
End If
If BomNum > 0 Then
  For i = 0 To BomNum - 1
    ret = BitBlt(Form1.Main_p.hdc, 5 + 16 * i, 352, 16, 16, Form1.Mm.hdc, 96, 0, SRCERASE)
    ret = BitBlt(Form1.Main_p.hdc, 5 + 16 * i, 352, 16, 16, Form1.Mm.hdc, 64, 0, SRCINVERT)
  Next i
End If
'なぜか
'If Sc > (Extend + 1) * 30000 then ....
'とできない！
If Sc > Extend Then M.Life = M.Life + 1: Extend = Extend + 50000
Printf Format(Sc, "000000") + "0", 472, 10, 20, vbGreen
If Nsc > 10 Then Sc = Sc + 10: Nsc = Nsc - 10
If Nsc > 1 Then Sc = Sc + 1: Nsc = Nsc - 1
Printf Format(Sc, "000000") + "0", 472, 10, 20, vbGreen
Form1.Main_p.Line (10, 340)-Step(Hit_Pow * 4, 10), vbGreen, BF
Printf Format(Hit_Num, "@@") & "HIT", 100, 334, 14, vbGreen
End Sub
Sub Conv_Sc() 'ハイスコアと比較
Dim i As Byte, j As Integer 'なぜかByte型でできない
For i = 0 To 4
  If Sc > HS(i) Then
    If i < 4 Then
      For j = 4 To i + 1 Step -1
        HS(j) = HS(j - 1)
        HN(j) = HN(j - 1)
      Next j
    End If
    HS(i) = Sc
    Do
      HN(i) = UCase(Left(InputBox("ハイスコア" & CStr(i + 1) & "位です。" _
            & CStr(Sc * 10) + "点" & vbCrLf & vbCrLf _
            & "半角英文字３文字以内で名前をいれてね。", _
            "ハイスコア！", "***"), 3))
      If HN(i) = "" Then HN(i) = "nak"
    Loop Until (Adp_Comp(HN(i)))
    Exit Sub
  End If
Next i
End Sub
Sub Draw_Md() '自弾の描画
Dim i As Byte, ret As Long
For i = 0 To 4
  If MD(i).Status = SS_BE Then
    ret = BitBlt(Form1.Mk_p.hdc, MD(i).X - 3, MD(i).Y - 5, 6, 11, Form1.Mm.hdc, 118, 0, SRCERASE)
    ret = BitBlt(Form1.Mk_p.hdc, MD(i).X - 3, MD(i).Y - 5, 6, 11, Form1.Mm.hdc, 112, 0, SRCINVERT)
  ElseIf MD(i).Status = SS_BM Then
    Form1.Mk_p.Circle (MD(i).X, MD(i).Y), 16 * MD(i).flg, vbGreen
  End If
Next i
End Sub
Sub Draw_Me() '自機の描画
Dim ret As Long
If M.flg > 0 And M.flg < 21 And (M.flg Mod 2) Then Exit Sub
If M.flg > 35 Then
  Dim Ax As Byte
  Ax = -4 * (Abs(M.flg - 55) > 5) - 8 * (Abs(M.flg - 55) < 6)
  ret = StretchBlt(Form1.Mk_p.hdc, M.X - 8 - Ax, M.Y - 8 - Ax, 16 + 2 * Ax, 16 + 2 * Ax, Form1.Mm.hdc, 96, 0, 16, 16, SRCERASE)
  ret = StretchBlt(Form1.Mk_p.hdc, M.X - 8 - Ax, M.Y - 8 - Ax, 16 + 2 * Ax, 16 + 2 * Ax, Form1.Mm.hdc, 64, 0, 16, 16, SRCINVERT)
ElseIf M.flg > 24 And M.flg < 34 Then
  Form1.Mk_p.Circle (M.X, M.Y), (34 - M.flg) ^ 2, vbRed
Else
  ret = BitBlt(Form1.Mk_p.hdc, M.X - 8, M.Y - 8, 16, 16, Form1.Mm.hdc, 96, 0, SRCERASE)
  ret = BitBlt(Form1.Mk_p.hdc, M.X - 8, M.Y - 8, 16, 16, Form1.Mm.hdc, 64, 0, SRCINVERT)
End If
End Sub
Sub Draw_T() '敵機の描画
Dim i As Byte, ret As Long
For i = 0 To MaxT - 1
  If T(i).Status = SS_BE Then
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 96, -16 * T(i).Vx, SRCERASE)
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 16 * T(i).N, -16 * T(i).Vx, SRCINVERT)
    If T(i).Hit Then
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 80, -16 * T(i).Vx, SRCERASE)
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 16 * T(i).N, -16 * T(i).Vx, SRCINVERT)
      T(i).Hit = False
    End If
  ElseIf T(i).Status = SS_BM Then
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 96, -16 * T(i).Vx, SRCERASE)
      ret = BitBlt(Form1.Mk_p.hdc, T(i).X - 8, T(i).Y - 8, 16, 16, Form1.Mm.hdc, 144, 0, SRCINVERT)
'    Form1.Mk_p.Circle (T(i).X, T(i).Y), T(i).flg ^ 2 - 2 * T(i).flg + 3, vbWhite
  End If
Next i
End Sub
Sub Draw_TD() '敵弾の描画
Dim i As Byte, ret As Long
For i = 0 To MaxD - 1
  If TD(i).Status = SS_BE Then
    ret = BitBlt(Form1.Mk_p.hdc, TD(i).X - 2, TD(i).Y - 2, 5, 5, Form1.Mm.hdc, 133, 0, SRCERASE)
    ret = BitBlt(Form1.Mk_p.hdc, TD(i).X - 2, TD(i).Y - 2, 5, 5, Form1.Mm.hdc, 128, 0, SRCINVERT)
  ElseIf TD(i).Status = SS_BM Then
    'Form1.Mk_p.Circle (TD(i).X-2, TD(i).Y-2), TD(i).flg ^ 2, vbYellow
    ret = BitBlt(Form1.Mk_p.hdc, TD(i).X - 2, TD(i).Y - 2, 5, 5, Form1.Mm.hdc, 133, 0, SRCERASE)
    ret = BitBlt(Form1.Mk_p.hdc, TD(i).X - 2, TD(i).Y - 2, 5, 5, Form1.Mm.hdc, 133, 6, SRCINVERT)
  End If
Next i
End Sub
Sub Help_Mode()
Dim i As Byte
Draw_St
Form1.Caption = "ghost"
Printf "＜ゴーストの遊びかた＞", 70, 12, 28, vbGreen
Printf "－ゲーム中－", 10, 60, 20, vbYellow
Printf "移    動", 30, 100, 14, vbYellow
Printf "ショット", 30, 120, 14, vbYellow
Printf "ボム", 30, 140, 14, vbYellow
Printf "一時停止", 30, 160, 14, vbYellow
Printf "タイトルに戻る", 30, 180, 14, vbYellow
Printf "－タイトル画面時－", 10, 220, 20, vbYellow
Printf "Ｆ２", 30, 260, 14, vbYellow
Printf "Ｆ５", 30, 280, 14, vbYellow
Printf "Ｆ８", 30, 300, 14, vbYellow
Printf "ＥＳＰ", 30, 320, 14, vbYellow
'
Printf "カーソルキー，テンキーの８，２，４，６", 160, 100, 14, vbWhite
Printf "スペースキー，Ｚキー", 160, 120, 14, vbWhite
Printf "シフトキー", 160, 140, 14, vbWhite
Printf "Ｆ３キー", 160, 160, 14, vbWhite
Printf "ＥＳＰキー", 160, 180, 14, vbWhite
Printf "スタート", 160, 260, 14, vbWhite
Printf "スピード入力", 160, 280, 14, vbWhite
Printf "ハイスコア画面", 160, 300, 14, vbWhite
Printf "終了", 160, 320, 14, vbWhite
Printf "Push Any Key", 450, 340, 12, vbGreen
'←↑→↓
i = Int(100 * Rnd)
If HS(0) > 50000 Then
  If i < 6 Or HS(0) > 250000 Then Printf "隠しコマンド１：↑↑↓↓←→←→Ｚ", 350, 240, 10, vbRed
  If i < 3 Or HS(0) > 280000 Then Printf "隠しコマンド２：↑↑↓↓←→←→Ｘ", 350, 260, 10, vbRed
  If i < 1 Or HS(0) > 300000 Then Printf "隠しコマンド３：↑↑↓↓←→←→Ｃ", 350, 280, 10, vbRed
End If
Command_Data(0) = 0
While Command_Data(0) = 0
  DoEvents
Wend
Show_Title
End Sub
Sub High_Score() 'ハイスコア画面
Dim i As Byte
Draw_St
Form1.Caption = "ghost"
Printf Space(6) + "High  Score", 60, 10, 40, vbGreen
For i = 0 To 4
  Printf Str(i + 1) + ".", 90, 78 + 56 * i, 35, vbYellow
  Printf Format(HS(i), "000000") + "0", 180, 78 + 56 * i, 35, vbWhite
  Printf ":", 380, 74 + 56 * i, 35, vbWhite
  Printf HN(i), 400, 78 + 56 * i, 35, vbWhite
Next i
Command_Data(0) = 0
While Command_Data(0) = 0
  DoEvents
Wend
Show_Title
End Sub
Sub Ini_Bs_Shot() 'ボスのショット
Dim i As Byte, j As Byte, k As Byte, r As Integer
If T(11).Status = SS_BM Then Exit Sub
'
r = Search_R(11)
'
If NowD + 1 > MaxD - 1 Then Exit Sub
If (T(11).flg Mod 12) = 0 Then
  Set_TD i, 11, r
End If
'
Select Case St
  Case 0
    If NowD + 5 > MaxD Then Exit Sub
    If (T(11).flg Mod 60) = 0 Then
    For j = 0 To 4
      Set_TD i, 11, r + 30 * (j - 2)
      TD(i).X = TD(i).X + (12 - 4 * Abs(j - 2)) * Cs(r + 30 * (j - 2))
      TD(i).Y = TD(i).Y + (12 - 4 * Abs(j - 2)) * Sn(r + 30 * (j - 2))
    Next j
    End If
  Case 1
    If NowD + 3 > MaxD Then Exit Sub
    If (T(11).flg Mod 36) = 0 Then
      For j = 0 To 2
        Set_TD i, 11, r + 30 * (j - 1)
        TD(i).X = T(11).X + 12 * Cs(r + 30 * (j - 1))
        TD(i).Y = T(11).Y + 12 * Sn(r + 30 * (j - 1))
      Next j
    End If
  Case 2
    If NowD + 4 > MaxD Then Exit Sub
    If (T(11).flg Mod 15) = 0 Then
      r = Int(360 * Rnd)
      For j = 0 To 3
        Set_TD i, 11, r + 90 * j
      Next j
    End If
  Case 3
    If NowD + 7 > MaxD Then Exit Sub
    If (T(11).flg Mod 24) = 0 Then
      For j = 0 To 6
        Set_TD i, 11, r + 15 * (j - 3)
        TD(i).X = T(11).X + 12 * Cs(r + 15 * (j - 3))
        TD(i).Y = T(11).Y + 12 * Sn(r + 15 * (j - 3))
      Next j
    End If
  Case 4
    If T(11).flg = 12 Then
      Ini_T
      Ini_T
      Ini_T
    End If
    If NowD + 3 > MaxD Then Exit Sub
    If (T(11).flg Mod 24) = 0 Then
      For j = 0 To 2
        Set_TD i, 11, r + 30 * (j - 1)
        TD(i).X = T(11).X + 12 * Cs(r + 30 * (j - 1))
        TD(i).Y = T(11).Y + 12 * Sn(r + 30 * (j - 1))
      Next j
    End If
    If NowD + 7 > MaxD Then Exit Sub
    If (T(11).flg Mod 36) = 0 Then
      For j = 0 To 6
        Set_TD i, 11, r + 15 * (j - 3)
        TD(i).X = T(11).X + 12 * Cs(r + 15 * (j - 3))
        TD(i).Y = T(11).Y + 12 * Sn(r + 15 * (j - 3))
      Next j
    End If
End Select
End Sub
Sub Ini_Data() '初期化
Dim i As Byte
For i = 0 To 25: Backs(i) = Int(9 * Rnd): Next i
For i = 0 To 4: MD(i).Status = 0: Next i
For i = 0 To 11: T(i).Status = 0:  Next i
For i = 0 To 99: TD(i).Status = 0: Next i
Sc = 0: Nsc = 0: Z = 0: St = 0: Hit_Num = 0: Hit_Pow = 0
Extend = 50000: NowD = 0: Hidden_flg = 0: BomNum = 2
Stage = ST_START: End_flg = False: Game_flg = False
End Sub
Sub Ini_Md() '自弾の出現
Dim i As Byte
If Kof(vbKeySpace) Or Kof(vbKeyZ) Then M.Shot = M.Shot + 4 + 3 * (M.Shot > 0) Else M.Shot = 0: Exit Sub
If M.Shot > 24 Then M.Shot = 0: Exit Sub
If (M.Shot Mod 4) = 0 Then
  While MD(i).Status <> SS_NO: i = i + 1: Wend
  If i = 5 Then Exit Sub
  With MD(i)
    .X = M.X + 3
    .Y = M.Y
    .Status = SS_BE
    .flg = 0
  End With
End If
End Sub
Sub Ini_T() '敵の出現
Dim i As Byte
Next_T = 10 + Int(8 * Rnd) - (St \ 2)
While T(i).Status <> SS_NO: i = i + 1: Wend
If i = MaxT Then Exit Sub
With T(i)
  .N = Int(4 * Rnd)
  .flg = 0
  '.Vx はグラフィック用
  .Vy = 0
  .Status = SS_BE
  .Shot = 14
  .Life = 2 + (.N > 1)
  If .N < 3 Then
    .X = 200
    .Y = 16 + Int(100 * Rnd)
  Else
    .X = 200 + Int(96 * Rnd)
    .Y = 64
  End If
End With
End Sub
Sub Ini_TD(ByVal k As Byte) '敵弾の出現
Dim i As Byte, j As Byte, r As Long
If NowD + 1 - 2 * (T(k).N = 1) - (T(k).N = 2) - 3 * (T(k).N = 3) > MaxD - 1 Then Exit Sub
T(k).Shot = 20 - 2 * St * (T(k).N = 0) - St * (T(k).N = 1) + 10 * (T(k).N = 2) + 2 * St * (T(k).N = 3)
r = Search_R(k)
If T(k).X < 200 Then
  Select Case T(k).N
    Case 0
      Set_TD i, k, r
    Case 1
      For j = 0 To 2
        Set_TD i, k, r + 30 * (j - 1)
      Next j
    Case 2
      For j = 0 To 1
        Set_TD i, k, r
        TD(i).X = T(k).X + 8 * Cs(r - 30 + 60 * j)
        TD(i).Y = T(k).Y + 8 * Sn(r - 30 + 60 * j)
      Next j
    Case 3
      r = Int(10 * Rnd)
      For j = 0 To 3
        Set_TD i, k, -45 * (r > 4) + j * 90
      Next j
  End Select
End If
End Sub
Sub Input_Command(ByVal KeyCode As Byte) 'キー入力
Dim i As Integer
If KeyCode < vbKeyReturn Or KeyCode > vbKeyF10 Then Exit Sub
For i = 0 To 8
  Command_Data(9 - i) = Command_Data(8 - i)
Next i
Command_Data(0) = KeyCode
End Sub
Sub Load_Ini() 'ハイスコア等のロード
Dim i As Byte
On Error GoTo Err_File
Open Start_Path + "ghost.Dat" For Input As #1
For i = 0 To 4
  Input #1, HN(i), HS(i)
Next i
Input #1, Sp
Sp = Int(Sp)
Close #1
Exit Sub
Err_File:
If Err <> 53 Then MsgBox "エラーです。ファイルが開けません", vbOKOnly, "エラー": End '強制終了
Open Start_Path + "ghost.Dat" For Output As #2
For i = 0 To 4
  Print #2, "***"; ","; 10000 - i * 1000
Next i
Print #2, 10
Close #2
Resume
End Sub
Sub Me_Dei() '死亡
'Printf "Your Score Is", 20 + 2 * M.X, 3 * M.Y - 24, 24, vbRed
'Printf Format(CStr(Sc), "000000") + "0", 20 + 2 * M.X, 3 * M.Y, 22, vbRed
Printf "Your Score Is", 20, 40, 24, vbRed
Printf Format(CStr(Sc), "000000") + "0", 300, 42, 22, vbRed
'<!--入力待ち
Command_Data(0) = 0
Game_flg = False
End_flg = True
While Command_Data(0) <> vbKeySpace _
  And Command_Data(0) <> vbKeyReturn _
  And Command_Data(0) <> vbKeyEscape _
  And Command_Data(0) <> vbKeyZ _
  And Command_Data(0) <> vbKeyF9
  DoEvents
Wend
End Sub
Sub Me_vs_Bs() '自機とボス
Dim Ax As Long
If Stage <> ST_BOSS Then Exit Sub
Ax = 16 + 4 * (Hidden_flg And 2) * (M.flg = 0) - 4 * (M.flg > 35) * (1 - (M.flg > 50 And M.flg < 56))
If M.flg > 24 And M.flg < 34 Then Ax = (33 - M.flg) ^ 2
If Abs(M.X - T(11).X) < Ax And Abs(M.Y - T(11).Y) < Ax + 2 Then
  Nsc = Nsc + 10
  If T(11).Status = SS_BM Then Exit Sub
  Hit_Pow = 20
  Hit_Num = Hit_Num + 1
  Sc = Sc - 1000 * (M.flg = 0)
  T(11).Hit = True
  T(11).Life = T(11).Life + (T(11).Life > 0)
  If M.flg = 0 Then
    M.flg = 34
    M.Life = M.Life + (M.Life > 0)
    BomNum = BomNum + 1
  End If
  If T(11).Life = 0 Then
    Sc = Sc + 10000 - 20000 * (St = 4)
    T(11).Status = SS_BM
    Exit Sub
  End If
End If
End Sub
Sub Md_vs_Bs() '自弾とボス
Dim i As Byte
If Stage <> ST_BOSS Then Exit Sub
For i = 0 To 4
  If MD(i).Status = SS_BE And T(11).Status = SS_BE Then
    If Abs(MD(i).X - T(11).X) < 12 And Abs(MD(i).Y - T(11).Y) < 18 Then
      T(11).Life = T(11).Life + (T(11).Life > 0)
      Nsc = Nsc + 1
      Hit_Pow = 20
      Hit_Num = Hit_Num + 1
      MD(i).Status = SS_BM
      T(11).Hit = True
      Sc = Sc + Bn(-Hit_Num * (Hit_Num < 19) - 18 * (Hit_Num > 18)) - (Hit_Num > 18) * (Hit_Num - 18) * 10
      If T(11).Life = 0 Then
        Sc = Sc + 10000 - 20000 * (St = 4)
        T(11).Status = SS_BM
        Exit Sub
      End If
    End If
  End If
Next i
End Sub
Sub Md_vs_T() '自弾と敵機
Dim i As Byte, j As Byte
For i = 0 To 4
  If MD(i).Status = SS_BE Then
    For j = 0 To MaxT - 1
      If T(j).Status = SS_BE Then
        If Abs(MD(i).X - T(j).X) < 8 And Abs(MD(i).Y - T(j).Y) < 12 Then
          MD(i).Status = SS_BM
          T(j).Life = T(j).Life + (T(j).Life > 0)
          T(j).Hit = True
          Nsc = Nsc + 1
          If T(j).Life = 0 Then
            T(j).Status = SS_BM
            T(j).flg = 3
            Hit_Num = Hit_Num + 1
            Hit_Pow = 20
            Sc = Sc + (10 - 5 * (T(j).N = 2) - 10 * (T(j).N = 0) - 20 * (T(j).N = 1)) * (1 - 9 * (T(j).Shot = 1))
            Sc = Sc + Bn(-Hit_Num * (Hit_Num < 19) - 18 * (Hit_Num > 18)) - (Hit_Num > 18) * (Hit_Num - 18) * 10
          End If
        End If
      End If
    Next j
  End If
Next i
End Sub
Sub Me_vs_T() '自機と敵機
Dim i As Byte, Ax As Long
Ax = 7 + 2 * (Hidden_flg And 2) * (M.flg = 0) - 4 * (M.flg > 35) * (1 - (M.flg > 50 And M.flg < 56))
If M.flg > 24 And M.flg < 34 Then Ax = (33 - M.flg) ^ 2
For i = 0 To MaxT - 1
  If T(i).Status = SS_BE Then
    If Abs(M.X - T(i).X) < Ax And Abs(M.Y - T(i).Y) < Ax + 1 Then
      T(i).Status = SS_BM
      T(i).flg = 3
      Nsc = Nsc + 1
      Hit_Pow = 20
      Hit_Num = Hit_Num + 1
      Sc = Sc - 10 * (M.flg < 25)
      Nsc = Nsc + 1
      If M.flg = 0 Then
        M.flg = 34
        M.Life = M.Life + (M.Life > 0)
        BomNum = BomNum + 1
      End If
    End If
  End If
Next i
End Sub
Sub Me_vs_Td() '自機と敵弾
Dim i As Byte, Ax As Long
Ax = 4 + (Hidden_flg And 2) * (M.flg = 0) - 4 * (M.flg > 35) * (1 - (M.flg > 50 And M.flg < 56))
If M.flg > 24 And M.flg < 34 Then Ax = (33 - M.flg) ^ 2
For i = 0 To MaxD - 1
  If TD(i).Status = SS_BE Then
    If Abs(M.X - TD(i).X) < Ax And Abs(M.Y - TD(i).Y) < Ax + 1 Then
      TD(i).Status = SS_BM
      TD(i).flg = 4
      Nsc = Nsc + 1
      Sc = Sc - 10 * (M.flg > 25)
      If M.flg = 0 Then
        M.flg = 34
        M.Life = M.Life + (M.Life > 0)
        BomNum = BomNum + 1
      End If
    End If
  End If
Next i
End Sub
Sub Mv_Bs() 'ボスの動き
If Stage <> ST_BOSS Then Exit Sub
T(11).flg = T(11).flg + 1
If T(11).flg > 118 Then T(11).flg = 0
If T(11).flg = 0 Then T(11).Vx = T(11).Vx - 4
If T(11).Status = SS_BE Then
  If St = 0 Or St = 2 Then
    T(11).X = 144 + 32 * Cs(6 * T(11).flg)
    T(11).Y = 64 - (32 - 16 * (St = 2)) * Sn(6 * T(11).flg)
  Else
    T(11).X = 144 + (16 - 16 * (St <> 1) - 4 * (St = 4)) * Cs(6 * T(11).flg)
    T(11).Y = 64 - (16 - 16 * (St <> 3)) * Sn(6 * T(11).flg)
  End If
  T(11).X = T(11).X + T(11).Vx
  If T(11).X < 32 Then T(11).Status = SS_BM
  Ini_Bs_Shot
Else 'if T(11).status=ss_bm then
  T(11).X = T(11).X - 4
  If T(11).X < -32 Then Stage_Clear
End If
End Sub
Sub Mv_Md() '自弾の動き
Dim i As Byte
For i = 0 To 4
  If MD(i).Status = SS_BE Then
    MD(i).X = MD(i).X + 6
    If MD(i).X > 192 + 3 Then MD(i).Status = SS_NO
  ElseIf MD(i).Status = SS_BM Then
    MD(i).flg = MD(i).flg + 1
    If MD(i).flg = 2 Then MD(i).Status = SS_NO
  End If
Next i
End Sub
Sub Mv_Me() '自機の動き
'M.flgで行動が変わる
'0      Normal
'1-10    移動可無敵
'10-24   オート無敵
'25-34  爆発中
'35-75  ボム
'DoEvents '<-なぜかここに書いてしまった
M.flg = M.flg + (M.flg > 0)
If M.flg > 10 Then
  If M.flg < 24 Then M.X = M.X + 2
  If M.flg = 24 Then
    If M.Life = 0 Then End_flg = True 'Me_Dei
    M.X = -16: M.Y = 64
  End If
  If M.flg < 34 Then Exit Sub
  If M.flg = 35 Then M.flg = 0
End If
M.Vx = (Kof(vbKeyLeft) Or Kof(vbKeyNumpad4)) - (Kof(vbKeyRight) Or Kof(vbKeyNumpad6))
M.Vy = (Kof(vbKeyUp) Or Kof(vbKeyNumpad8)) - (Kof(vbKeyDown) Or Kof(vbKeyNumpad2))
M.Vx = M.X + Msp * M.Vx * (1 - 1 * (M.flg > 35))
M.Vy = M.Y + Msp * M.Vy * (1 - 1 * (M.flg > 35))
If M.Vx > 8 And M.Vx < 160 Then M.X = M.Vx
If M.Vy > 16 And M.Vy < 112 Then M.Y = M.Vy
End Sub
Sub Mv_T() '敵機の動き
Dim i As Byte
For i = 0 To MaxT - 1
  If T(i).Status = SS_BE Then
    T(i).Shot = T(i).Shot + (T(i).Shot > 0)
    T(i).flg = T(i).flg + 1
    T(i).Vx = (M.X < T(i).X) '自機より右の時(-)
    'Select Caseは遅くなるのだが便利ではある。
    Select Case T(i).N
    Case 0
      T(i).X = T(i).X - 3
      T(i).Y = T(i).Y + Sgn(M.Y - T(i).Y)
    Case 1
      T(i).X = T(i).X + 2 * (T(i).flg < 10 Or T(i).flg > 40) - 1
    Case 2
      T(i).X = T(i).X - 3
      If T(i).Vy = 0 And Abs(M.X - T(i).X) < 48 Then T(i).Vy = Sgn(M.Y - T(i).Y)
      T(i).Vy = T(i).Vy - Sgn(T(i).Vy) * (Abs(T(i).Vy) < 8)
      T(i).Y = T(i).Y + T(i).Vy
    Case 3
      T(i).X = T(i).X - 2
      T(i).Y = 64 + 32 * Sin(15 * T(i).flg * NtoR)
    End Select
    If T(i).X < -8 Or T(i).Y < -8 Or T(i).Y > 136 Then T(i).Status = SS_NO
    If T(i).Status = SS_BE And T(i).Shot = 0 Then Ini_TD i
  ElseIf T(i).Status = SS_BM Then
    T(i).flg = T(i).flg - 1
    If T(i).flg = 0 Then T(i).Status = SS_NO
  End If
Next i
End Sub
Sub Mv_Td() '敵弾の動き
Dim i As Byte
For i = 0 To MaxD - 1
  If TD(i).Status = SS_BE Then
    TD(i).X = TD(i).X + TD(i).Vx
    TD(i).Y = TD(i).Y + TD(i).Vy
    If TD(i).X < -5 Or TD(i).X > 197 Or TD(i).Y < -5 Or TD(i).Y > 133 Then
      TD(i).Status = SS_NO
      NowD = NowD - 1
    End If
  ElseIf TD(i).Status = SS_BM Then
    TD(i).flg = TD(i).flg - 1
    If TD(i).flg = 0 Then
      TD(i).Status = SS_NO
      NowD = NowD - 1
    End If
  End If
Next i
End Sub
Sub Owari() '終了処理
Save_Ini
End
End Sub
Sub Printf(ByVal msg As String, ByVal cx As Integer, ByVal cy As Integer, ByVal fs As Byte, ByVal col As Long)
'Ｃみたい
'ただのメイン画面への出力
With Form1.Main_p
  .FontSize = fs * Src_X
  .ForeColor = col
  .CurrentX = cx
  .CurrentY = cy
End With
Form1.Main_p.Print msg
End Sub 'printf format(sc,"@@@@@@@0"),0,0,20,vbgreen
Sub Save_Ini() 'セーブします
Dim i As Byte
On Error GoTo Err_File
  Open Start_Path & "ghost.Dat" For Output As #1
  For i = 0 To 4
    Print #1, HN(i); ","; HS(i)
  Next i
  Print #1, Sp
  Close #1
  Exit Sub
Err_File:
  Close #1
  End
End Sub
Function Search_R(ByVal k As Byte) As Integer '角度を返す
'前作の使いまわし
Dim Ax As Long, Ay As Long, i As Long
Ax = M.X - T(k).X
Ay = M.Y - T(k).Y
If Ax <> 0 Then
   i = Atn(Ay / Ax) * 57.3 ' =180 / 3.14159
   Search_R = i - 180 * (Ax < 0)
ElseIf Ay >= 0 Then
  Search_R = 90
Else
  Search_R = 270
End If
End Function
Sub main() 'サブ的なメイン？（笑）
Dim X As Long, Y As Long, i As Byte
If App.PrevInstance Then
  MsgBox "複数起動はできません！" & vbCrLf, vbOKOnly, "ghost"
  End
End If
Start_Path = App.Path
If Right(Start_Path, 1) <> "\" Then Start_Path = Start_Path & "\"
Load_Ini
X = Screen.TwipsPerPixelX: Src_X = Int(X / 12)
Y = Screen.TwipsPerPixelY
Load Form1
Form1.Width = Form1.Width - (Form1.ScaleWidth * X) + 576 * X
Form1.Height = Form1.Height - (Form1.ScaleHeight * Y) + 362 * Y
Form1.Left = (Screen.Width - Form1.Width) / 2
Form1.Top = (Screen.Height - Form1.Height) / 2
If Dir(Start_Path & "ghost.Bmp") = "" Then
  Mk_Bmp
  SavePicture Form1.Mk_p.Image, Start_Path & "ghost.bmp"
End If
'For ret = 0 To 90
'  Sn(ret) = Sin(ret * NtoR): Cs(ret) = Cos(ret * NtoR)
'  Sn(180 - ret) = Sn(ret): Cs(180 - ret) = -Cs(ret)
'  Sn(180 + ret) = -Sn(ret): Cs(180 + ret) = -Cs(ret)
'  Sn(360 - ret) = -Sn(ret): Cs(360 - ret) = Cs(ret)
'Next ret
For i = 0 To 9
  Bn(i) = (i + 1) * 10
  Bn(i + 9) = (i + 1) * 100
Next i
Randomize
Ini_Data
Set_Form
Title
End Sub
Public Function Kof(ByVal KeyCode As Long) As Integer 'Key_On_Off（アーケードゲームではない^_^:）
Kof = False
If Sgn(GetAsyncKeyState(KeyCode)) Then Kof = True
End Function
Sub Set_Form() 'フォームにあるものの処理
With Form1
  'form
  'Form1.BorderStyle を３固定（実線）にして下さい。
  .AutoRedraw = True
  .BackColor = vbBlack
  .Caption = "ghost"
  .BackColor = vbBlack
  .Main_p.Font = "ＭＳ Ｐゴシック"
  'main_p
  'Form1.main_p.BorderStyle を0-なしにして下さい。
  Set_Pic .Main_p, 0, -8, 576, 378, True, vbBlack
  .Main_p.IMEMode = vbIMEDisable
  'mk_p
  'Form1.mk_p.BorderStyle を0-なしにして下さい。
   Set_Pic .Mk_p, 0, 0, 192, 128 + 8, False, vbBlack
  .Mk_p.FontSize = 6 * Src_X
  .Mk_p.ForeColor = vbGreen
  'mm
  'Form1.mm_p.BorderStyle を0-なしにして下さい。
  Set_Pic .Mm, 0, 0, 16 * 9 + 16, 32, False, vbWhite
  .Mm.Picture = LoadPicture(Start_Path & "ghost.bmp")
  'form
  .Show
End With
End Sub
Sub Set_Pic(ByVal co As Control, ByVal cx As Integer, ByVal cy As Integer, ByVal cw As Integer, ByVal ch As Integer, ByVal cv As Boolean, ByVal cc As Long)
With co
  co.ScaleMode = vbPixels
  co.Move cx, cy, cw, ch
  co.Visible = cv
  co.BackColor = cc
End With
End Sub
Sub Set_TD(ByRef i As Byte, ByVal j As Byte, ByVal p As Long)
'敵弾をセット
i = 0: While TD(i).Status <> SS_NO: i = i + 1: Wend
With TD(i)
  .Status = SS_BE
  .X = T(j).X
  .Y = T(j).Y
  .Vx = Int(TST * Cs(p))
  .Vy = Int(TST * Sn(p))
End With
NowD = NowD + 1
End Sub
Function Sn(ByVal r As Integer) As Single
Sn = Sin(r * NtoR)
End Function
Function Speed() As Integer
Dim ret As Long
ret = CLng(Sp)
Do
ret = Val(InputBox("半角数字で速さを入れて下さい" & vbCrLf + "１(速い)～１０００（遅い）", "スピード設定", CStr(ret)))
If ret < 1 Or ret > 1000 Then
  MsgBox "適度な数字を入れてください。", vbOKOnly, "エラー"
  ret = 0
End If
Loop While (ret = 0)
Command_Data(0) = 0
Speed = CInt(ret)
End Function
Sub Stage_Clear() 'クリア処理
Z = 0
St = St + 1
If St < 5 Then
  Stage = ST_START
  MaxD = 20 + 15 * St
  T(11).Status = SS_NO
  T(11).Vx = 0
  BomNum = BomNum + 1
Else
  Stage = ST_END
  Sc = Sc + Nsc
  Sc = Sc + 5000 * CLng(BomNum)
  If M.Life > 1 Then Sc = Sc + 1000 * (M.Life - 1)
End If
End Sub
Sub Stage_Status() 'ステージ状態監視
Dim i As Byte
If St < 5 Then
  If Z > 10 And Z < 30 Then Printf "Stage" & Str(St + 1), 180, 140, 45, vbWhite
  If Z = 50 Then Stage = ST_NORMAL
  If Z = 250 Then
    Stage = ST_BOSS
    T(11).X = 144
    T(11).Y = 64
    T(11).Vx = 0
    T(11).flg = 0
    T(11).Status = SS_BE
    T(11).Life = 25 + 10 * St
    For i = 0 To 5
      Form1.Main_p.Move 0, -16
      Form1.Main_p.Move 0, -8
      Form1.Main_p.Move 0, 0
      Form1.Main_p.Move 0, -8
    Next i
  End If
Else
  If Z > 20 And Z < 80 Then Printf "COMPLETE!", 10, 40, 40, vbWhite
  If Z > 80 And Z < 140 Then Printf "Thank You for", 10, 80, 40, vbWhite
  If Z > 140 And Z < 210 Then Printf "Your Playing!", 10, 120, 40, vbWhite
  If Z = 210 Then End_flg = True 'Me_Dei
End If
End Sub
Sub Title() 'タイトル画面（さびしい）
Dim i As Integer
Show_Title
Do
  If Command_Data(0) = vbKeyEnd Or Command_Data(0) = vbKeyF1 Then Help_Mode
  If Command_Data(0) = vbKeyF5 Then Sp = Speed()
  If Command_Data(0) = vbKeyF8 Then High_Score
  If Command_Data(0) = vbKeyF2 Or Command_Data(0) = vbKeyReturn Or Command_Data(0) = vbKeySpace Then Game_Main
  DoEvents
Loop Until (End_flg)
Owari
End Sub
Sub Touch_Form() 'フォームが変形（？）した
If Game_flg Then
  Pause = 1
  Form1.Caption = "ghost -Pause"
  Printf "PAUSE", 240, 200, 20, vbYellow
  DoEvents
End If
End Sub
Sub Play_game()
Dim T As Long, ret As Long
While (Not End_flg)
T = GetTickCount + Sp
If Not Pause Then
  Z = Z + 1
  Next_T = Next_T + (Next_T > 0)
  Hit_Pow = Hit_Pow + (Hit_Pow > 0)
  If Hit_Pow = 0 Then Hit_Num = 0
  If Not (Next_T > 0 Or Stage = ST_START Or Stage = ST_BOSS Or Stage = ST_END) Then Ini_T
  Mv_Me
  Mv_Md
  Ini_Md
  Mv_T
  Mv_Td
  Mv_Bs '
  Md_vs_T
  Md_vs_Bs '
  Me_vs_T
  Me_vs_Bs '
  Me_vs_Td
  Draw_Back
  Draw_T
  Draw_Bs '
  Draw_Md
  Draw_Me
  Draw_TD
  ret = StretchBlt(Form1.Main_p.hdc, 0, 0, 16 * 12 * 3, 16 * 8 * 3, Form1.Mk_p.hdc, 0, 0, 16 * 12, 16 * 8, SRCCOPY)
  '
  Etc
  ret = BitBlt(Form1.Mk_p.hdc, 0, 0, 192, 136, Form1.hdc, 0, 0, SRCCOPY)
  Stage_Status
End If
'------------------------------------------
While GetTickCount < T: DoEvents: Wend
Wend
End Sub
Sub Show_Title()
Form1.Caption = "ghost"
Draw_St
Printf "ghost", 170, 100, 50, vbGreen
Printf "B Version", 340, 160, 10, vbGreen
Printf "F2   START", 230, 230, 14, vbYellow
Printf "F5   SPEED", 230, 260, 14, vbYellow
Printf "F8   HISCORE", 230, 290, 14, vbYellow
Printf "Presents By nakatest1234 In 1998", 214, 330, 10, vbGreen
Command_Data(0) = 0
End_flg = False
End Sub
Public Sub Draw_St()
Dim i As Byte
Form1.Main_p.Cls
For i = 0 To 125
  Form1.Main_p.PSet (Int(576 * Rnd), Int(378 * Rnd)), QBColor(1 + Int(Rnd * 14))
Next i
End Sub
Public Sub Game_Main()
Dim i As Integer
'隠しコマンド入力のチェック
If Check_Command("383840403739373990") Then Hidden_flg = 1
If Check_Command("383840403739373988") Then Hidden_flg = 2
If Check_Command("383840403739373967") Then Hidden_flg = 3
'キーバッファのクリア
Command_Data(0) = 0
While (Kof(vbKeyZ) Or Kof(vbKeySpace))
  i = i - (i < 30000)
  If i > 5000 Then Printf "はよ はなせ", 180, 180, 30, vbYellow
  DoEvents
Wend
'ゲーム開始
Form1.Main_p.Cls
Stage = ST_START: Game_flg = True: End_flg = False: Pause = 0
MaxD = 20
Next_T = 15
With M
  .X = -16
  .Y = 64
  .flg = 24
  .Life = 3
End With
Msp = 3 + (Hidden_flg And 1) '隠しコマンドを入れておくと．．．
Game_Loop 'ゲームメインループ
Me_Dei
Conv_Sc
High_Score
Ini_Data
End Sub
Sub Game_Loop()
Dim T As Long, ret As Long
Do
  DoEvents
  While Pause
    DoEvents
    If End_flg Then Exit Sub
  Wend
  T = GetTickCount + Sp
  Z = Z + 1
  Next_T = Next_T + (Next_T > 0)
  Hit_Pow = Hit_Pow + (Hit_Pow > 0)
  Nsc = Nsc - (BomNum > 9 And St < 5)
  If Hit_Pow = 0 Then Hit_Num = 0
  If Not (Next_T > 0 Or Stage = ST_START Or Stage = ST_BOSS Or Stage = ST_END) Then Ini_T
  Mv_Me
  Mv_Md
  Ini_Md
  Mv_T
  Mv_Td
  Mv_Bs '
  Md_vs_T
  Md_vs_Bs '
  Me_vs_T
  Me_vs_Bs '
  Me_vs_Td
  Draw_Back
  Draw_T
  Draw_Bs '
  Draw_Md
  Draw_Me
  Draw_TD
  ret = StretchBlt(Form1.Main_p.hdc, 0, 0, 16 * 12 * 3, 16 * 8 * 3, Form1.Mk_p.hdc, 0, 0, 16 * 12, 16 * 8, SRCCOPY)
  '
  Etc
  ret = BitBlt(Form1.Mk_p.hdc, 0, 0, 192, 136, Form1.hdc, 0, 0, SRCCOPY)
  Stage_Status
While GetTickCount < T: DoEvents: Wend
Loop Until (End_flg) '終了のフラグが立つまでループ
End Sub
