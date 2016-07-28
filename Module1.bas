Attribute VB_Name = "Module1"
Dim p As Long, yerror As Integer, add As Integer, S As String, i As Integer
Dim Serror(3) As String, ToRead As Boolean
' toread=true :sign ;toread=false :num

Public Function JieCheng(a As Double, d As Double) As Double
a = Int(a): d = Int(d)
If a < 0 Or d < 1 Then yerror = 2: Exit Function
If a = 0 Then JieCheng = 1: Exit Function
On Error GoTo k
For i = a To 1 Step -d
JieCheng = JieCheng * i
Next i
GoTo l
k: yerror = 2
l: End Function

Private Sub readnum(a() As Double, b() As Integer, ps() As Integer, Optional x As Double)
If yerror > 0 Then Exit Sub
If p > Len(S) Then Exit Sub
If ToRead = True Then Exit Sub
Dim k As String, kk As String, q As Integer, SS As String: q = 1
k = Mid(S, p, 1): If p + 1 <= Len(S) Then kk = Mid(S, p, 2)
l:                                                          '读括号
If k = "(" Then add = add + 10: p = p + 1 Else GoTo z2
If p > Len(S) Then Exit Sub
k = Mid(S, p, 1): If p + 1 <= Len(S) Then kk = Mid(S, p, 2) Else kk = ""
GoTo l
z2: Do While k = "+" Or k = "-"                              '读负号
If k = "-" Then q = -q
p = p + 1: If p > Len(S) Then a(i) = q: Exit Sub
k = Mid(S, p, 1): If p + 1 <= Len(S) Then kk = Mid(S, p, 2) Else kk = ""
Loop
If kk = "pi" Then a(i) = 3.14159265358979 * q: ToRead = True: p = p + 2: Exit Sub       '读pi,e
If k = "e" Then a(i) = 2.71828182845905 * q: ToRead = True: p = p + 1: Exit Sub
If k = "X" Then If IsMissing(x) Then ToRead = True: Exit Sub Else a(i) = x * q: p = p + 1: ToRead = True: Exit Sub
Dim E As Boolean: E = False                                                     '计算E的个数
If Not (IsNumeric(k) Or k = "." Or k = "E") Then
    If p = 1 Then a(0) = 1
    Exit Sub
End If
Do While (IsNumeric(k) Or k = "." Or k = "E")
If kk = "E+" Or kk = "E-" Then
    If E Then GoTo Value Else E = True: p = p + 2: SS = SS & kk
ElseIf k = "E" Then
    If E Then GoTo Value Else E = True: p = p + 1: SS = SS & k
ElseIf k = "." Then
    If E Then GoTo Value Else p = p + 1: SS = SS & k
ElseIf IsNumeric(k) Then
    p = p + 1: SS = SS & k
Else: GoTo Value
End If
If p > Len(S) Then Exit Do
k = Mid(S, p, 1): If p + 1 <= Len(S) Then kk = Mid(S, p, 2) Else kk = ""
Loop
Value: If Left(SS, 1) = "E" Then SS = "1" & SS
a(i) = Val(SS) * q: ToRead = True
End Sub

Public Function NumP(a As Double, d As Double) As Double
a = Int(a): d = Int(d): NumP = 1: If d > 2 ^ 31 Then yerror = 2: Exit Function
If a < 0 Or a < d Then yerror = 4: Exit Function
Dim i As Long
On Error GoTo k
For i = 0 To d - 1
NumP = NumP * (a - i)
Next i
GoTo l
k: yerror = 2
l: End Function

Public Function NumC(a As Double, ByVal d As Double) As Double
a = Int(a): d = Int(d): NumC = 1: If d > 2 ^ 31 Then yerror = 2: Exit Function
If a < 0 Or a < d Then yerror = 4: Exit Function
If 2 * d > a Then d = a - d
Dim i As Long
On Error GoTo k
For i = 0 To d - 1
NumC = NumC * (a - i) / (i + 1)
Next i
GoTo l
k: yerror = 2
l: End Function

Private Sub ReadSign(a() As Double, b() As Integer, ps() As Integer, Optional x As Double)
Dim k(1 To 7) As String, flag As Integer
start1: flag = 0
If p > Len(S) Then Exit Sub
For j = 1 To 7: If p + j - 1 <= Len(S) Then k(j) = Mid(S, p, j) Else k(j) = ""
Next j
If k(1) = "+" Then
  If ToRead = True Then b(i) = 0 + add: ps(i) = 1: p = p + 1
ElseIf k(1) = "-" Then
  If ToRead = True Then b(i) = 0 + add: ps(i) = 0: p = p + 1
ElseIf k(1) = "*" Then
  If ToRead = True Then b(i) = 1 + add: ps(i) = 1: p = p + 1 Else yerror = 1
ElseIf k(1) = "/" Then
  If ToRead = True Then b(i) = 1 + add: ps(i) = 0: p = p + 1 Else yerror = 1
ElseIf k(1) = "\" Then
  If ToRead = True Then b(i) = 2 + add: ps(i) = 1: p = p + 1 Else yerror = 1
ElseIf k(1) = "%" Then
  If ToRead = True Then b(i) = 2 + add: ps(i) = 0: p = p + 1 Else yerror = 1
ElseIf k(1) = "^" Then
  If ToRead = True Then b(i) = 3 + add: ps(i) = 2: p = p + 1 Else yerror = 1
ElseIf k(1) = "P" Then
  If ToRead = True Then b(i) = 5 + add: ps(i) = 0: p = p + 1 Else yerror = 1
ElseIf k(1) = "C" Then
  If ToRead = True Then b(i) = 5 + add: ps(i) = 1: p = p + 1 Else yerror = 1
ElseIf k(1) = "&" Then
  If ToRead = True Then b(i) = 4 + add: ps(i) = 0: p = p + 1 Else yerror = 1
ElseIf k(1) = "m" Then
  If ToRead = True Then
    If k(3) = "mod" Then b(i) = 2 + add: ps(i) = 0: p = p + 3 Else yerror = 1
  Else
    yerror = 1
  End If
ElseIf k(1) = "l" Then
  If ToRead = True Then
    If k(3) = "lst" Then
      b(i) = 4 + add: ps(i) = 3: p = p + 3
    ElseIf k(2) = "ln" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 24: p = p + 2: i = i + 1
    ElseIf k(2) = "lg" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 25: p = p + 2: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(2) = "ln" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 24: p = p + 2
    If k(2) = "lg" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 25: p = p + 2
  End If
ElseIf k(1) = "i" Then
  If ToRead = True Then
    If k(3) = "int" Then b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6: ps(i + 1) = 22: i = i + 1: p = p + 3 Else yerror = 1
  Else
    yerror = 1
    If k(3) = "int" Then a(i) = 0: b(i) = 6 + add: ps(i) = 22: p = p + 3: yerror = 0
  End If
ElseIf k(1) = "x" Then
  If ToRead = True Then
    If k(3) = "xor" Then b(i) = 4 + add: ps(i) = 5: p = p + 3 Else yerror = 1
  Else
    yerror = 1
  End If
ElseIf k(1) = "o" Then
  If ToRead = True Then
    If k(2) = "or" Then b(i) = 4 + add: ps(i) = 1: p = p + 2 Else yerror = 1
  Else
    yerror = 1
  End If
ElseIf k(1) = "t" Then
  If ToRead = True Then
    If k(4) = "tanh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 9: p = p + 4: i = i + 1
    ElseIf k(3) = "tgh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 9: p = p + 3: i = i + 1
    ElseIf k(3) = "tan" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 3: p = p + 3: i = i + 1
    ElseIf k(2) = "tg" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 3: p = p + 2: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(4) = "tanh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 9: p = p + 4: yerror = 0
    If k(3) = "tgh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 9: p = p + 3: yerror = 0
    If k(3) = "tan" Then a(i) = 0: b(i) = 6 + add: ps(i) = 3: p = p + 3: yerror = 0
    If k(2) = "tg" Then a(i) = 0: b(i) = 6 + add: ps(i) = 3: p = p + 2: yerror = 0
  End If
ElseIf k(1) = "r" Then
  If ToRead = True Then
    If k(3) = "rst" Then
      b(i) = 4 + add: ps(i) = 4: p = p + 3
    ElseIf k(3) = "rtd" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 17: p = p + 3: i = i + 1
    ElseIf k(3) = "rtg" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 20: p = p + 3: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(3) = "rtd" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 17: p = p + 3
    If k(3) = "rtg" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 20: p = p + 3
  End If
ElseIf k(1) = "n" Then
  If ToRead = True Then
    If k(3) = "not" Then b(i) = 1 + add: ps(i) = 1 + add: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 23 + add: p = p + 3: i = i + 1 Else yerror = 1
  Else
    yerror = 1
    If k(3) = "not" Then a(i) = 0: b(i) = 6 + add: ps(i) = 23: p = p + 3
  End If
ElseIf k(1) = "a" Then
  If ToRead = True Then
    If k(7) = "arcsinh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 10: p = p + 7: i = i + 1
    ElseIf k(7) = "arccosh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 11: p = p + 7: i = i + 1
    ElseIf k(7) = "arctanh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 12: p = p + 7: i = i + 1
    ElseIf k(6) = "arctgh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 12: p = p + 6: i = i + 1
    ElseIf k(6) = "arcsin" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 4: p = p + 6: i = i + 1
    ElseIf k(6) = "arccos" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 5: p = p + 6: i = i + 1
    ElseIf k(6) = "arctan" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 6: p = p + 6: i = i + 1
    ElseIf k(5) = "arctg" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 6: p = p + 5: i = i + 1
    ElseIf k(3) = "abs" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 0: p = p + 3: i = i + 1
    ElseIf k(3) = "and" Then
      b(i) = 4 + add: ps(i) = 0: p = p + 3
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(7) = "arcsinh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 10: p = p + 7: yerror = 0
    If k(7) = "arccosh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 11: p = p + 7: yerror = 0
    If k(7) = "arctanh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 12: p = p + 7: yerror = 0
    If k(6) = "arctgh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 12: p = p + 6: yerror = 0
    If k(6) = "arcsin" Then a(i) = 0: b(i) = 6 + add: ps(i) = 4: p = p + 6: yerror = 0
                If k(6) = "arccos" Then a(i) = 0: b(i) = 6 + add: ps(i) = 5: p = p + 6: yerror = 0
                If k(6) = "arctan" Then a(i) = 0: b(i) = 6 + add: ps(i) = 6: p = p + 6: yerror = 0
                If k(5) = "arctg" Then a(i) = 0: b(i) = 6 + add: ps(i) = 6: p = p + 5: yerror = 0
                If k(3) = "abs" Then a(i) = 0: b(i) = 6 + add: ps(i) = 0: p = p + 3: yerror = 0
  End If
ElseIf k(1) = "d" Then
  If ToRead = True Then
    If k(3) = "dtr" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 16: p = p + 3: i = i + 1
    ElseIf k(3) = "dtg" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 18: p = p + 3: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(3) = "dtr" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 16: p = p + 3
    If k(3) = "dtg" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 18: p = p + 3
  End If
ElseIf k(1) = "g" Then
  If ToRead = True Then
    If k(3) = "gtr" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 21: p = p + 3: i = i + 1
    ElseIf k(3) = "gtd" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 19: p = p + 3: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(3) = "gtr" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 21: p = p + 3
    If k(3) = "gtd" Then yerror = 0: a(i) = 0: b(i) = 6 + add: ps(i) = 19: p = p + 3
  End If
ElseIf k(1) = "s" Then
  If ToRead = True Then
    If k(4) = "sinh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 7: p = p + 4: i = i + 1
    ElseIf k(3) = "sin" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 1: p = p + 3: i = i + 1
    ElseIf k(3) = "sec" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 14: p = p + 3: i = i + 1
    ElseIf k(3) = "sgn" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 26: p = p + 3: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(4) = "sinh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 7: p = p + 4: yerror = 0
    If k(3) = "sin" Then a(i) = 0: b(i) = 6 + add: ps(i) = 1: p = p + 3: yerror = 0
    If k(3) = "sec" Then a(i) = 0: b(i) = 6 + add: ps(i) = 14: p = p + 3: yerror = 0
    If k(3) = "sgn" Then a(i) = 0: b(i) = 6 + add: ps(i) = 26: p = p + 3: yerror = 0
  End If
ElseIf k(1) = "c" Then
  If ToRead = True Then
    If k(4) = "cosh" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 8: p = p + 4: i = i + 1
    ElseIf k(3) = "cos" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 2: p = p + 3: i = i + 1
    ElseIf k(3) = "csc" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 13: p = p + 3: i = i + 1
    ElseIf k(3) = "cot" Then
      b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 15: p = p + 3: i = i + 1
    Else
      yerror = 1
    End If
  Else
    yerror = 1
    If k(4) = "cosh" Then a(i) = 0: b(i) = 6 + add: ps(i) = 8: p = p + 4: yerror = 0
                If k(3) = "cos" Then a(i) = 0: b(i) = 6 + add: ps(i) = 2: p = p + 3: yerror = 0
                If k(3) = "csc" Then a(i) = 0: b(i) = 6 + add: ps(i) = 13: p = p + 3: yerror = 0
                If k(3) = "cot" Then a(i) = 0: b(i) = 6 + add: ps(i) = 15: p = p + 3: yerror = 0
  End If
ElseIf k(1) = "!" Then
  If ToRead = True Then
    b(i) = 7 + add: ps(i) = 0: p = p + 1
    If IsNumeric(Mid(S, p, 1)) Or Mid(S, p, 1) = "." Then a(i + 1) = 1: i = i + 1: flag = 1
  Else
    yerror = 1
  End If
ElseIf k(1) = "X" Then
  If IsMissing(x) Then
    yerror = 1
  Else
      If ToRead = True Then p = p + 1: b(i) = 1 + add: ps(i) = 1: a(i + 1) = x: flag = 1: i = i + 1 Else a(i) = x: flag = 1: p = p + 1
  End If
ElseIf k(1) = "e" Then
  If k(3) = "exp" Then
    If ToRead = True Then b(i) = 1 + add: ps(i) = 1: a(i + 1) = 0: b(i + 1) = 6 + add: ps(i + 1) = 40: p = p + 3: i = i + 1 Else a(i) = 0: b(i) = 6 + add: ps(i) = 40: p = p + 3
  Else
    If ToRead = True Then p = p + 1: b(i) = 1 + add: ps(i) = 1: a(i + 1) = 2.71828182845905: flag = 1: i = i + 1 Else a(i) = 2.71828182845905: flag = 1: p = p + 1
  End If
ElseIf k(2) = "pi" Then
  If ToRead = True Then p = p + 2: b(i) = 1 + add: ps(i) = 1: a(i + 1) = 3.14159265358979: flag = 1: i = i + 1 Else a(i) = 3.14159265358979: flag = 1: p = p + 2
ElseIf k(1) = "(" Then
  If ToRead = True Then b(i) = 1 + add: ps(i) = 1: add = add + 10: p = p + 1 Else add = add + 10: i = i - 1
ElseIf k(1) = ")" Then
  If ToRead = True Then add = add - 10: p = p + 1: flag = 1 Else yerror = 1
ElseIf k(1) = "]" Then
  If ToRead = True Then add = 0: p = p + 1: flag = 1: i = i - 1 Else yerror = 1
ElseIf k(1) = "A" Then
Else
  yerror = 1
End If
If yerror > 0 Then Exit Sub
ToRead = False: i = i + 1
If flag > 0 Then ToRead = True: i = i - 1: GoTo start1
End Sub

Public Sub Calc(ByVal SS As String, ByRef Ans As Double, ByRef Serr As String, Optional ByVal x As Double)
S = SS
Serror(0) = "句法错误或非法字符"
Serror(1) = "被零除或溢出(数据过大)"
Serror(2) = "堆栈溢出(运算符过多)"
Serror(3) = "参数不正确"
Dim a(41) As Double, b(40) As Integer, ps(40) As Integer, m As Integer
p = 1: i = 0: ToRead = False: yerror = 0: add = 0
If Len(S) = 0 Then Exit Sub
Do While p <= Len(S)
If add > 30000 Then yerror = 3: Exit Do
Call readnum(a(), b(), ps(), x)
If IsMissing(x) Then
    Call ReadSign(a(), b(), ps(), x)
Else
    Call ReadSign(a(), b(), ps(), x)
End If
If yerror > 0 Then Exit Do
Loop
If ToRead = False Then yerror = 1
If yerror > 0 Then Serr = Serror(yerror - 1): Exit Sub
m = i: If m < 2 Then GoTo z1
Dim qq As Integer, j() As Integer, t As Integer
Do
ReDim j(m)
qq = 0: t = 0: ii = 0
        Do While qq < m
        If b(qq) > b(qq + 1) Then
          Call Calc2(a(qq), b(qq), ps(qq), a(qq + 1)): t = t + 1: j(qq) = t: qq = qq + 1
        ElseIf b(qq) = b(qq + 1) And b(qq) < 6 Then
          Call Calc2(a(qq), b(qq), ps(qq), a(qq + 1)): t = t + 1: j(qq) = t: qq = qq + 1
        End If
        If yerror > 0 Then Exit Do
        j(ii) = t: qq = qq + 1: ii = ii + 1
        Loop
If yerror > 0 Then Exit Do
If t = 0 Then Exit Do
For qq = 0 To m - 1
If j(qq) > 0 Then b(qq) = b(qq + j(qq)): ps(qq) = ps(qq + j(qq)): a(qq + 1) = a(qq + 1 + j(qq))
Next qq
m = m - t
Loop Until m < 2
If m > 1 Then
        For qq = m - 1 To 1 Step -1
        Call Calc2(a(qq), b(qq), ps(qq), a(qq + 1))
        Next qq
End If
z1: Call Calc2(a(0), b(0), ps(0), a(1))
If yerror > 0 Then Serr = Serror(yerror - 1): Exit Sub
Ans = Val(a(0))
End Sub

Private Sub Calc2(a As Double, b As Integer, c As Integer, d As Double)
Const pi = 3.14159265358979
Const E = 2.71828182845975
Dim m As Long: m = 30 * (b Mod 10) + c
On Error GoTo k
GoTo l
k:: If m = 30 Or m = 61 Then yerror = 2 Else yerror = 4: Exit Sub
l: Select Case m
Case 1
        a = a + d
Case 0
        a = a - d
Case 30
        a = a / d
Case 31
        a = a * d
Case 60
        a = a Mod d
Case 61
        a = a \ d
Case 204
        a = Log(d)
Case 205
        a = Log(d) / Log(10)
Case 92
        a = a ^ d
Case 120
        a = a And d
Case 121
        a = a Or d
Case 203
        a = Not d
Case 123 'lst
        a = a * 2 ^ d
Case 124 'rst
        a = Int(a / (2 ^ d))
Case 125
        a = a Xor d
Case 210
        a = JieCheng(a, d)
        If yerror > 0 Then Exit Sub
Case 180
        a = Abs(d)
Case 181
        a = Sin(d)
Case 182
        a = Cos(d)
Case 183
        a = Tan(d)
Case 184
        If Abs(d) = 1 Then
        a = pi / 2 * Sgn(d)
        Else
        a = Atn(d / (1 - d * d) ^ 0.5)
        End If
        If yerror > 0 Then Exit Sub
Case 185
        If d = 0 Then
        a = pi / 2
        Else
        a = Atn(Abs((1 - d * d) ^ 0.5 / d))
        If d < 0 Then a = pi - a
        End If
        If yerror > 0 Then Exit Sub
Case 186
        a = Atn(d)
Case 187
        a = 0.5 * (E ^ d - E ^ (d))
Case 188
        a = 0.5 * (E ^ d + E ^ (d))
Case 189
        a = (E ^ d - E ^ (-d)) / (E ^ d + E ^ (-d))
Case 190
        a = Log(d + (d * d + 1) ^ 0.5)
Case 191
        a = Log(d + (d * d - 1) ^ 0.5)
Case 192
        a = 0.5 * Log((1 + d) / (1 - d))
Case 193
        a = 1 / Sin(d)
Case 194
        a = i / Cos(d)
Case 195
        If Abs(a / pi - Int(a / pi) - 0.5) < 0.00001 Then a = 0 Else a = 1 / Tan(d)
Case 196
        a = d / 180 * pi
Case 197
        a = d / pi * 180
Case 198
        a = d / 9 * 10
Case 199
        a = d * 0.9
Case 200
        a = d / pi * 200
Case 201
        a = d / 200 * pi
Case 202
        a = Int(d)
Case 206
        a = Sgn(d)
Case 220
        a = Exp(d)
Case 150
        a = NumP(a, d)
        If yerror > 0 Then Exit Sub
Case 151
        a = NumC(a, d)
        If yerror > 0 Then Exit Sub
End Select
End Sub

