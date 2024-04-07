Attribute VB_Name = "JSONParser"
Option Explicit

Function JsonParse(JSON, Path$, Optional Header As Boolean = True, Optional PAD$ = vbNullString, Optional Delimiter$ = "/")
'======================================================================================================
' Name        JsonParse
' Version     v3r0      2024-03-28      EGTools(egexcelvba@gmail.com)
' Param       JSON      JSON text, Array or Range
'             Path      �˻��� �̸�, �̸� ����� ��� "/"�� ����ؼ� ����
'                       Path�� �� �ܰ迡 ���ϵ�ī�� ��� ���� "*/�̸�", "�̸�1/*/�̸�*"
'                       "**/�̸�"�� ����� ��� �ܰ踦 �����ϰ� ����
'                       �̸���� ��ü�� "/"�� �ִ� ��� "\/"�� �����ؾ� ��
'             Header    �˻������ ����� ��� ���� ���� �ʵ��(Path)�� ǥ���� �� ����
'             PAD       �迭�� ���� �� �� ���� ä�� ��, �⺻���� ���ڿ�("")
'             Delimiter Header�� �����ʵ�� ����� �ٴܰ��� ��� ����� ������, �⺻���� "/"
' Descripton  JSON text�� Name:Value ���� �˻��Ͽ� Path�� �ش��ϴ� ����� ���
'             �˻� ����� ���� ����� ��� Header�� ���� �ʵ���� ǥ��
'             �˻� ����� ���� ����� �Ϻ� �ʵ��� Path�� ������ �׸��� Header�� ���� �ʵ������ ǥ��
'             �˻� ����� 1�� Field�� ��� Value���� ǥ��
'             Header�� ���� �ʵ尡 ���� �ܰ谡 ���� ��� �� �ܰ踦 Delimiter�� �����Ͽ� ǥ��
'======================================================================================================
 
    Dim vD, vRow, vResult, vPath, vName, vItem, iR&, iC&, iCnt&, iCol&, iRow&, oDict As Object, Key$, Val
     
    If Trim(Path) = "" Then JsonParse = CVErr(xlErrValue): Exit Function
    vD = JSONpair(JSON, Chr(7)) '// �����ڸ� (Bell)�� �����Ͽ� �۾��Կ� ����
    If IsError(vD) Then JsonParse = vD: Exit Function
    If Not IsArray(vD) Then JsonParse = CVErr(xlErrValue): Exit Function
    If UBound(vD, 1) = 1 And UBound(vD, 2) = 1 Then JsonParse = CVErr(xlErrNA): Exit Function
    Path = Replace(Replace(Replace(Path, "\/", vbTab), "/", Chr(7)), vbTab, "/")
    vPath = Split(Path, Chr(7))    '// �̸���θ� �迭�� ����
    
    '// �� ���� �˻��Ͽ� ����� ����
    ReDim vRow(1 To UBound(vD, 1), 1 To 4)  '�� 1�� ���ȣ, ��2�� ����ȣ ����
    For iR = 1 To UBound(vD, 1) '// �̸���ο� ��Ī�Ǵ� �ุ ���
        If vD(iR, 1) Like (Path & "*") Then
            vName = vD(iR, 1)
            If InStr(Path, "**") Then                               '// ��� ������ ��ġ �˻�
                vName = Mid(vName, InStr(vName, Replace(Path, "**" & Chr(7), "")))
            ElseIf InStr(Path, "*") Then                            '// ��� �Ϻ� ���ϵ� ī�� �˻�
                vName = Split(vD(iR, 1), Chr(7))
                vItem = Split(Path, Chr(7))
                For iC = LBound(vName) To Application.Min(UBound(vName), UBound(vItem))
                    If Not vName(iC) Like vItem(iC) Then Exit For
                Next iC
                Key = ""
                For iC = iC To Application.Min(UBound(vName), UBound(vItem))
                    Key = Key & IIf(Key = "", "", Chr(7)) & vName(iC)
                Next iC
            ElseIf Left(vName & Chr(7), Len(Path) + 1) = Path & Chr(7) Then  '// ������ġ �˻�
                vName = Mid(vName, Len(Path) + 2)
            Else
                GoTo NEXT_ROW
            End If
            iCnt = iCnt + 1
            vRow(iCnt, 1) = iR
            If IsArray(vName) Then vName = Join(vName, Delimiter)
            vRow(iCnt, 2) = vName
            vRow(iCnt, 3) = vD(iR, 1)
            vRow(iCnt, 4) = vD(iR, 2)
        ElseIf vD(iR, 1) = vbNullChar Then    '// ��ϱ��� ������ �״��
            iCnt = iCnt + 1
            vRow(iCnt, 1) = iR
        End If
NEXT_ROW:
    Next iR
    
    If iCnt = 0 Then JsonParse = CVErr(xlErrNA): Exit Function
    vRow = ResizeArray(vRow, iCnt, 4)
    
    '// �˻��� ���� �̸��� Dictionary�� �̿��Ͽ� ����ȣ�� ����ϰ� ����
    Set oDict = CreateObject("Scripting.Dictionary")
    ReDim vResult(1 To iCnt + IIf(Header, 1, 0), 1 To 100)
    iRow = 1 + IIf(Header, 1, 0)
    For iR = 1 To iCnt
      Val = vRow(iR, 4)
      Key = vRow(iR, 2)
      If Val <> """""" Then       '// �� ��ü�� ""�� ��쿡�� �ֵ���ǥ ���� ���ܽ�Ŵ
          If Left(Val, 1) = """" Then Val = Mid(Val, 2)
          If Right(Val, 1) = """" Then Val = Left(Val, Len(Val) - 1)
      End If
      
      If IsEmpty(vRow(iR, 3)) Or vRow(iR, 3) = vbNullChar Then  '// ��ϱ��� ������ �ٹٲ� ó��
          iRow = iRow + 1
          If iRow > UBound(vResult, 1) Then vResult = ResizeArray(vResult, iRow, UBound(vResult, 2))
      Else
          If Val <> vbNullString And Key = vbNullString Then    '// ���� �̸��� ���� ��
              Key = "'No_Name'"
          End If
          If Key <> vbNullString Then                           '// ���� �̸��� �ִ� ��
              If oDict.exists(Key) Then
                  vItem = oDict(Key)
                  vItem(1) = Application.Max(vItem(1) + 1, iRow)
                  If vItem(1) > iRow Then iRow = vItem(1) 'iRow + 1
                  If iRow > UBound(vResult, 1) Then vResult = ResizeArray(vResult, iRow, UBound(vResult, 2))
              Else
                  iCol = iCol + 1
                  If iCol > UBound(vResult, 2) Then vResult = ResizeArray(vResult, UBound(vResult, 1), iCol)
                  ReDim vItem(0 To 1): vItem(0) = iCol: vItem(1) = iRow   '1 + IIf(Header, 1, 0)
              End If
              oDict(Key) = vItem
              vResult(vItem(1), vItem(0)) = Val
          End If
      End If
    Next iR
    
    '// ���� ũ��� �迭 ����, ����� �ٹٲ޿� ���� ���κ� ���� ����
    If iRow > UBound(vResult, 1) Then iRow = UBound(vResult, 1)
    vItem = Trim(Join(Application.Index(vResult, iRow), ""))
    Do While vItem = ""
    iRow = iRow - 1: If iRow = 0 Then Exit Do
    vItem = Trim(Join(Application.Index(vResult, iRow), ""))
    Loop
    If iRow = 0 Then JsonParse = CVErr(xlErrNA): GoTo EXIT_RUN
    vResult = ResizeArray(vResult, iRow, iCol)
    
    '// ��� ���� ���� 1���� ��� �� ���� �״�� ���
    If iRow = 1 + IIf(Header, 1, 0) And iCol = 1 Then JsonParse = vResult(1 + IIf(Header, 1, 0), 1): GoTo EXIT_RUN
        
    If Header Then
        If iCol = 1 And oDict.keys()(0) = "'No_Name'" Then
        '// ���� 1�� ���̰�, �̸��� 'No_Name'�̸� Path�� ������ �׸��� �������� ���
            vResult(1, 1) = vPath(UBound(vPath))
        Else
        '// ���� �������� ��� �������� �� �̸��� ���� ���
            For Each vItem In oDict.keys
                vResult(1, oDict(vItem)(0)) = Replace(vItem, Chr(7), Delimiter)
            Next vItem
        End If
    End If
        
    If PAD <> vbNullChar Then       '// �� ä��� ������ ����
      For iR = 1 To UBound(vResult, 1)
        For iC = 1 To UBound(vResult, 2)
            If IsEmpty(vResult(iR, iC)) Then vResult(iR, iC) = PAD
        Next iC
      Next iR
    End If
    
    JsonParse = vResult

EXIT_RUN:
    Set oDict = Nothing

End Function

Function JSONtoArray(JSON, Optional PAD$ = vbNullString)
'======================================================================================================
' Name        JSONtoArray
' Version     v3r0      2024-03-28      EGTools(egexcelvba@gmail.com)
' Param       JSON      JSON text, Array or Range
'             PAD       �迭�� ���� �� �� ���� ä�� ��, �⺻���� ���ڿ�("")
' Descripton  JSON text�� 2D �迭�� ��ȯ
'             "�̸� ���"�� �� �ܰ躰�� �����Ͽ� �����ϰ� �� �����ʿ� "��"�� ����
'             "�̸� ���"���� �ֵ���ǥ�� ���� �̸��� ǥ�õ�
'             "��"�� ����/True/False ���� ���ڿ��� �ֵ���ǥ�� ǥ�õ�
'             ���("["�� "]"���)���� �� ����� �����ϱ� ���� �� ���� ���Ե�
'======================================================================================================
    Dim iR&, iC&, iCol&, vD, vItem
    
    vD = JSONpair(JSON, Chr(7))  '// (Bell)�� �����ڷ� ����Կ� ����
    
    If IsError(vD) Then JSONtoArray = vD: Exit Function
    If Not IsArray(vD) Then JSONtoArray = CVErr(xlErrValue): Exit Function
    For iR = 1 To UBound(vD, 1)
        iCol = Application.Max(iCol, Len(vD(iR, 1)) - Len(Replace(vD(iR, 1), Chr(7), "")))
    Next iR
    ReDim vResult(1 To UBound(vD, 1), 1 To iCol + 2)
    For iR = 1 To UBound(vD, 1)
        vItem = Split(vD(iR, 1), Chr(7))
        For iC = LBound(vItem) To UBound(vItem)
            vResult(iR, iC - LBound(vItem) + 1) = vItem(iC)
        Next iC
        vResult(iR, iC + 1) = IIf(IsEmpty(vD(iR, 2)), PAD, vD(iR, 2))
        For iC = iC + 2 To UBound(vResult, 2)
            vResult(iR, iC) = PAD
        Next iC
    Next iR
    JSONtoArray = vResult
End Function

Function JSONpair(JSON, Optional Delimiter$ = "/")
'======================================================================================================
' Name        JSONpair
' Version     v3r0      2024-03-28      EGTools(egexcelvba@gmail.com)
' Param       JSON      JSON text, Array or Range
'             Delimiter "�̸� ���" �� �ܰ踦 ������ ������, �⺻���� "/"
' Descripton  JSON text�� "�̸� ���"�� "��"���� ������ �迭�� ��ȯ
'             ù��° ���� "�̸� ���", �ι�° ���� "��"
'             "��"�� number/True/False ���� ���ڿ��� �ֵ���ǥ�� ǥ�õ� (JSON value �״��)
'             ���('['�� ']'���)���� �� ����� �����ϱ� ���� �� ���� ���Ե�
'======================================================================================================
    Dim Text$, idx&, sBuf$, iLvl&, iR&, iC&, iCnt&, vD, vValue, vH, vLvl, vList, X$, P$, N$, vItem, vResult
    Dim bName As Boolean, bVal As Boolean, inDQ As Boolean, BlankRow As Boolean, ArrayStart As Boolean
    
    '// ���� ���� ������ ���ڼ��� 32767�ڷ� TEXTJOIN�� ����ص� �Ѱ踦 ���� ���ϹǷ�
    '// Range�� �޾Ƽ� ���� VBA String ������ ������ �ֵ��� ��
    If TypeName(JSON) = "Range" Then
        If JSON.Cells.Count = 1 Then
            ReDim vD(0):        vD(0) = JSON.Value2
        Else
            vD = JSON.Value2
        End If
    ElseIf IsArray(JSON) Then
        vD = JSON
    Else
        ReDim vD(0):        vD(0) = JSON
    End If
    
    For Each vItem In vD: Text = Text & vItem: Next
    
    '// �迭�� ����� ũ�� ���ְ� �۾��� �ϴ� ���� �ӵ��� �ſ� �߿���
    iCnt = Len(Text) - Len(Replace(Replace(Replace(Text, "{", ""), "[", ""), ",", ""))
    ReDim vD(1 To iCnt, 1 To 100) '// ó���� ũ�� �ߴٰ� �������� ResizeArray�� ����
    ReDim vValue(1 To iCnt, 1 To 1)  '// ���� ������ �����ϱ� ���� ���
    ReDim vLvl(1 To 100)            '// �̸������ �ܰ踦 �����ϱ� ���� �迭
    ReDim vList(1 To 100)           '// �迭�� ���� ��ġ�� ǥ���ϱ� ���� �迭
    ReDim vH(1 To 1)                '// ���� �̸� ��θ� �����ϴ� �迭
    iR = 1
    idx = NextIdx(Text, idx, inDQ)  '// ������ Skip
    
    If InStr("{[", Mid(Text, idx, 1)) = 0 Then    '// Object�� Array�� �ƴ� ���
        ReDim vD(1 To 1, 1 To 1): vD(1, 1) = Text: JSONpair = vD: Exit Function
    End If
    
    Do While idx <= Len(Text)
        X = Mid(Text, idx, 1)
        Select Case X
        Case "{"                    '// ��ü ����, Level�� ����
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If P <> "[" And P <> ":" Then GoSub ADD_COL
            iLvl = iLvl + 1
            vLvl(iLvl) = iC
        Case "}"                    '// ��ü ����, Level�� ���
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If sBuf <> "" Then GoSub WRITE_BUF
            iLvl = iLvl - 1:
            If iLvl = 0 Then iR = iR + 1: Exit Do                  'Exit Function
            iC = vLvl(iLvl)
            ReDim Preserve vH(1 To iLvl)
            If InStr("]}", P) = 0 Then GoSub ADD_ROW
        Case "["                    '// �迭 ����, Level�� ����, �迭�ȿ� �迭�� ����
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If iLvl = 0 Then iLvl = iLvl + 1: vLvl(iLvl) = 1: iC = 1: ArrayStart = True
            vList(iLvl) = iLvl
            If P <> "{" And P <> ":" Then GoSub ADD_COL
        Case "]"                    '// �迭 ����, ���� ���� �ʿ�
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If InStr("}]", P) = 0 Then GoSub WRITE_BUF
            If InStr("}]", P) = 0 And InStr(",}]", NextChar(Text, idx)) Then GoSub ADD_ROW
            If ArrayStart And Trim(Join(Application.Index(vD, iR), "")) = Trim(Join(vH, "")) Then iR = iR - 1
        Case ":"                    '// �̸� ������, ���� ���� �ʿ�. �ֵ���ǥ ���� ���ۿ� �߰�
            If inDQ Then sBuf = sBuf & X: GoTo NEXT_IDX
            GoSub WRITE_BUF
            GoSub ADD_COL
        Case ","                    '// �� ������, �ձ��ڿ� ���� �Ǵ� �� ��ġ���� ���� ����, �迭�� ��� ����� ������ �� �ִ� �� �߰�
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If P = "]" Or P = "}" Then bName = True: bVal = False
            If P <> "}" And P <> "]" Then
                GoSub WRITE_BUF
                GoSub ADD_ROW
            ElseIf bVal Then
                GoSub WRITE_BUF
                GoSub ADD_ROW
            End If
            If vList(iLvl) = vLvl(iLvl) And NextChar(Text, idx) = "{" Then BlankRow = True: GoSub ADD_ROW

        Case """"                   '// �ֵ���ǥ In/Out�� ����, ������ �ֵ���ǥ�� �״�� ���
            inDQ = Not inDQ
            If Not isName(Text, idx, inDQ) Then sBuf = sBuf & X: GoTo NEXT_IDX
        Case "\"                    '// Escape ���� ó��
            N = NextChar(Text, idx)
            If N = """" Then sBuf = sBuf & """":      idx = idx + 1     '// Quotation Mark
            If N = "\" Then sBuf = sBuf & "\":        idx = idx + 1     '// Reverse Solidus
            If N = "/" Then sBuf = sBuf & "/":        idx = idx + 1     '// Solidus
            If N = "b" Then sBuf = sBuf & vbBack:     idx = idx + 1     '// Backspace
            If N = "f" Then sBuf = sBuf & vbFormFeed: idx = idx + 1     '// Form Feed
            If N = "n" Then sBuf = sBuf & vbLf:       idx = idx + 1     '// Line Feed or New Line
            If N = "r" Then sBuf = sBuf & vbCr:       idx = idx + 1     '// Carriage Return
            If N = "t" Then sBuf = sBuf & vbTab:      idx = idx + 1     '// Horizontal Tab
            If N = "u" Then sBuf = sBuf & ChrW(Application.Hex2Dec(Mid(Text, idx + 2, 4))): idx = idx + 5 '// Unicode
        Case Chr(7)
            sBuf = sBuf & "(Bell)"
        Case Else                   '// ������ �Ϲ� ���ڴ� ���ۿ� �߰�
            sBuf = sBuf & X
        End Select
        
NEXT_IDX:
        P = X
        idx = NextIdx(Text, idx, inDQ)
    Loop
    
    '// ����ũ��� �迭 ����, ����� �ٹٲ޿� ���� ���κ� ���ٵ� ����
    If iR > UBound(vD, 1) Then iR = UBound(vD, 1)
    If Trim(Join(Application.Index(vD, iR), "")) = "" Then iR = iR - 1
    vItem = Application.Index(vD, iR)
    If Trim(Join(vItem, "")) = "" Then iR = iR - 1
    vD = ResizeArray(vD, iR, Application.Max(vLvl) + 1)
    vValue = ResizeArray(vValue, iR, 1)
    
    If Delimiter = vbNullChar Then Delimiter = Chr(7)
    
    ReDim vResult(1 To iR, 1 To 2)
    For iR = 1 To UBound(vD, 1)
        For iC = 1 To UBound(vD, 2) - 1
            vItem = vD(iR, iC): GoSub REMOVE_DQ
            If vItem <> "" Then vResult(iR, 1) = vResult(iR, 1) & IIf(IsEmpty(vResult(iR, 1)), "", Chr(7)) & vItem
        Next iC
        If Delimiter <> Chr(7) Then vResult(iR, 1) = Replace(vResult(iR, 1), Chr(7), Delimiter)
        vResult(iR, 2) = vValue(iR, 1)
    Next iR
    
    JSONpair = vResult
    Exit Function
    
'========= GoSub Labels =========================================================
    
WRITE_BUF:                          '// ������ ���� �̸����� ������ Ȯ���Ͽ� ó��
    bName = isName(Text, idx, inDQ)
    If bName Then
        ReDim Preserve vH(1 To iLvl)
        For iCnt = 1 To iLvl - 1
            vD(iR, iCnt) = vH(iCnt)
        Next iCnt
        vH(iLvl) = sBuf
        vD(iR, vLvl(iLvl)) = sBuf
        bName = False
        bVal = True
        
    Else
        If iLvl <= UBound(vH) Then  '// {}ó�� ���� ��ü�� �ִ� ��� ���� ����
          For iCnt = 1 To iLvl - 1
              vD(iR, iCnt) = vH(iCnt)
          Next iCnt
          vD(iR, vLvl(iLvl)) = vH(iLvl)
          vValue(iR, 1) = sBuf
        End If
    End If
    sBuf = ""
    iC = vLvl(iLvl)
    Return
    
ADD_COL:                            '// ������, �迭���� ũ�� ũ�� ����
    iC = iC + 1
    If iC > UBound(vD, 2) Then GoSub RESIZE
    Return
    
ADD_ROW:                            '// ������, �迭���� ũ�� ũ�� ����, + ���� �̸���� �Է�
    iR = iR + 1
    If iR > UBound(vD, 1) Then GoSub RESIZE
    For iCnt = 1 To UBound(vD, 2)
        If BlankRow Then            '// ����� ��Ұ� ���н� ù��° ĭ���� vbNullChar -> ������ ����
            If iCnt = 1 Then vD(iR - 1, iCnt) = vbNullChar Else vD(iR - 1, iCnt) = Empty
        End If
    Next iCnt
    BlankRow = False
    Return

RESIZE:
    vD = ResizeArray(vD, Application.Max(iR, UBound(vD, 1)), Application.Max(iC, UBound(vD, 2)))
    vValue = ResizeArray(vValue, Application.Max(iR, UBound(vValue, 1)), 1)
    Return
        
REMOVE_DQ:
    If vItem <> """""" Then       '// �� ��ü�� ""�� ��쿡�� �ֵ���ǥ ���� ���ܽ�Ŵ
        If Left(vItem, 1) = """" Then vItem = Mid(vItem, 2)
        If Right(vItem, 1) = """" Then vItem = Left(vItem, Len(vItem) - 1)
    End If
    Return
        
End Function

Private Function NextIdx(ByRef JSON, ByVal idx&, inDQ)
    Do While idx < Len(JSON) And InStr(IIf(inDQ, "", " ") & vbCr & vbLf & vbTab, Mid(JSON, idx + 1, 1)) > 0
        idx = idx + 1
        If idx > Len(JSON) Then NextIdx = idx: Exit Do
    Loop
    NextIdx = idx + 1
End Function

Private Function NextChar(ByRef JSON, ByVal idx&)
    Do While InStr(" " & vbCr & vbLf & vbTab, Mid(JSON, idx + 1, 1)) > 0
        idx = idx + 1
        If idx > Len(JSON) Then NextChar = "": Exit Function
    Loop
    NextChar = Mid(JSON, idx + 1, 1)
End Function

Private Function isName(ByRef JSON, ByVal idx&, inDQ)
    Dim X$, in_DQ As Boolean
    in_DQ = inDQ
    If Mid(JSON, idx, 1) = """" Then idx = idx + 1
    X = Mid(JSON, idx, 1)
    If X = """" Then in_DQ = Not in_DQ
    Do
        If Not in_DQ Then
            If X = ":" Then isName = True: Exit Function
            If InStr("{[,]}", X) > 0 Then isName = False: Exit Function
        End If
        idx = idx + 1
        If idx > Len(JSON) Then isName = False: Exit Function
        X = Mid(JSON, idx, 1)
        If X = """" Then in_DQ = Not in_DQ
    Loop
    isName = X = ":"
End Function

Private Function ResizeArray(srcArray As Variant, Rows As Long, Cols As Long)
    Dim NewArr, r&, c&, lbR&, ubR&, lbC&, ubC&, iD&
    
    iD = getDimension(srcArray)
    If iD <> 2 Then ResizeArray = CVErr(xlErrRef): Exit Function
    If Rows < 1 Or Cols < 1 Then ResizeArray = CVErr(xlErrValue): Exit Function
    
    lbR = LBound(srcArray, 1):    ubR = UBound(srcArray, 1)
    lbC = LBound(srcArray, 2):    ubC = UBound(srcArray, 2)
    
        ReDim NewArr(lbR To lbR + Rows - 1, lbC To lbC + Cols - 1)
        For r = lbR To UBound(NewArr, 1)
            For c = lbC To UBound(NewArr, 2)
                If r <= ubR And c <= ubC Then
                    NewArr(r, c) = srcArray(r, c)
                Else
                    NewArr(r, c) = Empty
                End If
            Next
        Next
        ResizeArray = NewArr

End Function

Private Function getDimension(var As Variant) As Long
    Dim i&, tmp&
    On Error GoTo Err
    Do While True
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function

