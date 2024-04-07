Attribute VB_Name = "JSONParser"
Option Explicit

Function JsonParse(JSON, Path$, Optional Header As Boolean = True, Optional PAD$ = vbNullString, Optional Delimiter$ = "/")
'======================================================================================================
' Name        JsonParse
' Version     v3r0      2024-03-28      EGTools(egexcelvba@gmail.com)
' Param       JSON      JSON text, Array or Range
'             Path      검색할 이름, 이름 경로인 경우 "/"를 사용해서 연결
'                       Path의 각 단계에 와일드카드 사용 가능 "*/이름", "이름1/*/이름*"
'                       "**/이름"를 사용한 경우 단계를 무시하고 적용
'                       이름경로 자체에 "/"가 있는 경우 "\/"로 기재해야 함
'             Header    검색결과가 목록인 경우 제목에 하위 필드명(Path)을 표시할 지 여부
'             PAD       배열을 만든 후 빈 셀에 채울 값, 기본값은 빈문자열("")
'             Delimiter Header에 하위필드명 기재시 다단계인 경우 사용할 구분자, 기본값은 "/"
' Descripton  JSON text중 Name:Value 쌍을 검색하여 Path에 해당하는 결과를 출력
'             검색 결과가 정상 목록인 경우 Header에 따라 필드명을 표시
'             검색 결과가 정상 목록의 일부 필드인 Path의 마지막 항목을 Header에 따라 필등명으로 표시
'             검색 결과가 1개 Field인 경우 Value값만 표시
'             Header에 하위 필드가 여러 단계가 있을 경우 각 단계를 Delimiter로 연결하여 표시
'======================================================================================================
 
    Dim vD, vRow, vResult, vPath, vName, vItem, iR&, iC&, iCnt&, iCol&, iRow&, oDict As Object, Key$, Val
     
    If Trim(Path) = "" Then JsonParse = CVErr(xlErrValue): Exit Function
    vD = JSONpair(JSON, Chr(7)) '// 구분자를 (Bell)로 지정하여 작업함에 주의
    If IsError(vD) Then JsonParse = vD: Exit Function
    If Not IsArray(vD) Then JsonParse = CVErr(xlErrValue): Exit Function
    If UBound(vD, 1) = 1 And UBound(vD, 2) = 1 Then JsonParse = CVErr(xlErrNA): Exit Function
    Path = Replace(Replace(Replace(Path, "\/", vbTab), "/", Chr(7)), vbTab, "/")
    vPath = Split(Path, Chr(7))    '// 이름경로를 배열로 변경
    
    '// 각 행을 검색하여 결과를 기재
    ReDim vRow(1 To UBound(vD, 1), 1 To 4)  '열 1에 행번호, 열2에 열번호 기재
    For iR = 1 To UBound(vD, 1) '// 이름경로에 매칭되는 행만 기록
        If vD(iR, 1) Like (Path & "*") Then
            vName = vD(iR, 1)
            If InStr(Path, "**") Then                               '// 경로 무작위 일치 검색
                vName = Mid(vName, InStr(vName, Replace(Path, "**" & Chr(7), "")))
            ElseIf InStr(Path, "*") Then                            '// 경로 일부 와일드 카드 검색
                vName = Split(vD(iR, 1), Chr(7))
                vItem = Split(Path, Chr(7))
                For iC = LBound(vName) To Application.Min(UBound(vName), UBound(vItem))
                    If Not vName(iC) Like vItem(iC) Then Exit For
                Next iC
                Key = ""
                For iC = iC To Application.Min(UBound(vName), UBound(vItem))
                    Key = Key & IIf(Key = "", "", Chr(7)) & vName(iC)
                Next iC
            ElseIf Left(vName & Chr(7), Len(Path) + 1) = Path & Chr(7) Then  '// 완전일치 검색
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
        ElseIf vD(iR, 1) = vbNullChar Then    '// 목록구분 빈줄은 그대로
            iCnt = iCnt + 1
            vRow(iCnt, 1) = iR
        End If
NEXT_ROW:
    Next iR
    
    If iCnt = 0 Then JsonParse = CVErr(xlErrNA): Exit Function
    vRow = ResizeArray(vRow, iCnt, 4)
    
    '// 검색된 하위 이름은 Dictionary를 이용하여 열번호를 기록하고 재사용
    Set oDict = CreateObject("Scripting.Dictionary")
    ReDim vResult(1 To iCnt + IIf(Header, 1, 0), 1 To 100)
    iRow = 1 + IIf(Header, 1, 0)
    For iR = 1 To iCnt
      Val = vRow(iR, 4)
      Key = vRow(iR, 2)
      If Val <> """""" Then       '// 값 자체가 ""인 경우에는 쌍따옴표 제거 예외시킴
          If Left(Val, 1) = """" Then Val = Mid(Val, 2)
          If Right(Val, 1) = """" Then Val = Left(Val, Len(Val) - 1)
      End If
      
      If IsEmpty(vRow(iR, 3)) Or vRow(iR, 3) = vbNullChar Then  '// 목록구분 빈행은 줄바꿈 처리
          iRow = iRow + 1
          If iRow > UBound(vResult, 1) Then vResult = ResizeArray(vResult, iRow, UBound(vResult, 2))
      Else
          If Val <> vbNullString And Key = vbNullString Then    '// 하위 이름이 없는 것
              Key = "'No_Name'"
          End If
          If Key <> vbNullString Then                           '// 하위 이름이 있는 것
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
    
    '// 최종 크기로 배열 조정, 목록형 줄바꿈에 의한 끝부분 빈줄 삭제
    If iRow > UBound(vResult, 1) Then iRow = UBound(vResult, 1)
    vItem = Trim(Join(Application.Index(vResult, iRow), ""))
    Do While vItem = ""
    iRow = iRow - 1: If iRow = 0 Then Exit Do
    vItem = Trim(Join(Application.Index(vResult, iRow), ""))
    Loop
    If iRow = 0 Then JsonParse = CVErr(xlErrNA): GoTo EXIT_RUN
    vResult = ResizeArray(vResult, iRow, iCol)
    
    '// 결과 값이 오직 1개인 경우 그 값을 그대로 출력
    If iRow = 1 + IIf(Header, 1, 0) And iCol = 1 Then JsonParse = vResult(1 + IIf(Header, 1, 0), 1): GoTo EXIT_RUN
        
    If Header Then
        If iCol = 1 And oDict.keys()(0) = "'No_Name'" Then
        '// 열이 1개 뿐이고, 이름이 'No_Name'이면 Path의 마지막 항목을 제목으로 출력
            vResult(1, 1) = vPath(UBound(vPath))
        Else
        '// 열이 여러개인 경우 각열마다 열 이름을 정상 출력
            For Each vItem In oDict.keys
                vResult(1, oDict(vItem)(0)) = Replace(vItem, Chr(7), Delimiter)
            Next vItem
        End If
    End If
        
    If PAD <> vbNullChar Then       '// 빈셀 채우기 지정시 적용
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
'             PAD       배열을 만든 후 빈 셀에 채울 값, 기본값은 빈문자열("")
' Descripton  JSON text를 2D 배열로 변환
'             "이름 경로"를 각 단계별로 구분하여 기재하고 그 오른쪽에 "값"을 기재
'             "이름 경로"에는 쌍따옴표가 없이 이름만 표시됨
'             "값"은 숫자/True/False 외의 문자열은 쌍따옴표로 표시됨
'             목록("["과 "]"사용)에는 각 목록을 구분하기 위해 빈 행이 삽입됨
'======================================================================================================
    Dim iR&, iC&, iCol&, vD, vItem
    
    vD = JSONpair(JSON, Chr(7))  '// (Bell)을 구분자로 사용함에 주의
    
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
'             Delimiter "이름 경로" 각 단계를 연결할 구분자, 기본값은 "/"
' Descripton  JSON text를 "이름 경로"와 "값"으로 구성된 배열로 변환
'             첫번째 열에 "이름 경로", 두번째 열에 "값"
'             "값"은 number/True/False 외의 문자열은 쌍따옴표로 표시됨 (JSON value 그대로)
'             목록('['과 ']'사용)에는 각 목록을 구분하기 위해 빈 행이 삽입됨
'======================================================================================================
    Dim Text$, idx&, sBuf$, iLvl&, iR&, iC&, iCnt&, vD, vValue, vH, vLvl, vList, X$, P$, N$, vItem, vResult
    Dim bName As Boolean, bVal As Boolean, inDQ As Boolean, BlankRow As Boolean, ArrayStart As Boolean
    
    '// 셀에 저장 가능한 글자수가 32767자로 TEXTJOIN을 사용해도 한계를 넘지 못하므로
    '// Range로 받아서 값을 VBA String 변수에 연결해 넣도록 함
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
    
    '// 배열을 충분히 크게 해주고 작업을 하는 것이 속도에 매우 중요함
    iCnt = Len(Text) - Len(Replace(Replace(Replace(Text, "{", ""), "[", ""), ",", ""))
    ReDim vD(1 To iCnt, 1 To 100) '// 처음에 크게 했다가 마지막에 ResizeArray로 조정
    ReDim vValue(1 To iCnt, 1 To 1)  '// 값을 별도로 기재하기 위해 사용
    ReDim vLvl(1 To 100)            '// 이름경로의 단계를 저장하기 위한 배열
    ReDim vList(1 To 100)           '// 배열이 사용된 위치를 표시하기 위한 배열
    ReDim vH(1 To 1)                '// 현재 이름 경로를 저장하는 배열
    iR = 1
    idx = NextIdx(Text, idx, inDQ)  '// 공백을 Skip
    
    If InStr("{[", Mid(Text, idx, 1)) = 0 Then    '// Object나 Array가 아닌 경우
        ReDim vD(1 To 1, 1 To 1): vD(1, 1) = Text: JSONpair = vD: Exit Function
    End If
    
    Do While idx <= Len(Text)
        X = Mid(Text, idx, 1)
        Select Case X
        Case "{"                    '// 개체 시작, Level을 증가
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If P <> "[" And P <> ":" Then GoSub ADD_COL
            iLvl = iLvl + 1
            vLvl(iLvl) = iC
        Case "}"                    '// 개체 종료, Level을 축소
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If sBuf <> "" Then GoSub WRITE_BUF
            iLvl = iLvl - 1:
            If iLvl = 0 Then iR = iR + 1: Exit Do                  'Exit Function
            iC = vLvl(iLvl)
            ReDim Preserve vH(1 To iLvl)
            If InStr("]}", P) = 0 Then GoSub ADD_ROW
        Case "["                    '// 배열 시작, Level은 유지, 배열안에 배열도 가능
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If iLvl = 0 Then iLvl = iLvl + 1: vLvl(iLvl) = 1: iC = 1: ArrayStart = True
            vList(iLvl) = iLvl
            If P <> "{" And P <> ":" Then GoSub ADD_COL
        Case "]"                    '// 배열 종료, 버퍼 쓰기 필요
            If inDQ And bVal Then sBuf = sBuf & X: GoTo NEXT_IDX
            If InStr("}]", P) = 0 Then GoSub WRITE_BUF
            If InStr("}]", P) = 0 And InStr(",}]", NextChar(Text, idx)) Then GoSub ADD_ROW
            If ArrayStart And Trim(Join(Application.Index(vD, iR), "")) = Trim(Join(vH, "")) Then iR = iR - 1
        Case ":"                    '// 이름 구분자, 버퍼 쓰기 필요. 쌍따옴표 안은 버퍼에 추가
            If inDQ Then sBuf = sBuf & X: GoTo NEXT_IDX
            GoSub WRITE_BUF
            GoSub ADD_COL
        Case ","                    '// 값 구분자, 앞글자에 따라 또는 값 위치에서 버퍼 쓰기, 배열의 요소 변경시 구분할 수 있는 행 추가
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

        Case """"                   '// 쌍따옴표 In/Out을 변경, 값에는 쌍따옴표를 그대로 사용
            inDQ = Not inDQ
            If Not isName(Text, idx, inDQ) Then sBuf = sBuf & X: GoTo NEXT_IDX
        Case "\"                    '// Escape 문자 처리
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
        Case Else                   '// 나머지 일반 문자는 버퍼에 추가
            sBuf = sBuf & X
        End Select
        
NEXT_IDX:
        P = X
        idx = NextIdx(Text, idx, inDQ)
    Loop
    
    '// 최종크기로 배열 조정, 목록형 줄바꿈에 의한 끝부분 빈줄도 삭제
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
    
WRITE_BUF:                          '// 쓰려는 값이 이름인지 값인지 확인하여 처리
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
        If iLvl <= UBound(vH) Then  '// {}처럼 공갈 개체가 있는 경우 오류 방지
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
    
ADD_COL:                            '// 열증가, 배열보다 크면 크기 조정
    iC = iC + 1
    If iC > UBound(vD, 2) Then GoSub RESIZE
    Return
    
ADD_ROW:                            '// 행증가, 배열보다 크면 크기 조정, + 좌측 이름경로 입력
    iR = iR + 1
    If iR > UBound(vD, 1) Then GoSub RESIZE
    For iCnt = 1 To UBound(vD, 2)
        If BlankRow Then            '// 목록의 요소간 구분시 첫번째 칸에만 vbNullChar -> 구분행 기준
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
    If vItem <> """""" Then       '// 값 자체가 ""인 경우에는 쌍따옴표 제거 예외시킴
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

