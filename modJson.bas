Attribute VB_Name = "modJson"
Option Explicit

Private Type ParserContext
    JSONString As String
    I As Long
    Length As Long
    LineNo As Long
    Column As Long
End Type

Const JSONErrCode As Long = 2000

Private Function IsSpace(ByVal Char As String) As Boolean
IsSpace = True
Select Case Char
Case vbCr
Case vbLf
Case vbTab
Case " "
Case Else
    IsSpace = False
End Select
End Function

Private Function GetPositionString(Ctx As ParserContext) As String
GetPositionString = "line " & Ctx.LineNo & " column " & Ctx.Column
End Function

Private Function IsEndOfString(Ctx As ParserContext) As Boolean
IsEndOfString = Ctx.I > Ctx.Length
End Function

Private Function PeekChar(Ctx As ParserContext) As String
PeekChar = Mid$(Ctx.JSONString, Ctx.I, 1)
End Function

Private Sub SkipChar(Ctx As ParserContext, PeekedChar As String)
Ctx.I = Ctx.I + 1
If PeekedChar = vbLf Then
    Ctx.LineNo = Ctx.LineNo + 1
    Ctx.Column = 1
Else
    Ctx.Column = Ctx.Column + 1
End If
End Sub

Private Function GetChar(Ctx As ParserContext) As String
GetChar = PeekChar(Ctx)
SkipChar Ctx, GetChar
End Function

Private Sub SkipSpaces(Ctx As ParserContext)
Dim CurChar As String
Do
    CurChar = PeekChar(Ctx)
    If IsSpace(CurChar) = False Then Exit Do
    SkipChar Ctx, CurChar
Loop
End Sub

Private Function HexCharToVal(ByVal HexCharAsc As Long) As Long
If HexCharAsc >= &H30 And HexCharAsc <= &H39 Then
    HexCharToVal = HexCharAsc - &H30
ElseIf HexCharAsc >= &H41 And HexCharAsc <= &H46 Then
    HexCharToVal = HexCharAsc - &H41 + 10
ElseIf HexCharAsc >= &H61 And HexCharAsc <= &H66 Then
    HexCharToVal = HexCharAsc - &H61 + 10
Else
    Err.Raise JSONErrCode, "JSON Parser"
End If
End Function

Private Function ParseString(Ctx As ParserContext, Optional ByVal IsObjectKey As Boolean = False) As String
Dim CurChar As String
Dim Escape As Boolean
Dim EscapeHex As Boolean
Dim HexNumDigits As Long
Dim HexVal As Long
Dim StartLineNo As Long, StartColumn As Long
StartLineNo = Ctx.LineNo
StartColumn = Ctx.Column - 1
Do
    CurChar = GetChar(Ctx)
    If Len(CurChar) = 0 Then Err.Raise JSONErrCode, "JSON Parser", "Unterminated string starting at " & "line " & StartLineNo & " column " & StartColumn
    If Escape Then
        If EscapeHex Then
            HexVal = HexVal * &H10 + HexCharToVal(AscW(CurChar))
            HexNumDigits = HexNumDigits + 1
            If HexNumDigits = 4 Then
                If IsObjectKey And HexVal < &H20 Then Err.Raise JSONErrCode, "JSON Parser", "Invalid control character at " & GetPositionString(Ctx)
                ParseString = ParseString & ChrW$(HexVal)
                EscapeHex = False
                Escape = False
            End If
        Else
            Escape = False
            Select Case CurChar
            Case """"
                ParseString = ParseString & CurChar
            Case "\"
                ParseString = ParseString & CurChar
            Case "/"
                ParseString = ParseString & CurChar
            Case "b"
                ParseString = ParseString & vbBack
            Case "f"
                ParseString = ParseString & vbFormFeed
            Case "n"
                ParseString = ParseString & vbLf
            Case "r"
                ParseString = ParseString & vbCr
            Case "t"
                ParseString = ParseString & vbTab
            Case "u"
                Escape = True
                EscapeHex = True
                HexNumDigits = 0
                HexVal = 0
                Err.Description = "Invalid \uXXXX escape at " & GetPositionString(Ctx)
            Case Else
                Err.Raise JSONErrCode, "JSON Parser", "Invalid \escape at " & GetPositionString(Ctx)
            End Select
        End If
    Else
        If IsObjectKey And AscW(CurChar) < &H20 Then Err.Raise JSONErrCode, "JSON Parser", "Invalid control character at " & GetPositionString(Ctx)
        If CurChar = "\" Then
            Escape = True
        ElseIf CurChar = """" Then
            Exit Do
        Else
            ParseString = ParseString & CurChar
        End If
    End If
Loop
End Function

Private Function GetNumeric(Ctx As ParserContext) As String
Dim CurChar As String
Do
    CurChar = PeekChar(Ctx)
    If IsNumeric(CurChar) Then
        GetNumeric = GetNumeric & CurChar
        SkipChar Ctx, CurChar
    Else
        Exit Do
    End If
Loop
If Len(GetNumeric) = 0 Then Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)
End Function

Private Function NumericToInteger(Numeric As String) As Variant
On Error GoTo Try1

NumericToInteger = CLng(Numeric)
Exit Function

Try1:
Err.Clear
NumericToInteger = CCur(Numeric)
End Function

Private Function ParseNumber(Ctx As ParserContext, ByVal FirstChar As String) As Variant
Dim IsSigned As Boolean
Dim NumberString As String
Dim CurChar As String
Dim IsSignedExp As Boolean

If FirstChar = "-" Then
    IsSigned = True
    SkipChar Ctx, FirstChar
End If
NumberString = GetNumeric(Ctx)
ParseNumber = NumericToInteger(NumberString)

CurChar = PeekChar(Ctx)
If CurChar = "." Then
    SkipChar Ctx, CurChar
    NumberString = GetNumeric(Ctx)
    ParseNumber = CDbl(ParseNumber) + CDbl(NumberString) / (10 ^ Len(NumberString))
End If

CurChar = PeekChar(Ctx)
If LCase$(CurChar) = "e" Then
    SkipChar Ctx, CurChar
    CurChar = PeekChar(Ctx)
    If CurChar = "-" Then
        SkipChar Ctx, CurChar
        IsSignedExp = True
    End If
    NumberString = GetNumeric(Ctx)
    If IsSignedExp Then NumberString = "-" & NumberString
    ParseNumber = CDbl(ParseNumber) * (10 ^ CDbl(NumberString))
End If

If IsSigned Then ParseNumber = -ParseNumber
End Function

Function IsEmptyArray(TestArray As Variant) As Boolean
IsEmptyArray = True
On Local Error Resume Next
Dim I As Long
I = -1
I = UBound(TestArray)
If I >= 0 Then IsEmptyArray = False
End Function

Private Function ParseList(Ctx As ParserContext) As Variant
Dim CurChar As String
Dim RetList() As Variant
Dim ItemCount As Long

SkipSpaces Ctx
CurChar = PeekChar(Ctx)
If CurChar = "]" Then
    SkipChar Ctx, CurChar
    ParseList = RetList
    Exit Function
End If

ReDim RetList(8)
Do
    ParseSubString Ctx, RetList(ItemCount)
    ItemCount = ItemCount + 1
    If ItemCount >= UBound(RetList) + 1 Then ReDim Preserve RetList(ItemCount * 3 / 2 + 1)
    
    SkipSpaces Ctx
    CurChar = PeekChar(Ctx)
    If CurChar = "]" Then
        SkipChar Ctx, CurChar
        If ItemCount Then
            ReDim Preserve RetList(ItemCount - 1)
        Else
            Erase RetList
        End If
        ParseList = RetList
        Exit Function
    ElseIf CurChar = "," Then
        SkipChar Ctx, CurChar
    Else
        Err.Raise JSONErrCode, "JSON Parser", "Unexpected `" & CurChar & "` at " & GetPositionString(Ctx)
    End If
Loop

End Function

Private Function ParseObject(Ctx As ParserContext) As Variant
Dim JObject As Object
Dim SubItem As Variant
Dim CurChar As String

Set JObject = CreateObject("Scripting.Dictionary")

SkipSpaces Ctx
CurChar = PeekChar(Ctx)
If CurChar = "}" Then
    SkipChar Ctx, CurChar
    Set ParseObject = JObject
    Exit Function
End If

Dim KeyName As String
Do
    CurChar = PeekChar(Ctx)
    If CurChar = """" Then
        SkipChar Ctx, CurChar
        KeyName = ParseString(Ctx, True)
    ElseIf CurChar = "'" Then
        Err.Raise JSONErrCode, "JSON Parser", "Expecting property name enclosed in double quotes at " & GetPositionString(Ctx)
    Else
        Err.Raise JSONErrCode, "JSON Parser", "Key name must be string at " & GetPositionString(Ctx)
    End If
    
    SkipSpaces Ctx
    CurChar = PeekChar(Ctx)
    If CurChar <> ":" Then Err.Raise JSONErrCode, "JSON Parser", "Expecting ':' delimiter at " & GetPositionString(Ctx)
    SkipChar Ctx, CurChar
    SkipSpaces Ctx
    ParseSubString Ctx, SubItem
    JObject.Add KeyName, SubItem
    
    SkipSpaces Ctx
    CurChar = PeekChar(Ctx)
    If CurChar = "}" Then
        SkipChar Ctx, CurChar
        Exit Do
    ElseIf CurChar = "," Then
        SkipChar Ctx, CurChar
        SkipSpaces Ctx
    Else
        Err.Raise JSONErrCode, "JSON Parser", "Expecting ',' delimiter at " & GetPositionString(Ctx)
    End If
Loop

Set ParseObject = JObject
End Function

Private Function ParseBoolean(Ctx As ParserContext, ByVal ExpectedValue As Boolean) As Variant
Dim CurChar As String
Dim Word As String, ExpectedWord As String
Dim I As Long

If ExpectedValue = False Then
    ExpectedWord = "false"
Else
    ExpectedWord = "true"
End If

For I = 1 To Len(ExpectedWord)
    CurChar = GetChar(Ctx)
    If Len(CurChar) Then Word = Word & CurChar Else Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)
Next
If Word = ExpectedWord Then
    ParseBoolean = ExpectedValue
Else
    Err.Raise JSONErrCode, "JSON Parser", "Unknown identifier `" & Word & "` at " & GetPositionString(Ctx)
End If

End Function

Private Sub ParseSubString(Ctx As ParserContext, outParsed As Variant)
SkipSpaces Ctx
If IsEndOfString(Ctx) Then Err.Raise JSONErrCode, "JSON Parser", "Expecting value at " & GetPositionString(Ctx)

Dim CurChar As String
CurChar = PeekChar(Ctx)
If CurChar = """" Then
    SkipChar Ctx, CurChar
    outParsed = ParseString(Ctx)
ElseIf IsNumeric(CurChar) = True Or CurChar = "-" Then
    outParsed = ParseNumber(Ctx, CurChar)
ElseIf CurChar = "[" Then
    SkipChar Ctx, CurChar
    outParsed = ParseList(Ctx)
ElseIf CurChar = "{" Then
    SkipChar Ctx, CurChar
    Set outParsed = ParseObject(Ctx)
ElseIf CurChar = "t" Then
    outParsed = ParseBoolean(Ctx, True)
ElseIf CurChar = "f" Then
    outParsed = ParseBoolean(Ctx, False)
Else
    Err.Raise JSONErrCode, "JSON Parser", "Unexpected `" & CurChar & "` at " & GetPositionString(Ctx)
End If
End Sub

Private Function NewParserContext(JSONString As String) As ParserContext
With NewParserContext
.JSONString = JSONString
.I = 1
.Length = Len(JSONString)
.LineNo = 1
.Column = 1
End With
End Function

Function ParseJSONString(JSONString As String) As Variant
Dim Ctx As ParserContext
Ctx = NewParserContext(JSONString)
ParseSubString Ctx, ParseJSONString
SkipSpaces Ctx
If IsEndOfString(Ctx) = False Then Err.Raise JSONErrCode, "JSON Parser", "Extra data at " & GetPositionString(Ctx)
End Function

Sub ParseJSONString2(JSONString As String, ReturnParsed As Variant)
Dim Ctx As ParserContext
Ctx = NewParserContext(JSONString)
ParseSubString Ctx, ReturnParsed
SkipSpaces Ctx
If IsEndOfString(Ctx) = False Then Err.Raise JSONErrCode, "JSON Parser", "Extra data at " & GetPositionString(Ctx)
End Sub

Private Function Hex4(ByVal Value As Long) As String
Hex4 = Right$("000" & Hex$(Value), 4)
End Function

Private Function EscapeString(ByVal SourceStr As String) As String
Dim I As Long, EI As Long, CurChar As String, CharCode As Long, ToAppend As String

EI = Len(SourceStr)
For I = 1 To EI
    CurChar = Mid$(SourceStr, I, 1)
    CharCode = CLng(AscW(CurChar)) And &HFFFF&
    Select Case CharCode
    Case 0
        ToAppend = "\0"
    Case 1 To 7, &HB, &HE To &H1F, &HD800& To &HDFFF&
        ToAppend = "\u" & Hex4(CharCode)
    Case 8
        ToAppend = "\b"
    Case 9
        ToAppend = "\t"
    Case &HA
        ToAppend = "\n"
    Case &HC
        ToAppend = "\f"
    Case &HD
        ToAppend = "\r"
    Case &H5C
        ToAppend = "\\"
    Case Else
        ToAppend = CurChar
    End Select
    EscapeString = EscapeString & ToAppend
Next
End Function

Function JSONToString(JSONData As Variant, Optional ByVal Indent As Long = 0, Optional ByVal IndentChar = " ", Optional ByVal CurIndentLevel As Long = 0) As String
If IsArray(JSONData) Then
    If IsEmptyArray(JSONData) Then
        JSONToString = "[]"
        Exit Function
    End If
    Dim I As Long, U As Long
    U = UBound(JSONData)
    JSONToString = "["
    CurIndentLevel = CurIndentLevel + 1
    If Indent Then GoSub IndentNextLine
    For I = 0 To U
        JSONToString = JSONToString & JSONToString(JSONData(I), Indent, IndentChar, CurIndentLevel + 1)
        If I <> U Then
            JSONToString = JSONToString & ","
            If Indent Then GoSub IndentNextLine
        End If
    Next
    CurIndentLevel = CurIndentLevel - 1
    If Indent Then GoSub IndentNextLine
    JSONToString = JSONToString & "]"
ElseIf IsObject(JSONData) Then
    Dim JObj As Object, KeyName As Variant, IsNotFirst As Boolean
    Set JObj = JSONData
    If JObj.Count = 0 Then
        JSONToString = "{}"
        Exit Function
    End If
    JSONToString = "{"
    If Indent Then GoSub IndentNextLine
    For Each KeyName In JObj
        If IsNotFirst Then
            JSONToString = JSONToString & ","
            If Indent Then GoSub IndentNextLine
        End If
        JSONToString = JSONToString & """" & KeyName & """: " & JSONToString(JObj(KeyName), Indent, IndentChar, CurIndentLevel + 1)
        IsNotFirst = True
    Next
    CurIndentLevel = CurIndentLevel - 1
    If Indent Then GoSub IndentNextLine
    JSONToString = JSONToString & "}"
Else
    Select Case VarType(JSONData)
    Case vbString
        JSONToString = """" & EscapeString(JSONData) & """"
    Case Else
        If IsNumeric(JSONData) Then
            JSONToString = JSONData
            If Left$(JSONToString, 1) = "." Then
                JSONToString = "0" & JSONToString
            Else
                JSONToString = Replace(JSONToString, "-.", "-0.")
            End If
            JSONToString = Replace(LCase$(JSONToString), "e+", "e")
        Else
            Err.Raise JSONErrCode, "JSON Parser", "Unknown variant type `" & VarType(JSONData) & "`"
        End If
    End Select
End If
Exit Function
AddIndent:
    JSONToString = JSONToString & String(Indent * CurIndentLevel, IndentChar)
    Return
    
AddNewLine:
    JSONToString = JSONToString & vbCrLf
    Return

IndentNextLine:
    GoSub AddNewLine
    GoSub AddIndent
    Return
End Function
