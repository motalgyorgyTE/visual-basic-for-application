'   Json converter/Parser module.
'
' Author
'   Mnyerczán Sándor
'   <mnyerczan@outlook.hu>
'
' Standard
'   The JavaScript Object Notation (JSON) Data Interchange Format
'   https://tools.ietf.org/html/rfc7158
'
'
'


Private p As Long               ' Counter
Private token As Variant
Private translate As Object
Private sstck As Stack          ' Structural level Stack
private gao As Boolean          ' Grammatical analisis only


' These are the six structural characters:
Private Enum sc
    leftSquareBracket = &H5B    ' [ left square bracket
    leftCurlyBracket = &H7B     ' { left curly bracket
    rightSquareBracket = &H5D   ' ] right square bracket
    rightCurlyBracket = &H7D    ' } right curly bracket
    colon = &H3A                ' : colon
    comma = &H2C                ' , comma
End Enum



' Insignificant whitespace is allowed before or after any of the six
' structural characters.
Private Enum ws
    spac_e = &H20               ' Space
    horizontalTab = &H9         ' Horizontal tab
    lineFeed = &HA              ' Line feed or New line
    carrageReturn = &HD         ' Carriage return
End Enum



'-------------------------------------------------------------------
' Public functions
'-------------------------------------------------------------------



' Json parsing and coding algorithm
'
Public Function JsonEncode( _
        ByVal jsonPattern As String, _
        Optional grammaticalAnalisisOnly As Boolean = False) As String

    If Len(js) = 0  Then Err.Raise 1020, "json.JsonEncode",
                            "Empty string is not parsable!"

    gao = grammaticalAnalisisOnly

    Set sstck = New Stack
    sstck.SetType ("Long")

    p = 0
    Set translate = CreateObject("Scripting.dictionary")
    translate.Add key:=sc.leftCurlyBracket, Item:=sc.rightCurlyBracket
    translate.Add key:=sc.leftSquareBracket, Item:=sc.rightSquareBracket
    translate.Add key:=sc.rightCurlyBracket, Item:=sc.leftCurlyBracket
    translate.Add key:=sc.rightSquareBracket, Item:=sc.leftSquareBracket

    JsonEncode = JsonEncodeEngine(jsonPattern)
End Function


' Decode json encoded string
'
Public Function JsonDecode(ByVal jsonPattern As String) As String
    Dim i As Long, e As Byte
    Dim s, cache, char, js As String

    i = 1
    e = 0
    js = jsonPattern

    While i < Len(js)
        If Mid(js, i, 1) = "\" And Mid(js, i + 1, 1) = "u" Then
            If IsNumeric(Mid(js, i + 2, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 3, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 4, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 5, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 Then
                For e = 0 To 3
                    cache = cache & Mid(js, i + 2 + e, 1)
                Next

                char = ChrW("&H" & cache)
                s = s & char

                i = i + 5
                cache = ""
            End If
        ElseIf Mid(js, i, 1) = "\" And Mid(js, i + 1, 1) = Chr(&H22) Then
            s = s & "\" & Chr(&H22)
            i = i + 1
        Else
            s = s & Mid(js, i, 1)
        End If

        i = i + 1
    Wend

    JsonDecode = s
End Function



' Convert json data to valid vba Dictionary/Array structure.
'
Public Function Parse(json As String, Optional typeReset As Boolean = False) As Variant
    p = 2

    token = Tokenize(json)
    On Error GoTo errorHandler

    If token(1) = "{" Then
        Set Parse = ParseObj
    ElseIf token(1) = "[" Then
        Parse = ParseArr
    Else
        Err.Raise 1011, "JsonParser.Parse", "Invalid Json format."
    End If


    If typeReset Then
        If VarType(Parse) = vbObject Then
            Set Parse = Reset(Parse)
        Else
            Parse = Reset(Parse)
        End If
    End If

    Exit Function

errorHandler:
    Err.Raise 1011, "JsonParser.Parse", "Invalid Json format."
End Function



'-------------------------------------------------------------------
' Private functions
'-------------------------------------------------------------------



Private Function JsonEncodeEngine(ByVal js As String) As String
    Dim cstck As Stack ' Structural, Counter Stack
    Dim cp As Long ' Code point
    Dim s As String

    Set cstck = New Stack
    cstck.SetType ("Long")

    Do:
        p = p + 1

        cp = CLng(AscW(Mid(js, p, 1)))
        Select Case cp

            ' STRUCTURAL CHARACTERS
            Case sc.leftCurlyBracket:                           ' KEY: "{"
                PlaceChk cstck, cp
                sstck.Push (cp)
                cstck.Push (cp)
                s = s & Chr(cp) & JsonEncodeEngine(js) & Chr(translate(cp))
                sstck.Pop


            Case sc.leftSquareBracket:                          ' KEY: "["
                PlaceChk cstck, cp
                sstck.Push (cp)
                cstck.Push (cp)
                s = s & Chr(cp) & JsonEncodeEngine(js) & Chr(translate(cp))
                sstck.Pop


            Case sc.rightCurlyBracket:                          ' KEY: "}"
                ObjectChk cstck
                StackChk (cp)
                Exit Do


            Case sc.rightSquareBracket:                         ' KEY: "]"
                ArrayChk cstck
                StackChk (cp)
                Exit Do

            Case sc.comma:                                      ' KEY: ","
                PlaceChk cstck, cp
                'CommaChk cstck
                s = s + Chr(cp)
                cstck.Push (sc.comma)

            Case sc.colon:                                      ' KEY: ":"
                If sstck.ReadLast = sc.leftSquareBracket Then
                    Err.Raise 1021, "json.JsonEncode", _
                        "Syntax error. An array cannot contain " & _
                        "a colon '" & Chr(sc.colon) & "', at: " & p
                End If
                PlaceChk cstck, cp
                s = s + Chr(cp)
                cstck.Push (sc.colon)


            ' INSIGNIFICANT WHITESPACES
            Case ws.spac_e                                      ' KEY: " "
                s = s & ChrW(cp)

            Case ws.horizontalTab                               ' KEY: "\t"
                s = s & ChrW(cp)

            Case ws.lineFeed                                    ' KEY: "\n"
                s = s & ChrW(cp)

            Case ws.carrageReturn                               ' KEY: "\r"
                s = s & ChrW(cp)


            ' LITERAL NAMES
            Case &H74                                       ' KEY: "true"
                If Mid(js, p + 1, 1) <> &H72 Or _
                    Mid(js, p + 2, 1) <> &H75 Or _
                    Mid(js, p + 3, 1) <> &H65 Then
                    Err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "true"
                p = p + 3
                cstck.Push (cp)

            Case &H66                                           ' KEY: "false"
                If Mid(js, p + 1, 1) <> &H61 Or _
                    Mid(js, p + 2, 1) <> &H6C Or _
                    Mid(js, p + 3, 1) <> &H73 Or _
                    Mid(js, p + 4, 1) <> &H65 Then
                    Err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "false"
                p = p + 4
                cstck.Push (cp)

            Case &H6E                                           ' KEY: "null"
                If Mid(js, p + 1, 1) <> &H75 Or _
                    Mid(js, p + 2, 1) <> &H6C Or _
                    Mid(js, p + 3, 1) <> &H6C Then
                    Err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "null"
                p = p + 3
                cstck.Push (cp)


            ' STRING
            Case &H22:                                          ' KEY: '"'
                PlaceChk cstck, cp
                cstck.Push (cp)
                s = strHandler(s, js)


            ' NUMBER
            Case &H30, _
                &H31, _
                &H32, _
                &H33, _
                &H34, _
                &H35, _
                &H36, _
                &H37, _
                &H38, _
                &H39, _
                &H2D, _
                &H2B, _
                &H2E, _
                &H45, _
                &H65:                                           ' KEY:  "0", "1", "2", "3", "4",
                                                                '       "5", "6", "7", "8", "9",
                                                                '       "-", "+", ".", "e", "E"
                PlaceChk cstck, cp
                s = s + Chr(cp)
                ' Save, if it does not already exist: "0"
                If cstck.ReadLast <> &H30 Then cstck.Push (CLng(&H30))

            Case Else:                                          ' KEY: Other forbidden
                Err.Raise 1023, "json.JsonEncode", _
                    "Syntax error, forbidden character, at: " & _
                    p & Chr(10) & "Code point:  0x" & Right(&H30 & &H30 & &H30 & Hex(cp), 4)
        End Select

        If sstck.Count <> 0 Then
            If Len(js) = p Then Err.Raise 1024, "json.JsonEncode", _
                    "Syntax error. Missing '" & _
                    Chr(translate(sstck.ReadLast())) & "', at: " & p
        Else
            Exit Do
        End If
    Loop

    JsonEncodeEngine = s
End Function




' word processing algorithm
'
Private Function strHandler(ByVal s As String, js As String) As String
    s = s & Chr(&H22)
    Do
        p = p + 1
        cp = CInt(AscW(Mid(js, p, 1)))
        Select Case cp

            Case &H22:
                If Mid(js, p - 1, 1) <> &H5C Then               ' KEY: "\""
                    strHandler = s & Chr(&H22)
                    Exit Do
                Else
                    s = s & ChrW(cp)
                End If

            Case &H20 To &H22, _
                &H23 To &H5B, _
                &H5D To &HFFFF:                                 ' KEY: 0x20-21 / 0x23-5B / 0x5D-FFFF

                s = s & ChrW(cp)
            Case Else
                If gao Then
                    s = s & ChrW(cp)
                Else
                    s = s & Chr(&H5C) & Chr(&H75)
                    s = s & Right(&H30 & &H30 & &H30 & StrConv(Hex(cp), vbLowerCase), 4)
                End If
        End Select
    Loop
End Function




' post-process control
'
Private Function ObjectChk(cstck As Stack)
    ' Object contains max one key/value pair
    If cstck.Count > 0 And cstck.Count < 4 Then
        If cstck.Count = 1 Then
            Err.Raise 1025, "json.JsonEncode", _
                "Syntax error. Missing separator '" & Chr(sc.colon) & "', at: " & p
        ElseIf cstck.Count Mod 4 <> 3 Then
            Select Case (cstck.Count Mod 4)
                Case 0:
                    Err.Raise 1026, "json.JsonEncode", _
                        "Syntax error. Unexpected separator '" & Chr(sc.comma) & "', at: " & p
                Case 1:
                    Err.Raise 1025, "json.JsonEncode", _
                        "Syntax error. Expected separator '" & Chr(sc.colon) & "', at: " & p
                Case 2:
                    Err.Raise 1027, "json.JsonEncode", _
                        "Syntax error. Missing value for key, at: " & p
            End Select
        End If
    ' Object contains more key/value pairs
    ElseIf cstck.Count > 3 Then
        If cstck.Count Mod 4 = 0 Then
            Err.Raise 1027, "json.JsonEncode", _
                "Syntax error. Missing value for key, at: " & p
        End If
    End If
End Function


Private Function ArrayChk(cstck As Stack)
    If cstck.Count > 0 And cstck.Count Mod 2 <> 1 Then
        Err.Raise 1028, "json.JsonEncode", _
            "Syntax error. To mutch separator in array, at: " & p
    End If
End Function


' in-process control
'
Private Function PlaceChk(ByVal cstck As Stack, cp)
    ' Array
    If sstck.ReadLast = sc.leftSquareBracket Then
        If cp = sc.comma Then
            If cstck.ReadLast = sc.comma Then Err.Raise 1026, "json.JsonEncode", _
                "Syntax error. Unexpected separator '" & Chr(sc.comma) & "', at: " & p

        ElseIf cstck.Count Mod 2 = 1 And cstck.ReadLast() <> &H30 Then
            Err.Raise 1025, "json.JsonEncode", _
                "Syntax error. Expected separator '" & Chr(sc.comma) & "', at: " & p

        End If
    ' Object
    ElseIf sstck.ReadLast = sc.leftCurlyBracket Then
        Select Case cstck.Count Mod 4
            Case 0:
                If cp <> &H22 Then Err.Raise 1029, "json.JsonEncode", _
                    "Syntax error. Only string can be key of object, at: " & p
            Case 1:
                If cp <> sc.colon Then If cp <> &H22 Then Err.Raise 1025, "json.JsonEncode", _
                    "Syntax error. Expected separator '" & Chr(sc.colon) & "', at: " & p
            Case 2:
                ' Do nothing
            Case 3:
                If cp <> sc.comma And cstck.ReadLast() <> &H30 Then Err.Raise 1025, "json.JsonEncode", _
                    "Syntax error. Expected separator '" & Chr(sc.comma) & "', at: " & p
        End Select
    End If
End Function



Private Function StackChk(cp)
    If translate(cp) <> sstck.ReadLast Then
        Err.Raise 1022, "json.JsonEncode", _
            "Syntax error. Expected structural character '" & _
                Chr(translate(sstck.ReadLast)) & "', at: " & p
    End If
End Function



Private Function ParseObj() As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.dictionary")
    Dim e As Integer
    Do:
        Select Case token(p)
            Case "]":
                        Set ParseObj = dict
                        Exit Function
            Case "}":
                        Exit Do
            Case ",", ":":
                        ' do nothing
            Case Else:
                        If token(p + 2) = "[" Then      ' Add dictionary
                            e = p
                            p = p + 3
                            dict.Add key:=token(e), Item:=ParseArr()

                        ElseIf token(p + 2) = "{" Then  ' Add array
                            e = p
                            p = p + 3
                            dict.Add key:=token(e), Item:=ParseObj()

                        Else
                            dict.Add key:=token(p), Item:=token(p + 2)
                            p = p + 2
                        End If
        End Select
        p = p + 1
    Loop
    Set ParseObj = dict
End Function



Private Function ParseArr() As Variant
    Dim arr() As Variant
    Dim e As Integer
    e = 0
    Do:
        Select Case token(p)
            Case "}":
                        ' do nothing
            Case "{":
                        ReDim Preserve arr(e)
                        p = p + 1
                        Set arr(e) = ParseObj

            Case "[":
                        ReDim Preserve arr(e)
                        arr(e) = ParseArr

            Case "]":
                        Exit Do
            Case ",":
                        e = e + 1
            Case Else:
                        ReDim Preserve arr(e)
                        arr(e) = token(p)
        End Select
        p = p + 1
    Loop

    ParseArr = arr
End Function



Private Function Tokenize(s)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[ee][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = Rextract(s, Pattern, True)
End Function


Private Function Rextract(s, Pattern, Optional bGroup1bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .Ignorecase = True
    .Pattern = Pattern
    If .test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  Rextract = v
End Function



Private Function Reset(jObj As Variant) As Variant

    ' Dictionary
    If VarType(jObj) = vbObject Then
        Dim k As Variant
        For Each k In jObj.Keys()
            vSwitcher jObj, k
        Next k
        Set Reset = jObj
        Exit Function

    ' Variant()
    ElseIf VarType(jObj) = vbArray + vbVariant Then
        Dim i As Long
        For i = 0 To UBound(jObj)
            vSwitcher jObj, i
        Next
    Else
        If IsNumeric(jObj) Then
            jObj = CDec(jObj)
        ElseIf jObj = "true" Then
            jObj = True
        ElseIf jObj = "false" Then
            jObj = False
        ElseIf jObj = "null" Then
            jObj = Null
        End If
    End If

    Reset = jObj
End Function



' Because variant type, needed a switcher beetwen
' object and array definition.
'
Private Function vSwitcher(ByRef jObj As Variant, ByVal k As String)
    If VarType(jObj(k)) = vbObject Then
        Set jObj(k) = Reset(jObj(k))
    Else
        jObj(k) = Reset(jObj(k))
    End If
End Function


