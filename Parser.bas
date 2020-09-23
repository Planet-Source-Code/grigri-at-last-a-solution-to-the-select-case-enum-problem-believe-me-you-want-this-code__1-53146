Attribute VB_Name = "Parser"
Option Explicit

Public Function GetProcParams(sProcLine As String, arrParams() As String) As Boolean
    Dim i As Long
    Dim iStartParams As Long
    Dim iToken As Long
    Dim bWantRetVal As Boolean
    Dim sProcName As String
    
    Dim arrTokens() As String
    Dim arrTokens2() As String
    Dim sTemp As String
    Dim Index As Long
    
    If LenB(sProcLine) = 0 Then Exit Function
    
    iStartParams = InStr(sProcLine, "(")
    If iStartParams = 0 Then Exit Function
    
    arrTokens() = Split(Left$(sProcLine, iStartParams - 1), " ")
    iToken = 0
CHECK_PROC_TYPE:
    If arrTokens(iToken) = "Private" Then
        iToken = iToken + 1
        GoTo CHECK_PROC_TYPE
    End If
    If arrTokens(iToken) = "Public" Then
        iToken = iToken + 1
        GoTo CHECK_PROC_TYPE
    End If
    If arrTokens(iToken) = "Static" Then
        iToken = iToken + 1
        GoTo CHECK_PROC_TYPE
    End If
    If arrTokens(iToken) = "Sub" Then
        bWantRetVal = False
        GoTo START_ANALYZE_PROC
    End If
    If arrTokens(iToken) = "Function" Then
        bWantRetVal = True
        sProcName = arrTokens(UBound(arrTokens()))
        GoTo START_ANALYZE_PROC
    End If
    If arrTokens(iToken) = "Property" Then
        If arrTokens(iToken + 1) = "Get" Then
            bWantRetVal = True
            sProcName = arrTokens(UBound(arrTokens()))
            GoTo START_ANALYZE_PROC
        ElseIf arrTokens(iToken + 1) = "Let" Then
            bWantRetVal = False
            GoTo START_ANALYZE_PROC
        ElseIf arrTokens(iToken + 1) = "Set" Then
            bWantRetVal = False
            GoTo START_ANALYZE_PROC
        End If
    End If
    Exit Function
START_ANALYZE_PROC:
    ' From this point on, we're definately in a procedure
    GetProcParams = True
    ' Tokenize the parameter list
    ' (empty quotes first to avoid problems)
    arrTokens = Split(EmptyQuotes(Mid$(sProcLine, iStartParams + 1)), ", ")
    ' Setup the return values
    Index = UBound(arrParams) + 1
    ReDim Preserve arrParams(Index + UBound(arrTokens) * 2 + 3)
    If bWantRetVal Then
        ' The last 2 entries in the array will be
        ' the procedure name and its return value
        arrParams(UBound(arrParams) - 1) = sProcName
    Else
    End If
    
    ' Deal with the return value
    sTemp = arrTokens(UBound(arrTokens))
    i = InStrRev(sTemp, ") As ")
    If i = 0 Then
        ' No return value indicated
        If bWantRetVal Then
            arrParams(UBound(arrParams)) = "Variant"
        End If
        arrTokens(UBound(arrTokens)) = Left$(sTemp, Len(sTemp) - 1)
    Else
        If bWantRetVal Then
            arrParams(UBound(arrParams)) = Mid$(sTemp, i + 5, Len(sTemp) - 5)
        End If
        arrTokens(UBound(arrTokens)) = Left$(sTemp, i - 1)
    End If
    
    ' Deal with the parameters
    For i = 0 To UBound(arrTokens())
        arrTokens2 = Split(arrTokens(i), " ")
        iToken = 0
CHECK_PARAM_TOKEN:
        ' Check in case of error
        If iToken > UBound(arrTokens2) Then GoTo DO_NEXT_PARAM
        ' Skip specifiers
        If arrTokens2(iToken) = "Optional" Then iToken = iToken + 1: GoTo CHECK_PARAM_TOKEN
        If arrTokens2(iToken) = "ByVal" Then iToken = iToken + 1: GoTo CHECK_PARAM_TOKEN
        If arrTokens2(iToken) = "ByRef" Then iToken = iToken + 1: GoTo CHECK_PARAM_TOKEN
        ' Grab variable name
        arrParams(Index) = arrTokens2(iToken)
        If iToken + 2 > UBound(arrTokens2) Then
            arrParams(Index + 1) = "Variant"
        Else
            arrParams(Index + 1) = arrTokens2(iToken + 2)
        End If
        Index = Index + 2
DO_NEXT_PARAM:
    Next
End Function

Public Function GetDeclaredVars(sLine As String, arrVars() As String) As Boolean
    Dim arrTokens() As String
    
    If Left$(sLine, 4) = "Dim " Then
        arrTokens = Split(EmptyQuotes(Mid$(sLine, 5)), ", ")
        GetDeclaredVars = True
    End If
    If Left$(sLine, 8) = "Private " Then
        arrTokens = Split(EmptyQuotes(Mid$(sLine, 9)), ", ")
        GetDeclaredVars = True
    End If
    If Left$(sLine, 7) = "Public " Then
        arrTokens = Split(EmptyQuotes(Mid$(sLine, 8)), ", ")
        GetDeclaredVars = True
    End If
    If Left$(sLine, 7) = "Static " Then
        arrTokens = Split(EmptyQuotes(Mid$(sLine, 8)), ", ")
        GetDeclaredVars = True
    End If
    If Not GetDeclaredVars Then Exit Function
    
    Dim i As Long, j As Long
    Dim Index As Long
    Index = UBound(arrVars) + 1
    ReDim Preserve arrVars(Index + UBound(arrTokens) * 2 + 1)
    For i = 0 To UBound(arrTokens)
        j = InStr(arrTokens(i), " As ")
        If j = 0 Then
            arrVars(Index) = arrTokens(i)
            arrVars(Index + 1) = "Variant"
        Else
            arrVars(Index) = Trim$(Left$(arrTokens(i), j - 1))
            arrVars(Index + 1) = Trim$(Mid$(arrTokens(i), j + 4))
        End If
        Index = Index + 2
    Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : JoinAndTrimContinuedLines
' Purpose   : Check for continued lines, starting at the given line in the given module
'             Optionally return the last line, always return the joined lines, trimmed
'---------------------------------------------------------------------------------------
Public Function JoinAndTrimContinuedLines(cm As CodeModule, ByVal StartLine As Long, Optional ByRef EndLine As Long) As String
    Dim sLine As String
    EndLine = StartLine
    Do
        If EndLine > cm.CountOfLines Then Exit Function
        sLine = Trim$(cm.Lines(EndLine, 1))
        If Right$(sLine, 2) = " _" Then
            JoinAndTrimContinuedLines = JoinAndTrimContinuedLines & Left$(sLine, Len(sLine) - 1)
        Else
            JoinAndTrimContinuedLines = JoinAndTrimContinuedLines & sLine
            Exit Function
        End If
        EndLine = EndLine + 1
    Loop
End Function

Public Function GetLocalVars(cm As CodeModule, ByVal StartLine As Long) As String()
    Dim arrVars() As String
    ' Initialize the array
    arrVars = Split("", " ")
    
    Dim iLine As Long, sLine As String
    For iLine = StartLine To 1 Step -1
        If iLine > 1 Then
            ' Check for continued lines [ignore if it is]
            If Right$(cm.Lines(iLine - 1, 1), 2) = " _" Then GoTo SKIP_THIS_LINE
        End If
        sLine = StripVBComments(JoinAndTrimContinuedLines(cm, iLine))
        If GetProcParams(sLine, arrVars) Then
            ' Note for returning arrays from functions:
            ' The Exit/End Function statement must be directly after the
            ' assignment, as this makes the array transfered [fast].
            ' If there are any statements in between, the array
            ' is copied [slow] instead.
            GetLocalVars = arrVars
            Exit Function
        End If
        Call GetDeclaredVars(sLine, arrVars)
SKIP_THIS_LINE:
    Next
    GetLocalVars = arrVars
End Function

Public Function SearchLocalVars(cm As CodeModule, ByVal StartLine As Long, sVarName As String) As String
    Dim arrStrings() As String, i As Long
    arrStrings = GetLocalVars(cm, StartLine)
    For i = 0 To UBound(arrStrings) - 1 Step 2
        If arrStrings(i) = sVarName Then
            SearchLocalVars = arrStrings(i + 1)
            Exit Function
        End If
    Next
End Function

Public Function SearchModuleForVar(cm As CodeModule, sVarName As String, ByVal WantPrivate As Boolean) As String
    Dim mem As Member
    For Each mem In cm.Members
        If mem.Type <> vbext_mt_Event Then
            If Not (WantPrivate And mem.Scope = vbext_Private) Then
                If mem.Name = sVarName Then
                    SearchModuleForVar = SearchLocalVars(cm, mem.CodeLocation, sVarName)
                    If SearchModuleForVar <> "" Then Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function SearchProjectForVar(prj As VBProject, sVarName As String) As String
    Dim cmp As VBComponent
    On Error Resume Next
    For Each cmp In prj.VBComponents
        SearchProjectForVar = SearchModuleForVar(cmp.CodeModule, sVarName, False)
        If SearchProjectForVar <> "" Then Exit Function
    Next
End Function

Public Function GetWithVariable(cm As CodeModule, ByVal StartLine As Long)
    Dim i As Long, sLine As String, iWithLevel As Long
    If StartLine < 1 Then Exit Function
    For i = StartLine To 1 Step -1
        sLine = cm.Lines(i, 1)
        If Len(sLine) > 5 Then
            sLine = LTrim$(sLine)
            
            If Left$(sLine, 5) = "With " Then
                If iWithLevel = 0 Then
                    sLine = StripVBComments(JoinAndTrimContinuedLines(cm, i))
                    GetWithVariable = Mid$(sLine, 6)
                    If Left$(GetWithVariable, 1) = "." Then
                        ' I don't believe it! Someone's done a With inside
                        ' another With, just to piss us off - like this for example:
                        '     With SomeVar
                        '         With .SomeOtherVar
                        '            Select Case .YetAnotherOne
                        '            End Select
                        '         End With
                        '     End With
                        ' We're going to have to recurse
                        GetWithVariable = GetWithVariable(cm, i - 1) & GetWithVariable
                    End If
                    Exit Function
                Else
                    iWithLevel = iWithLevel - 1
                End If
            ElseIf Left$(sLine, 8) = "End With" Then
                iWithLevel = iWithLevel + 1
            End If
        End If
    Next
End Function

Public Function GetExpressionType(sExpr As String, cm As CodeModule, ByVal StartLine As Long) As String
    Dim i As Long, sVar As String, sTrail As String, sVarType As String
    ' Check if there is a space in the expression
    ' If their is [assuming literal quotes have been emptied]
    ' then the expression is complex, like "a + b"
    ' and therefore we can't easily deduce the type
    i = InStr(sExpr, " ")
    If i > 0 Then Exit Function
    ' Find the position of the first dot
    ' Remembering of course that literal quotes have been emptied
    ' If we hadn't done that then an expression like "Select Case col("blob.blob").Var"
    ' would wreck our plans...
    i = InStr(sExpr, ".")
    Select Case i
    Case 0
        ' Nice and easy, no trailing specifiers
        sVar = sExpr
        sTrail = ""
    Case 1
        ' We need to look for a "With"
        sVar = GetWithVariable(cm, StartLine)
        If sVar = "" Then Exit Function
        ' Append the expression to the with variable and try again
        GetExpressionType = GetExpressionType(EmptyQuotes(sVar) & sExpr, cm, StartLine)
        Exit Function
    Case Else
        sVar = Left$(sExpr, i - 1)
        sTrail = Mid$(sExpr, i + 1)
    End Select
    
    sVarType = GetVarType(sVar, cm, StartLine)
    If sTrail = "" Then
        GetExpressionType = sVarType
        Exit Function
    End If
    GetExpressionType = GetExprType2(cm, sVarType, sTrail)
    
End Function

Public Function GetExprType2(cm As CodeModule, sParentType As String, sChildName As String) As String
    Dim i As Long, j As Long, k As Long
    i = InStr(sChildName, ".")
    If i > 0 Then
        ' Argh! A complex expression
        ' We'll resolve it recursively
        GetExprType2 = GetExprType2(cm, GetExprType2(cm, sParentType, Left$(sChildName, i - 1)), Mid$(sChildName, i + 1))
        Exit Function
    End If
    
    On Error Resume Next
    ' Check if sParentType is a project-level class
    Dim cmp As VBComponent, mem As Member
    Set cmp = VBInstance.ActiveVBProject.VBComponents(sParentType)
    If Not cmp Is Nothing Then
        ' It is a class! Bravo :)
        ' Now we need to get the member
        For Each mem In cmp.CodeModule.Members
            If mem.Scope <> vbext_Private Then
                If mem.Name = sChildName Then
                    GetExprType2 = GetVarType(sChildName, cmp.CodeModule, mem.CodeLocation)
                    Exit Function
                End If
            End If
        Next
    End If
    
    ' Check if sParentType is a module-level type [search public and private]
    Dim arrStrings() As String
    arrStrings = GetModuleTypes(cm, True, True)
    Err.Clear
    i = UBound(arrStrings)
    If Err.Number = 0 Then
    For i = 0 To UBound(arrStrings) - 1 Step 2
        If arrStrings(i) = sParentType Then
            ' It's a match! Get the type of the child...
            j = InStr(arrStrings(i + 1), "|" & sChildName & ":")
            If j > 0 Then
                k = InStr(j + 1, arrStrings(i + 1), "|")
                If k = 0 Then
                    ' Last one
                    GetExprType2 = Mid$(arrStrings(i + 1), j + Len(sChildName) + 2)
                Else
                    ' In the middle...
                    GetExprType2 = Mid$(arrStrings(i + 1), j + Len(sChildName) + 2, k - j - Len(sChildName) - 2)
                End If
                Exit Function
            End If
        End If
    Next
    End If
        
    ' Check if it's a Type declared in another module [search only public]
    For Each cmp In VBInstance.ActiveVBProject.VBComponents
        arrStrings = GetModuleTypes(cmp.CodeModule, cmp.Type = vbext_ct_StdModule, False)
        Err.Clear
        i = UBound(arrStrings)
        If Err.Number = 0 Then
        For i = 0 To UBound(arrStrings) - 1 Step 2
            If arrStrings(i) = sParentType Then
                ' It's a match! Get the type of the child...
                j = InStr(arrStrings(i + 1), "|" & sChildName & ":")
                If j > 0 Then
                    k = InStr(j + 1, arrStrings(i + 1), "|")
                    If k = 0 Then
                        ' Last one
                        GetExprType2 = Mid$(arrStrings(i + 1), j + Len(sChildName) + 2)
                    Else
                        ' In the middle...
                        GetExprType2 = Mid$(arrStrings(i + 1), j + Len(sChildName) + 2, k - j - Len(sChildName) - 2)
                    End If
                    Exit Function
                End If
            End If
        Next
        End If
    Next
    
    ' Ok, time to break out the typelibs
    On Error Resume Next
    Dim ref As Reference, lib As TypeLibInfo, ti As TypeInfo, mi As MemberInfo, ii As InterfaceInfo
    
'    ' Check if the Parent is actually the name of a library reference
'    Set ref = VBInstance.ActiveVBProject.References(sParentType)
'    If Not ref Is Nothing Then
'        Set lib = Nothing
'        Set lib = TypeLibInfoFromFile(ref.FullPath)
'        If Not lib Is Nothing Then
'        End If
'    End If
    
    ' Ok, check all referenced typelibs for it
    For Each ref In VBInstance.ActiveVBProject.References
        Set lib = Nothing
        Set lib = TypeLibInfoFromFile(ref.FullPath)
        If Not lib Is Nothing Then
            For Each ti In lib.TypeInfos
                Select Case ti.TypeKind
                Case TKIND_INTERFACE, TKIND_RECORD
                    If ti.Name = sParentType Then
                        For Each mi In ti.Members
                            If mi.Name = sChildName Then
                                ' Got it!
                                GetExprType2 = mi.ReturnType.TypeInfo.Name
                                Exit Function
                            End If
                        Next
                    End If
                Case TKIND_COCLASS
                    If ti.Name = sParentType Then
                        For Each ii In ti.Interfaces
                            For Each mi In ii.Members
                                If mi.Name = sChildName Then
                                    GetExprType2 = mi.ReturnType.TypeInfo.Name
                                    Exit Function
                                End If
                            Next
                        Next
                    End If
                End Select
            Next
        End If
    Next
End Function

Public Function GetVarType(sVarName As String, cm As CodeModule, ByVal StartLine As Long) As String
    GetVarType = SearchLocalVars(cm, StartLine, sVarName)
    If GetVarType <> "" Then Exit Function
    
    GetVarType = SearchModuleForVar(cm, sVarName, True)
    If GetVarType <> "" Then Exit Function
    
    GetVarType = SearchProjectForVar(VBInstance.ActiveVBProject, sVarName)
    If GetVarType <> "" Then Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : EmptyQuotes
' Purpose   : Given a string in VB statement format, returns the same string
'             but with any quotes emptied, for easier processing.
'---------------------------------------------------------------------------------------
Public Function EmptyQuotes(sLine As String) As String
    Dim i As Long, j As Long, InQuote As Boolean
    i = 1: j = 1
    InQuote = False
    Do
        If Mid$(sLine, j, 1) = """" Then
            If InQuote Then
                InQuote = False
            Else
                EmptyQuotes = EmptyQuotes & Mid$(sLine, i, j - i + 1)
                InQuote = True
            End If
            j = j
            i = j
        End If
        j = j + 1
    Loop While j < Len(sLine)
    If Not InQuote Then
        EmptyQuotes = EmptyQuotes & Mid$(sLine, i)
    End If
End Function

Public Function GetModuleEnums(cm As CodeModule, DefaultIsPublic As Boolean, WantPrivate As Boolean) As String()
    Dim i As Long, sLine As String, j As Long
    Dim arrEnums() As String
    Dim CurrentIndex As Long
    Dim bInEnum As Boolean
    ReDim arrEnums(10)
    CurrentIndex = 0
    
    i = 1
    While i <= cm.CountOfDeclarationLines
        sLine = StripVBComments(JoinAndTrimContinuedLines(cm, i, i))
        If bInEnum Then
            If sLine = "End Enum" Then
                bInEnum = False
                CurrentIndex = CurrentIndex + 1
                If CurrentIndex > UBound(arrEnums) Then
                    ReDim Preserve arrEnums(UBound(arrEnums) + 11)
                End If
            Else
                j = InStr(sLine, "=")
                If j = 0 Then
                    ' No explicit value
                    arrEnums(CurrentIndex) = arrEnums(CurrentIndex) & "|" & sLine
                Else
                    ' We don't want the value part
                    arrEnums(CurrentIndex) = arrEnums(CurrentIndex) & "|" & Trim$(Left$(sLine, j - 1))
                End If
            End If
        Else
            If sLine Like "*Enum *" Then
                ' It looks like an enum declaration
                If Left$(sLine, 5) = "Enum " Then
                    If DefaultIsPublic Or WantPrivate Then
                        arrEnums(CurrentIndex) = Mid$(sLine, 6)
                        CurrentIndex = CurrentIndex + 1
                        bInEnum = True
                        GoTo DO_NEXT_LINE
                    End If
                End If
                If Left$(sLine, 12) = "Public Enum " Then
                    arrEnums(CurrentIndex) = Mid$(sLine, 13)
                    CurrentIndex = CurrentIndex + 1
                    bInEnum = True
                    GoTo DO_NEXT_LINE
                End If
                If Left$(sLine, 13) = "Private Enum " Then
                    If WantPrivate Then
                        arrEnums(CurrentIndex) = Mid$(sLine, 14)
                        CurrentIndex = CurrentIndex + 1
                        bInEnum = True
                        GoTo DO_NEXT_LINE
                    End If
                End If
            End If
        End If
DO_NEXT_LINE:
        i = i + 1
    Wend
    
    ReDim Preserve arrEnums(CurrentIndex - 1)
    GetModuleEnums = arrEnums
End Function

Public Function StripVBComments(ByVal sLine As String)
    Dim i As Long, j As Long, InQuotes As Boolean
    i = InStr(sLine, "'")
    ' No apostrophe => no comments => no problem
    If i = 0 Then StripVBComments = sLine: Exit Function
    ' Check if there are any quotes in the line
    j = InStr(sLine, """")
    ' There are no quotes, so the apostrophe we found marks the comment
    If j = 0 Then StripVBComments = Left$(sLine, i - 1)
    
    ' This is the nasty bit
    ' Basically, we need to process the string char by char, left to right
    ' If we find an apostrophe not inside quotes, then it's the comment
    ' marker and we can all go home
    InQuotes = False
    For j = 1 To i
        Select Case AscW(Mid$(sLine, j, 1))
        Case 34 ' Double-quote => string
            InQuotes = Not InQuotes
        Case 39 ' Single-quote
            If Not InQuotes Then
                ' Hurrah! We've got the beggar
                StripVBComments = Left$(sLine, j - 1)
            End If
        End Select
    Next
End Function

Public Function GetSelectCaseVariable(cm As CodeModule, ByVal StartLine As Long)
    Dim i As Long, sTempLine As String
    Dim iSwitchLevel As Long
    For i = StartLine To 1 Step -1
        sTempLine = Trim$(cm.Lines(i, 1))
        If Left$(sTempLine, 12) = "Select Case " Then
            If iSwitchLevel = 0 Then
                GetSelectCaseVariable = Mid$(sTempLine, 13)
                Exit For
            Else
                iSwitchLevel = iSwitchLevel - 1
            End If
        End If
        If Left$(sTempLine, 11) = "End Select" Then
            iSwitchLevel = iSwitchLevel + 1
        End If
    Next
End Function

Public Function GetEnumValuesForCodePane(cp As CodePane) As String
    Dim cm As CodeModule
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim sSwitchVar As String
    Dim sSwitchVarType As String

    cp.GetSelection y1, x1, y2, x2
    ' Ignore spanning selections
    If x1 <> x2 Then Exit Function
    If y1 <> y2 Then Exit Function
    ' If we're at the first column then the line cannot
    ' begin with "Case"
    If x1 = 0 Then Exit Function
    
    Set cm = cp.CodeModule
    ' Check if the current line begins with "Case" (not including spaces)
    ' Remember that this line is being typed, so we can't guarantee proper case
    If UCase$(Left$(LTrim$(cm.Lines(y1, 1)), 5)) <> "CASE " Then Exit Function
    
    ' Find the switch variable
    
    sSwitchVar = GetSelectCaseVariable(cm, y1 - 1)
    If sSwitchVar = "" Then Exit Function
    
    'OutputDebugString "SwitchVariable = " & sSwitchVar & vbCrLf
    
    ' Empty any quotes inside it [for easier processing]
    sSwitchVar = EmptyQuotes(sSwitchVar)
    
    'OutputDebugString "Unquoted SwitchVariable = " & sSwitchVar & vbCrLf
    
    ' Resolve it to its type (as a string)
    sSwitchVarType = GetExpressionType(sSwitchVar, cm, y1 - 1)
    If sSwitchVarType = "" Then Exit Function
        
    'OutputDebugString "SwitchVariable Type = " & sSwitchVarType & vbCrLf
    
    GetEnumValuesForCodePane = SearchForEnum(sSwitchVarType, VBInstance.ActiveVBProject, cm)
End Function

Public Function SearchForEnum(sEnumName As String, prj As VBProject, cm As CodeModule) As String
    If sEnumName = "Variant" Then Exit Function
    If sEnumName = "Long" Then Exit Function
    If sEnumName = "Integer" Then Exit Function
    If sEnumName = "Byte" Then Exit Function
    If sEnumName = "String" Then Exit Function
    If sEnumName = "Date" Then Exit Function
    If sEnumName = "Currency" Then Exit Function
    If sEnumName = "Boolean" Then Exit Function
    If sEnumName = "Object" Then Exit Function
    
    SearchForEnum = SearchModuleForEnum(cm, sEnumName, True)
    If SearchForEnum <> "" Then Exit Function
    
    SearchForEnum = SearchProjectForEnum(prj, sEnumName, True)
    If SearchForEnum <> "" Then Exit Function
    
    SearchForEnum = SearchRefsForEnum(prj.References, sEnumName)
    If SearchForEnum <> "" Then Exit Function
End Function

Public Function SearchModuleForEnum(cm As CodeModule, sEnumName As String, ByVal WantPrivate As Boolean) As String
    Dim arrStrings() As String, i As Long
    On Error GoTo ERROR_HANDLER
    
    arrStrings = GetModuleEnums(cm, True, WantPrivate)
    For i = 0 To UBound(arrStrings) - 1 Step 2
        If arrStrings(i) = sEnumName Then
            SearchModuleForEnum = arrStrings(i + 1)
            Exit Function
        End If
    Next
    
ERROR_HANDLER:
End Function

Public Function SearchProjectForEnum(prj As VBProject, sEnumName As String, ByVal WantPrivate As Boolean) As String
    Dim cmp As VBComponent, cm As CodeModule
    On Error Resume Next
    For Each cmp In prj.VBComponents
        Set cm = Nothing
        Set cm = cmp.CodeModule
        If cm Is Nothing Then GoTo NEXT_COMPONENT
        
        If Not WantPrivate Then
            Select Case cmp.Type
            Case vbext_ct_ClassModule
                ' Exclude private class modules
                If cmp.Properties("Instancing") = 1 Then GoTo NEXT_COMPONENT
            Case vbext_ct_ActiveXDesigner, vbext_ct_UserControl
                ' Exclude private components
                If Not cmp.Properties("Public") Then GoTo NEXT_COMPONENT
            Case Else
                GoTo NEXT_COMPONENT
            End Select
        End If
        SearchProjectForEnum = SearchModuleForEnum(cm, sEnumName, False)
        If SearchProjectForEnum <> "" Then Exit Function
NEXT_COMPONENT:
    Next
End Function

Public Function SearchRefsForEnum(Refs As References, sEnumName As String) As String
    Dim ref As Reference
    For Each ref In Refs
        SearchRefsForEnum = SearchRefForEnum(ref, sEnumName)
        If SearchRefsForEnum <> "" Then Exit Function
    Next
End Function

Public Function SearchRefForEnum(ref As Reference, sEnumName As String) As String
    Dim lib As TypeLibInfo, ti As TypeInfo, mi As MemberInfo
    On Error GoTo ERROR_HANDLER
    Set lib = TypeLibInfoFromFile(ref.FullPath)
    If lib Is Nothing Then GoTo ERROR_HANDLER
    
    For Each ti In lib.TypeInfos
        If ti.TypeKind = TKIND_ENUM Then
            If ti.Name = sEnumName Then
                SearchRefForEnum = ""
                For Each mi In ti.Members
                    SearchRefForEnum = SearchRefForEnum & "|" & mi.Name
                Next
                Exit Function
            End If
        End If
    Next
    
ERROR_HANDLER:
End Function

Public Function GetModuleTypes(cm As CodeModule, DefaultIsPublic As Boolean, WantPrivate As Boolean) As String()
    Dim i As Long, sLine As String, j As Long
    Dim arrTypes() As String
    Dim CurrentIndex As Long
    Dim bInType As Boolean
    ReDim arrTypes(10)
    CurrentIndex = 0
    
    i = 1
    While i <= cm.CountOfDeclarationLines
        sLine = StripVBComments(JoinAndTrimContinuedLines(cm, i, i))
        If bInType Then
            If sLine = "End Type" Then
                bInType = False
                CurrentIndex = CurrentIndex + 1
                If CurrentIndex > UBound(arrTypes) Then
                    ReDim Preserve arrTypes(UBound(arrTypes) + 11)
                End If
            Else
                j = InStr(sLine, " As ")
                If j = 0 Then
                    ' No explicit type => Variant
                    arrTypes(CurrentIndex) = arrTypes(CurrentIndex) & "|" & sLine & ":Variant"
                Else
                    ' Extract the name and value, concatenate
                    arrTypes(CurrentIndex) = arrTypes(CurrentIndex) & "|" & Trim$(Left$(sLine, j - 1)) & ":" & Trim$(Mid$(sLine, j + 4))
                End If
            End If
        Else
            If sLine Like "*Type *" Then
                ' It looks like a Type declaration
                If Left$(sLine, 5) = "Type " Then
                    If DefaultIsPublic Or WantPrivate Then
                        arrTypes(CurrentIndex) = Mid$(sLine, 6)
                        CurrentIndex = CurrentIndex + 1
                        bInType = True
                        GoTo DO_NEXT_LINE
                    End If
                End If
                If Left$(sLine, 12) = "Public Type " Then
                    arrTypes(CurrentIndex) = Mid$(sLine, 13)
                    CurrentIndex = CurrentIndex + 1
                    bInType = True
                    GoTo DO_NEXT_LINE
                End If
                If Left$(sLine, 13) = "Private Type " Then
                    If WantPrivate Then
                        arrTypes(CurrentIndex) = Mid$(sLine, 14)
                        CurrentIndex = CurrentIndex + 1
                        bInType = True
                        GoTo DO_NEXT_LINE
                    End If
                End If
            End If
        End If
DO_NEXT_LINE:
        i = i + 1
    Wend
    
    ReDim Preserve arrTypes(CurrentIndex - 1)
    GetModuleTypes = arrTypes
End Function

Public Function Test2()
    Dim s As String
    s = "A = ""string"" : Z = ""harry said """"hi"""" to me"": B = 4"
    Debug.Print "'" & s & "'"
    Debug.Print "'" & EmptyQuotes(s) & "'"
End Function

Public Function Test()
    Dim arr() As String, s As String
    s = "Public Function JoinAndTrimContinuedLines(cm As CodeModule, Optional ByVal StartLine As Long = 2) As String"
    Debug.Print "Getting Params from:"
    Debug.Print s
    Debug.Print "-----------------------"
    If GetProcParams(s, arr) Then
        Debug.Print Join(arr, ",")
    Else
        Debug.Print "Not a function"
    End If
End Function
