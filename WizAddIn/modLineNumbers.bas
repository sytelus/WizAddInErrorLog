Attribute VB_Name = "modLineNumbers"
Option Explicit

Public Sub AddLineNumbersToProjects(ByVal voVBInstance As VBIDE.VBE)
    Call AddOrRemoveLineNumbers(voVBInstance, True, False, False)
End Sub

Public Sub RemoveLineNumbersFromProjects(ByVal voVBInstance As VBIDE.VBE)
    Call AddOrRemoveLineNumbers(voVBInstance, False, False, False)
End Sub

Public Sub AddErrorHandlerToProjects(ByVal voVBInstance As VBIDE.VBE)
    Call AddOrRemoveLineNumbers(voVBInstance, False, True, True)
End Sub

Private Sub AddOrRemoveLineNumbers(ByVal voVBInstance As VBIDE.VBE, ByVal vblnAddOrRemoveLineNumbers As Boolean, ByVal vblnDoNotChangeLineNumbers As Boolean, ByVal vblnAddErrorHandlers As Boolean)
    
    On Error GoTo ErrorTrap
    
    Dim sReadOnlyModuleNames As String
    
    sReadOnlyModuleNames = vbNullString 'Store module names which were readonly to display in error message
    
    Dim oVBProject As VBIDE.VBProject
    For Each oVBProject In voVBInstance.VBProjects  'For each project in group
        Dim oVBComponent As VBIDE.VBComponent
        For Each oVBComponent In oVBProject.VBComponents 'For each component
            'Get the code module for this component
            Dim oCodeModule As VBIDE.CodeModule
            Set oCodeModule = oVBComponent.CodeModule
            
            'If code module is available
            If Not (oCodeModule Is Nothing) Then
                'Go through each of the methods.
                Dim lVBMethodIndex As Long
                lVBMethodIndex = 1
                'Use Do instead of For because method count might change in between
                Do While lVBMethodIndex <= oCodeModule.Members.Count
                    Dim oVBMethod As VBIDE.Member
                    'Get the method object
                    Set oVBMethod = oCodeModule.Members(lVBMethodIndex)
                    'If this member is constant, variable, event - ignore it
                    Dim bIsVBProcedure As Boolean
                    Dim eVBMemberType As vbext_ProcKind
                    Call GetVBCodeMemberType(oCodeModule, oVBMethod, bIsVBProcedure, eVBMemberType)
                     
                    If bIsVBProcedure = True Then
                        
                        'Now we know whether the code member is actually a method/property or not.
                        'We need to put line numbers and error handlers ONLY in methods and properties.
                        If bIsVBProcedure Then
                            
                            'VB bug workaround: Due to bug in VB, Members Collection does not includes Let if there also exist Get property. Here's the worka around
                            Dim lPropertyTypeIndex As Long
                            Dim lPropertyTypeCount As Long
                            
                            If ((eVBMemberType = vbext_pk_Get) Or (eVBMemberType = vbext_pk_Let) Or (eVBMemberType = vbext_pk_Set)) And (oVBMethod.Type = vbext_mt_Property) Then
                                lPropertyTypeCount = 3
                            Else
                                lPropertyTypeCount = 1
                            End If
                            
                            For lPropertyTypeIndex = 1 To lPropertyTypeCount
                                Dim lMethodStartLineFromIDEModel As Long
                                Dim lMethodLineIndex As Long
                                On Error Resume Next
                                lMethodStartLineFromIDEModel = oCodeModule.ProcStartLine(oVBMethod.Name, eVBMemberType)
                                If Err.Number = 35 Then 'Sub or Function does not exist
                                    'Let or Set does not exist
                                    Err.Clear
                                    GoTo NextPropertyType
                                ElseIf Err.Number <> 0 Then
                                    GoTo ErrorTrap
                                End If
                                On Error GoTo ErrorTrap
                                
                                Dim lMethodLineScaneIndexStart As Long
                                Dim lMethodLineScaneIndexStop As Long
                                Call GetMethodStartStopScanLineNumbers(oCodeModule, oVBMethod, vblnAddOrRemoveLineNumbers, vblnAddErrorHandlers, _
                                    lMethodLineScaneIndexStart, lMethodLineScaneIndexStop)
                                
                                '
                                'Now we have the line numbers between which we should browse through statements
                                '
                                
                                'Add error handler logic
                                If vblnAddErrorHandlers = True Then
                                    Dim bIsOkToPutErrorHandler As Boolean
                                    bIsOkToPutErrorHandler = IsOkToPutErrorHandler(oCodeModule, oVBMethod, eVBMemberType, lMethodLineScaneIndexStart, lMethodLineScaneIndexStop)
                                    If bIsOkToPutErrorHandler = True Then
                                        Call InsertErrorHandlerCode(oVBProject, oVBComponent, oCodeModule, oVBMethod, lMethodStartLineFromIDEModel, eVBMemberType, lMethodLineScaneIndexStart, lMethodLineScaneIndexStop)
                                    Else
                                        'add bookmark here
                                    End If
                                Else
                                    'Don't add error handlers because parameter is true
                                End If
                                
                                If vblnDoNotChangeLineNumbers = False Then
                                    Dim bSelectStatementStarted As Boolean
                                    bSelectStatementStarted = False
                                    Dim bSplittedLineStarted As Boolean 'Lines ending with _ are splitted ones
                                    bSplittedLineStarted = False
                                    
                                    Call AddRemoveLineNumbersInMethod(oCodeModule, oVBMethod, vblnAddOrRemoveLineNumbers, lMethodLineScaneIndexStart, _
                                        lMethodLineScaneIndexStop, bSelectStatementStarted, bSplittedLineStarted)

                                Else
                                    'Do not add line numbers because parameter is false
                                End If
                                
NextPropertyType:
                                Select Case eVBMemberType
                                    Case vbext_pk_Get
                                        eVBMemberType = vbext_pk_Let
                                    Case vbext_pk_Let
                                        eVBMemberType = vbext_pk_Set
                                    Case vbext_pk_Set
                                        eVBMemberType = vbext_pk_Get
                                End Select
                                
                            Next lPropertyTypeIndex
                        End If
                    End If
                    lVBMethodIndex = lVBMethodIndex + 1
                Loop
            End If
ResumeNextComponent:
        Next oVBComponent
    Next oVBProject
    
    If sReadOnlyModuleNames <> vbNullString Then
        Err.Raise 1000, , "Line numbers can not be add to/removed from following modules because they are Read Only: " & vbCrLf & sReadOnlyModuleNames
    End If
    
Exit Sub
ErrorTrap:
    If Err.Number = 40198 Then  'Can not edit module
        sReadOnlyModuleNames = IIf(sReadOnlyModuleNames <> vbNullString, ", ", vbNullString) & sReadOnlyModuleNames & oVBComponent.Name
        Resume ResumeNextComponent
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetMethodStartLine(ByVal voCodeModule As CodeModule, ByVal voCodeMember As Member) As String
    Dim sMethodDefinationLine As String
    Dim lSearchStartLineOffset As Long
    lSearchStartLineOffset = 0
    Do
        sMethodDefinationLine = Trim(voCodeModule.Lines(voCodeMember.CodeLocation + lSearchStartLineOffset, 1))
        lSearchStartLineOffset = lSearchStartLineOffset + 1
    Loop While ((sMethodDefinationLine = vbNullString) Or (Left(sMethodDefinationLine, 1) = "'"))
    
    GetMethodStartLine = sMethodDefinationLine
End Function

Public Sub AddErrorHandlerToProjects_not_used(ByVal voVBInstance As VBIDE.VBE)
    
    On Error GoTo ErrorTrap
    
    Dim oVBProject As VBIDE.VBProject
    Dim oVBComponent As VBIDE.VBComponent
    Dim oCodeModule As VBIDE.CodeModule
    Dim oVBMethod As VBIDE.Member
    Dim lMethodStartLine As Long
    Dim lMethodLineCount As Long
    Dim lMethodLineIndex As Long
    Dim sOriginalLineNumber As String
    Dim lFirstWordStart As Long
    Dim lFirstWordLen As Long
    Dim sMethodLine As String
    Dim bLineNumberExist As Boolean
    Dim evptVBMethodType As vbext_ProcKind
    Dim bValidMethod As Boolean
    Dim bManualLineNumber As Boolean
    Dim sReadOnlyModuleNames As String
    
    sReadOnlyModuleNames = vbNullString
    For Each oVBProject In voVBInstance.VBProjects
        For Each oVBComponent In oVBProject.VBComponents
            Set oCodeModule = oVBComponent.CodeModule
            If Not (oCodeModule Is Nothing) Then
                For Each oVBMethod In oCodeModule.Members
                    If (oVBMethod.Type = vbext_mt_Method) Or (oVBMethod.Type = vbext_mt_Property) Then
                        bValidMethod = True
                        If oVBMethod.Type <> vbext_mt_Property Then
                            evptVBMethodType = vbext_pk_Proc
                        Else
                            Dim sMethodDefinationLine As String
                            Dim oWords As New Collection
                            Dim lSearchStartLineOffset As Long
                            lSearchStartLineOffset = 0
                            Do
                                sMethodDefinationLine = Trim(oCodeModule.Lines(oVBMethod.CodeLocation + lSearchStartLineOffset, 1))
                                lSearchStartLineOffset = lSearchStartLineOffset + 1
                            Loop While ((sMethodDefinationLine = vbNullString) Or (Left(sMethodDefinationLine, 1) = "'"))  'Skipp comment and blank lines
                            
                            Call MakeWordList(sMethodDefinationLine, oWords)
                            If oWords.Count >= 2 Then   'Check the second word
                                Select Case LCase(oWords(2))
                                    Case "get"
                                        evptVBMethodType = vbext_pk_Get
                                    Case "let"
                                        evptVBMethodType = vbext_pk_Let
                                    Case "set"
                                        evptVBMethodType = vbext_pk_Set
                                    Case Else   'Check the 3rd word
                                        If oWords.Count >= 3 Then
                                            Select Case LCase(oWords(3))
                                                Case "get"
                                                    evptVBMethodType = vbext_pk_Get
                                                Case "let"
                                                    evptVBMethodType = vbext_pk_Let
                                                Case "set"
                                                    evptVBMethodType = vbext_pk_Set
                                                Case Else
                                                    bValidMethod = False
                                            End Select
                                        Else
                                            bValidMethod = False
                                        End If
                                End Select
                            Else
                                bValidMethod = False
                            End If
                            Set oWords = Nothing
                        End If
                        
                        If bValidMethod Then
                            
                            'Due to bug in VB, Members Collection does not includes Let if there also exist Get property. Here's the worka around
                            Dim lPropertyTypeIndex As Long
                            Dim lPropertyTypeCount As Long
                            
                            If ((evptVBMethodType = vbext_pk_Get) Or (evptVBMethodType = vbext_pk_Let) Or (evptVBMethodType = vbext_pk_Set)) And (oVBMethod.Type = vbext_mt_Property) Then
                                lPropertyTypeCount = 3
                            Else
                                lPropertyTypeCount = 1
                            End If
                            
                            For lPropertyTypeIndex = 1 To lPropertyTypeCount
                                
                                Dim sPropertyTypeString As String
                                If oVBMethod.Type = vbext_mt_Property Then
                                    Select Case evptVBMethodType
                                        Case vbext_pk_Get
                                            sPropertyTypeString = "Get_"
                                        Case vbext_pk_Let
                                            sPropertyTypeString = "Let_"
                                        Case vbext_pk_Set
                                            sPropertyTypeString = "Set_"
                                        Case Else
                                            sPropertyTypeString = vbNullString
                                    End Select
                                Else
                                    sPropertyTypeString = vbNullString
                                End If
                                
                                On Error Resume Next
                                lMethodStartLine = oCodeModule.ProcStartLine(oVBMethod.Name, evptVBMethodType)
                                If Err.Number = 35 Then 'Sub or Function does not exist
                                    'Let or Set does not exist
                                    Err.Clear
                                    GoTo NextPropertyType
                                ElseIf Err.Number <> 0 Then
                                    GoTo ErrorTrap
                                End If
                                On Error GoTo ErrorTrap
                                lMethodLineCount = oCodeModule.ProcCountLines(oVBMethod.Name, evptVBMethodType)
                                Dim bSplittedLineStarted As Boolean 'Lines ending with _ are splitted ones
                                Dim bSelectStatementStarted As Boolean
                                Dim bThisIsCaseStatement As Boolean
                                Dim sTrimmedLine As String
                                
                                bSplittedLineStarted = False
                                bSelectStatementStarted = True
                                bThisIsCaseStatement = False
                                
                                'Find the last line of method
                                'Look for the non blank/non comment line from the end of the procedure
                                Dim lMethodActualLastLineNumber As Long
                                For lMethodActualLastLineNumber = (lMethodStartLine + lMethodLineCount - 1) To (lMethodStartLine + 1) Step -1
                                    sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualLastLineNumber, 1))
                                    If (sTrimmedLine <> vbNullString) _
                                        And (Left(sTrimmedLine, 1) <> "'") _
                                        And (Right(sTrimmedLine, 1) <> "_") _
                                        And (Left(sTrimmedLine, 1) <> "#") _
                                        Then
                                        Exit For
                                    End If
                                Next lMethodActualLastLineNumber
                                
                                'Find the method start line
                                Dim lMethodActualStartLineNumber As Long
                                For lMethodActualStartLineNumber = lMethodStartLine To lMethodActualLastLineNumber - 1
                                    sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualStartLineNumber, 1))
                                    If (sTrimmedLine <> vbNullString) _
                                        And (Left(sTrimmedLine, 1) <> "'") _
                                        And (Right(sTrimmedLine, 1) <> "_") _
                                        And (Left(sTrimmedLine, 1) <> "#") _
                                        Then
                                        Exit For
                                    End If
                                Next lMethodActualStartLineNumber
                                
                                'Check if error handler already exists OR refrence to Err.Number or "'#NO_ERROR_HANDLER" is there
                                Dim bDontPutErrorHandler As Boolean
                                Dim bEmptyMethod As Boolean
                                Dim lMethodActualLineCount As Long
                                Dim bContinueChecks As Boolean
                                Dim bMethodMayContainCalls As Boolean
                                
                                bDontPutErrorHandler = False
                                bEmptyMethod = True
                                bContinueChecks = True
                                bMethodMayContainCalls = False
                                
                                lMethodActualLineCount = 0
                                For lMethodLineIndex = (lMethodActualStartLineNumber + 1) To (lMethodActualLastLineNumber - 1)
                                    sMethodLine = Trim(oCodeModule.Lines(lMethodLineIndex, 1))
                                    If UCase(sMethodLine) = "'#USE_ERROR_HANDLER" Then
                                        bDontPutErrorHandler = False
                                        bContinueChecks = False
                                    ElseIf (UCase(sMethodLine) = "'#NO_ERROR_HANDLER") Or (UCase(sMethodLine) = "REM *** FailSafe SKIP") Then
                                        bDontPutErrorHandler = True
                                        bContinueChecks = False
                                    End If
                                    
                                    If (sMethodLine <> vbNullString) And (Left(sMethodLine, 1) <> "'") Then
                                        bEmptyMethod = False
                                        lMethodActualLineCount = lMethodActualLineCount + 1
                                        If Not bMethodMayContainCalls Then
                                            If ((InStr(1, sMethodLine, "(") <> 0) And (InStr(1, sMethodLine, ")") <> 0)) Or (InStr(1, sMethodLine, "call ") <> 0) Then
                                                bMethodMayContainCalls = True
                                            End If
                                        End If
                                        If bContinueChecks Then
                                            If InStr(1, sMethodLine, "err.", vbTextCompare) <> 0 Then
                                                bDontPutErrorHandler = True
                                                'Don't exit for loop. See next lines for commands.
                                            End If
                                            'Remove line number if any
                                            Dim oFirstWords As New Collection
                                            Call MakeWordList(sMethodLine, oFirstWords)
                                                'If first and second word is On Error
                                                If oFirstWords.Count >= 2 Then
                                                    If (oFirstWords(1) = "On") And (oFirstWords(2) = "Error") Then
                                                        bDontPutErrorHandler = True
                                                    Else
                                                        If oFirstWords.Count >= 3 Then
                                                            'First is line number and then On Error
                                                            If (oFirstWords(2) = "On") And (oFirstWords(3) = "Error") Then
                                                                bDontPutErrorHandler = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Set oFirstWords = Nothing
                                        End If
                                    End If
                                Next lMethodLineIndex
                                
                                'VB bug: For API declarations in code VB returns lMethodStartLine=1 and lMethodLineCount=true line num!!!
                                If Not ((lMethodStartLine = 1) And (GetMethodCodeLocation(oVBMethod) <> 1)) Then
                                    If (Not bDontPutErrorHandler) And (Not bEmptyMethod) Then
                                        'No error handlers for simple property procedures to improve performance and avoide resetting of errors in error handler code when it tries to access properties
                                        If Not (((evptVBMethodType = vbext_pk_Get) Or (evptVBMethodType = vbext_pk_Let) Or (evptVBMethodType = vbext_pk_Set)) And (lMethodActualLineCount = 1) And (Not bMethodMayContainCalls)) Then
                                            'Insert start
                                            Call oCodeModule.InsertLines(lMethodActualStartLineNumber + 1, vbCrLf & "    Const sMETHOD_NAME as String = " & sQUOTE & oVBProject.Name & "." & oVBComponent.Name & "." & sPropertyTypeString & oVBMethod.Name & sQUOTE & vbCrLf & "    On Error Goto ErrorTrap" & vbCrLf)
                                            
                                            'update last line number
                                            lMethodActualLastLineNumber = lMethodActualLastLineNumber + oCodeModule.ProcCountLines(oVBMethod.Name, evptVBMethodType) - lMethodLineCount
                                            
                                            'Parse the last line to get what the type of method
                                            Dim sVBMethodType As String
                                            Call MakeWordList(oCodeModule.Lines(lMethodActualLastLineNumber, 1), oWords)
                                            If oWords.Count = 2 Then 'Proper end line
                                                sVBMethodType = oWords(2)
                                                Dim sErrorHandlerCode As String
                                                sErrorHandlerCode = vbCrLf & "Exit " & sVBMethodType & vbCrLf & "ErrorTrap:" & vbCrLf
                                                If InStr(1, oVBMethod.Name, "_") = 0 Then
                                                    sErrorHandlerCode = sErrorHandlerCode & "    Call HandleError(sMETHOD_NAME, " & sQUOTE & oVBProject.Name & "." & oVBComponent.Name & sQUOTE & ")"
                                                Else
                                                    sErrorHandlerCode = sErrorHandlerCode & "    Call HandleError(sMETHOD_NAME, " & oVBProject.Name & "." & oVBComponent.Name & ", *Put here custom error code for event*)"
                                                End If
                                                Call oCodeModule.InsertLines(lMethodActualLastLineNumber, sErrorHandlerCode)
                                            End If
                                        End If
                                    End If
                                End If
 
NextPropertyType:
                                Select Case evptVBMethodType
                                    Case vbext_pk_Get
                                        evptVBMethodType = vbext_pk_Let
                                    Case vbext_pk_Let
                                        evptVBMethodType = vbext_pk_Set
                                    Case vbext_pk_Set
                                        evptVBMethodType = vbext_pk_Get
                                End Select
                                
                            Next lPropertyTypeIndex
                        End If
                    End If
                Next oVBMethod
            End If
ResumeNextComponent:
        Next oVBComponent
    Next oVBProject
    
    If sReadOnlyModuleNames <> vbNullString Then
        Err.Raise 1000, , "Line numbers can not be add to/removed from following modules because they are Read Only: " & vbCrLf & sReadOnlyModuleNames
    End If
    
Exit Sub
ErrorTrap:
    If Err.Number = 40198 Then  'Can not edit module
        sReadOnlyModuleNames = IIf(sReadOnlyModuleNames <> vbNullString, ", ", vbNullString) & sReadOnlyModuleNames & oVBComponent.Name
        Resume ResumeNextComponent
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetMethodCodeLocation(ByVal voVBMethod As Object) As Long
    
    Dim lGetMethodCodeLocation As Long
        
    On Error Resume Next
    lGetMethodCodeLocation = voVBMethod.CodeLocation
    
    If Err.Number = 424 Then
        GetMethodCodeLocation = 0
    ElseIf Err.Number = 0 Then
        GetMethodCodeLocation = lGetMethodCodeLocation
    Else
        Dim lErrorNumber As Long
        Dim sErrorDescription As String
        
        lErrorNumber = Err.Number
        sErrorDescription = Err.Source
        
        On Error GoTo ErrorTrap
        Err.Raise Err.Number, , Err.Description
    End If
    
Exit Function
ErrorTrap:
    ReRaiseError
End Function

Private Function IsOkToPutErrorHandler(ByVal voCodeModule As VBIDE.CodeModule, _
    ByVal voVBMethod As VBIDE.Member, ByVal veVBMemberType As vbext_ProcKind, _
    ByVal vlMethodStartScanLine As Long, ByVal vlMethodEndScaneLine As Long) As Boolean
    
    'Check if error handler already exists OR refrence to Err.Number or "'#NO_ERROR_HANDLER" is there
    Dim bDontPutErrorHandler As Boolean
    Dim bEmptyMethod As Boolean
    Dim bContinueChecks As Boolean
    Dim bMethodMayContainCalls As Boolean
    Dim bIsErrorHandlerForced As Boolean
    bDontPutErrorHandler = False
    bEmptyMethod = True
    bContinueChecks = True
    bMethodMayContainCalls = False
    bIsErrorHandlerForced = False
    
    Dim lMethodLineIndex As Long
    Dim lMethodCodeLineCount As Long
    lMethodCodeLineCount = 0
    For lMethodLineIndex = vlMethodStartScanLine To vlMethodEndScaneLine
        Dim sMethodLine As String
        sMethodLine = Trim(voCodeModule.Lines(lMethodLineIndex, 1))
        If UCase(sMethodLine) = "'#USE_ERROR_HANDLER" Then
            bDontPutErrorHandler = False
            bContinueChecks = False
            bIsErrorHandlerForced = True
        ElseIf (UCase(sMethodLine) = "'#NO_ERROR_HANDLER") Or (UCase(sMethodLine) = "REM *** FAILSAFE SKIP") Then
            bDontPutErrorHandler = True
            bContinueChecks = False
        End If
        
        If (sMethodLine <> vbNullString) And (Left(sMethodLine, 1) <> "'") Then
            bEmptyMethod = False
            lMethodCodeLineCount = lMethodCodeLineCount + 1
            If Not bMethodMayContainCalls Then
                If ((InStr(1, sMethodLine, "(") <> 0) And (InStr(1, sMethodLine, ")") <> 0)) Or (InStr(1, sMethodLine, "call ") <> 0) Or (InStr(1, sMethodLine, ", ") <> 0) Then
                    bMethodMayContainCalls = True
                End If
            End If
            
            If bContinueChecks Then
                'If method uses some error properties
                If (InStr(1, sMethodLine, "err.", vbTextCompare) <> 0) And (InStr(1, sMethodLine, "err.raise", vbTextCompare) = 0) Then
                    bDontPutErrorHandler = True
                   'Don't exit for loop. See next lines for commands.
                End If
                'See if error handler already exist
                Dim oFirstWords As New Collection
                Call MakeWordList(sMethodLine, oFirstWords)
                    'If first and second word is On Error
                    If oFirstWords.Count >= 2 Then
                        If (oFirstWords(1) = "On") And (oFirstWords(2) = "Error") Then
                            bDontPutErrorHandler = True
                        Else
                            If oFirstWords.Count >= 3 Then
                                'First is line number and then On Error
                                If (oFirstWords(2) = "On") And (oFirstWords(3) = "Error") Then
                                    bDontPutErrorHandler = True
                                End If
                            End If
                        End If
                    End If
                Set oFirstWords = Nothing
            End If
        End If
    Next lMethodLineIndex
    
    If (bDontPutErrorHandler = False) And (bIsErrorHandlerForced = False) Then
        If ((vlMethodStartScanLine = 1) And (GetMethodCodeLocation(voVBMethod) <> 1)) = True Then   'VB IDE object model bug
            bDontPutErrorHandler = True
        ElseIf bEmptyMethod = True Then
            bDontPutErrorHandler = True
        ElseIf (((veVBMemberType = vbext_pk_Get) Or (veVBMemberType = vbext_pk_Let) Or (veVBMemberType = vbext_pk_Set)) And (lMethodCodeLineCount = 1) And (Not bMethodMayContainCalls)) Then
            bDontPutErrorHandler = True
        End If
    End If
    
    IsOkToPutErrorHandler = Not bDontPutErrorHandler
End Function

Private Sub InsertErrorHandlerCode(ByVal voVBProject As VBIDE.VBProject, ByVal voVBComponent As VBIDE.VBComponent, _
    ByVal voCodeModule As VBIDE.CodeModule, ByVal voVBMethod As VBIDE.Member, _
    ByVal vlMethodStartLineNumberFromIDEModel As Long, _
    ByVal veVBMemberType As vbext_ProcKind, _
    ByVal vlMethodStartScanLine As Long, ByRef rlMethodEndScaneLine As Long)
    
    Dim sPropertyTypeString As String
    sPropertyTypeString = GetPropertyTypeNameString(voVBMethod, veVBMemberType)
    
    'Parse the last line to get what the type of method
    Dim sVBMethodType As String
    Dim oWords As New Collection
    Call MakeWordList(voCodeModule.Lines(rlMethodEndScaneLine + 1, 1), oWords)
    If oWords.Count = 2 Then 'Proper end line
        sVBMethodType = oWords(2)
        Dim sErrorHandlerCode As String
        sErrorHandlerCode = vbCrLf & "Exit " & sVBMethodType & vbCrLf & "WizAddInErrorTrap:" & vbCrLf
        Dim bIsPublicUserInterfaceOrEventMethod As Boolean
        bIsPublicUserInterfaceOrEventMethod = (InStr(1, voVBMethod.Name, "_") <> 0) And (voVBComponent.Type <> vbext_ct_ClassModule) And (voVBComponent.Type <> vbext_ct_StdModule)
            
        sErrorHandlerCode = sErrorHandlerCode & "    Call RecordError(" & sQUOTE & voVBComponent.Name & sQUOTE & _
            ", " & sQUOTE & sPropertyTypeString & voVBMethod.Name & sQUOTE & _
            ", " & "ewehLog" & IIf(bIsPublicUserInterfaceOrEventMethod, " Or ewehDisplay", " Or ewehReRaise") & _
            ")"
        Call voCodeModule.InsertLines(rlMethodEndScaneLine + 1, sErrorHandlerCode)
        
        'Insert this in after above code so the rlMethodEndScaneLine is not changed
        Call voCodeModule.InsertLines(vlMethodStartScanLine, vbCrLf & "    On Error Goto WizAddInErrorTrap" & vbCrLf)
        
        rlMethodEndScaneLine = rlMethodEndScaneLine + 8
    End If
End Sub

Private Function GetPropertyTypeNameString(ByVal voVBMethod As VBIDE.Member, ByVal veVBMethodType As vbext_ProcKind) As String
    Dim sPropertyTypeString As String
    If voVBMethod.Type = vbext_mt_Property Then
        Select Case veVBMethodType
            Case vbext_pk_Get
                sPropertyTypeString = "Get_"
            Case vbext_pk_Let
                sPropertyTypeString = "Let_"
            Case vbext_pk_Set
                sPropertyTypeString = "Set_"
            Case Else
                sPropertyTypeString = vbNullString
        End Select
    Else
        sPropertyTypeString = vbNullString
    End If
    GetPropertyTypeNameString = sPropertyTypeString
End Function

Private Sub GetMethodStartStopScanLineNumbers(ByVal voCodeModule As VBIDE.CodeModule, _
    ByVal voMethod As VBIDE.Member, _
    ByVal vblnAddOrRemove As Boolean, ByVal vblnAddErrorHandlers As Boolean, _
    ByRef rlMethodScanStartLine As Long, ByRef rlMethodScanStopLine As Long)
    
    Dim bIsVBProcedure As Boolean
    Dim eVBMemberType As vbext_MemberType
    Call GetVBCodeMemberType(voCodeModule, voMethod, bIsVBProcedure, eVBMemberType)
    
    If bIsVBProcedure = True Then
        Dim lMethodLineCount As Long
        Dim lMethodStartLineFromIDEModel As Long
        lMethodLineCount = voCodeModule.ProcCountLines(voMethod.Name, eVBMemberType)
        lMethodStartLineFromIDEModel = voCodeModule.ProcStartLine(voMethod.Name, eVBMemberType)
        
        'For adding line numbers or error handlers, go deep inside method where first line of code is
        'For removing line numbers, scane each of the line regardless of where it is
        If (vblnAddOrRemove = True) Or (vblnAddErrorHandlers = True) Then
            Dim sTrimmedLine As String
            'Find the last line of method
            'Look for the non blank/non comment line from the end of the procedure
            Dim lMethodActualLastLineNumber As Long
            For lMethodActualLastLineNumber = (lMethodStartLineFromIDEModel + lMethodLineCount - 1) To (lMethodStartLineFromIDEModel + 1) Step -1
                sTrimmedLine = Trim(voCodeModule.Lines(lMethodActualLastLineNumber, 1))
                'Ignore comments, directives and multiline method definations
                If (sTrimmedLine <> vbNullString) _
                    And (Left(sTrimmedLine, 1) <> "'") _
                    And (Right(sTrimmedLine, 1) <> "_") _
                    And (Left(sTrimmedLine, 1) <> "#") _
                    Then
                    Exit For
                End If
            Next lMethodActualLastLineNumber
            
            'Find the method start line
            Dim lMethodActualStartLineNumber As Long
            For lMethodActualStartLineNumber = lMethodStartLineFromIDEModel To lMethodActualLastLineNumber - 1
                sTrimmedLine = Trim(voCodeModule.Lines(lMethodActualStartLineNumber, 1))
                If (sTrimmedLine <> vbNullString) _
                    And (Left(sTrimmedLine, 1) <> "'") _
                    And (Right(sTrimmedLine, 1) <> "_") _
                    And (Left(sTrimmedLine, 1) <> "#") _
                    Then
                    Exit For
                End If
            Next lMethodActualStartLineNumber
            
            rlMethodScanStartLine = lMethodActualStartLineNumber + 1
            rlMethodScanStopLine = lMethodActualLastLineNumber - 1
        Else
            rlMethodScanStartLine = lMethodStartLineFromIDEModel
            rlMethodScanStopLine = lMethodStartLineFromIDEModel + lMethodLineCount - 1
        End If
    Else
        'This is some other code member otherthen method/property
    End If
End Sub

Private Sub GetVBCodeMemberType(ByVal voCodeModule As VBIDE.CodeModule, ByVal voVBMethod As VBIDE.Member, _
    ByRef rblnIsVBProcedure As Boolean, ByRef reVBMemberType As vbext_MemberType)
    
    If (voVBMethod.Type = vbext_mt_Method) Or (voVBMethod.Type = vbext_mt_Property) Then
        rblnIsVBProcedure = True
        
        'VB Bug workaround: Analyze method defination line to see if it's API declaration. In some API
        'Declare type of statement, it appears as method instead of declaration.
        Dim sMethodDefinationLine As String
        sMethodDefinationLine = GetMethodStartLine(voCodeModule, voVBMethod)
        Dim oWords As New Collection
        Call MakeWordList(sMethodDefinationLine, oWords)
        If voVBMethod.Type <> vbext_mt_Property Then
            reVBMemberType = vbext_pk_Proc
            If oWords.Count >= 2 Then
                If (LCase(oWords(1)) = "declare") Or (LCase(oWords(2)) = "declare") Then
                    rblnIsVBProcedure = False
                Else
                    'Method is not API declaration
                End If
            Else
                'Method is not API declaration
            End If
        Else
            'VB bug workaround: VB object model sometime errornously returns
            'other code members as property which they are not. Parse method
            'defination line to confirm.
            If oWords.Count >= 2 Then   'Check the second word
                Select Case LCase(oWords(2))
                    Case "get"
                        reVBMemberType = vbext_pk_Get
                    Case "let"
                        reVBMemberType = vbext_pk_Let
                    Case "set"
                        reVBMemberType = vbext_pk_Set
                    Case Else   'Check the 3rd word
                        If oWords.Count >= 3 Then
                            Select Case LCase(oWords(3))
                                Case "get"
                                    reVBMemberType = vbext_pk_Get
                                Case "let"
                                    reVBMemberType = vbext_pk_Let
                                Case "set"
                                    reVBMemberType = vbext_pk_Set
                                Case Else
                                    rblnIsVBProcedure = False
                            End Select
                        Else
                            rblnIsVBProcedure = False
                        End If
                End Select
            Else
                rblnIsVBProcedure = False
            End If
            Set oWords = Nothing
        End If
    Else
        rblnIsVBProcedure = False
    End If
End Sub

Private Sub AddRemoveLineNumbersInMethod(ByVal voCodeModule As VBIDE.CodeModule, ByVal voVBMethod As VBIDE.Member, _
    ByVal vblnAddOrRemoveLineNumbers As Boolean, ByVal vlMethodLineScaneIndexStart As Long, ByVal vlMethodLineScaneIndexStop As Long, _
    ByRef rblnIsSelectStatementStarted As Boolean, ByRef rblnIsSplittedLineStarted As Boolean)
    
    Dim bThisIsCaseStatement As Boolean
    bThisIsCaseStatement = False
    Dim lMethodLineIndex As Long
    For lMethodLineIndex = vlMethodLineScaneIndexStart To vlMethodLineScaneIndexStop
        Dim sMethodLine As String
        Dim lFirstWordStart As Long
        Dim lFirstWordLen As Long
        Dim sTrimmedLine As String
        
        sMethodLine = voCodeModule.Lines(lMethodLineIndex, 1)
        sTrimmedLine = Trim(sMethodLine)
        If (sTrimmedLine <> vbNullString) And (Not (Left(sTrimmedLine, 1)) = "'") And (Not (Left(sTrimmedLine, 1) = "#")) And (Not rblnIsSplittedLineStarted) Then
            
            'Check if line number already exist
            Dim sOriginalLineNumber As String
            sOriginalLineNumber = GetNextWord(sMethodLine, 1, , lFirstWordStart)
            lFirstWordLen = Len(sOriginalLineNumber)
            
            Dim bLineNumberExist As Boolean
            Dim bManualLineNumber As Boolean
            
            bLineNumberExist = False
            bManualLineNumber = False
            If lFirstWordLen <> 0 Then
                If (Left(sOriginalLineNumber, 1) <> "'") Then   'If not comment
                    If (Right(sOriginalLineNumber, 1) = ":") Then 'Remove last colon
                        If lFirstWordLen <> 1 Then
                            sOriginalLineNumber = Left(sOriginalLineNumber, lFirstWordLen - 1)
                            bManualLineNumber = True
                        End If
                    End If
                    If IsNumeric(sOriginalLineNumber) Then
                        bLineNumberExist = True
                    End If
                End If
            End If
            
            If InStr(1, sTrimmedLine, "Select Case", vbTextCompare) = 1 Then
                rblnIsSelectStatementStarted = True
            ElseIf rblnIsSelectStatementStarted = True Then
                If InStr(1, sTrimmedLine, "End Select", vbTextCompare) = 1 Then
                    rblnIsSelectStatementStarted = False
                End If
            End If
            
            If rblnIsSelectStatementStarted = True Then
                If InStr(1, sTrimmedLine, "Case", vbTextCompare) = 1 Then
                    bThisIsCaseStatement = True
                Else
                    bThisIsCaseStatement = False
                End If
            Else
                bThisIsCaseStatement = False
            End If
            
            If vblnAddOrRemoveLineNumbers = True Then 'Add
                If bLineNumberExist And (Not bManualLineNumber) And (Not bThisIsCaseStatement) Then    'Remove prev ones
                    sMethodLine = Mid(sMethodLine, lFirstWordStart + lFirstWordLen + 1)
                End If
                
                If ((Not bLineNumberExist) Or (bLineNumberExist And (Not bManualLineNumber))) And (Not bThisIsCaseStatement) Then
                    Call voCodeModule.ReplaceLine(lMethodLineIndex, lMethodLineIndex & " " & sMethodLine)
                End If
            Else                    'Remove
                If bLineNumberExist And (Not bManualLineNumber) And (Not bThisIsCaseStatement) Then
                    Call voCodeModule.ReplaceLine(lMethodLineIndex, Mid(sMethodLine, lFirstWordStart + lFirstWordLen + 1))
                End If
            End If
            
        End If

        If Right(sTrimmedLine, 1) = "_" Then
            rblnIsSplittedLineStarted = True
        Else
            rblnIsSplittedLineStarted = False
        End If
    Next lMethodLineIndex
End Sub
