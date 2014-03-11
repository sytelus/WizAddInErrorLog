Attribute VB_Name = "Module1"
Option Explicit

'These options could be combined so it must be set with values like 0,1,2,4,8,16,32,64,128
Public Enum enmErrorHandlingOption
    eehoNone = 0
    eehoReRaise = 1
    eehoDisplay = 2
    eehoLog = 4 'These options could be combined so it must be set with values like 0,1,2,4,8,16,32,64,128
End Enum

'Re raise the error back to the caller
Private Sub ReRaiseError()
    With Err
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With
End Sub

'In most functions, we put error handler which can trap the error, log the error and raise it back to the caller. The purpose of doing this is to produce a call trace at run time which can be analysed by support team to find the cause of the problem. This eliminates the usually costly process of duplicating the problem on a development machine with source code.
'Parameters:
'ModuleName - In which module error occurred?
'MethodName - In which method error occurred?
'HandlingOptions - Enumeration values which can be combined togather.
'    eehoLog - Log error
'    eehoReRaise - Reraise error
'    eehoDisplay - Display error
'    eehoNone - Do nothing
'The handling options can be combines like this:
'eehoLog Or eehoReRaise - Log the error and raise it again so the error propogates to the caller.
'eehoLog Or eehoDisplay - Log the error and display it. Do not raise it again.
'
'There are 3 override parameters which if specified will be used instead of value in Err object.
Public Sub RecordAndReRaiseError(ByVal ModuleName As String, ByVal MethodName As String, _
                    Optional ByVal ErrorHandlingOptions As enmErrorHandlingOption = eehoLog Or eehoReRaise, _
                    Optional ByVal ErrorNumberOveride As Long = 0, Optional ByVal ErrorSourceOveride As String = vbNullString, _
                    Optional ByVal ErrorDescriptionOveride As String = vbNullString)
    
    'Decide whether to use overides or values in Err object
    Dim lErrorNumber As Long
    Dim sErrorSource As String
    Dim sErrorDescription As String
    
    lErrorNumber = IIf(ErrorNumberOveride = 0, Err.Number, ErrorNumberOveride)
    sErrorSource = IIf(ErrorSourceOveride = vbNullString, Err.Source, ErrorSourceOveride)
    sErrorDescription = IIf(ErrorDescriptionOveride = vbNullString, Err.Description, ErrorDescriptionOveride)

    'Should we log the error?
    If (ErrorHandlingOptions And eehoLog) <> 0 Then
        If Not g_objKLog Is Nothing Then
            ' Report error
            Call g_objKLog.Report(lErrorNumber, App.EXEName, ModuleName, MethodName, Erl, "[" & sErrorSource & "] - " & sErrorDescription)
        Else
            ' Don't report
        End If
    End If
    
    'Should we display the error?
    If (ErrorHandlingOptions And eehoDisplay) <> 0 Then
        'Bypass KLog dialog as it hides error
        Call MsgBox("Following error was occured during the operation. Error " & lErrorNumber & _
            " : " & Err.Description & vbCrLf & "Source: " & Err.Source & IIf(Erl <> 0, " Line: " & Erl, vbNullString), vbInformation Or vbOKOnly, "Vector Error")
    End If
    
    'Should we re-raise the error
    If (ErrorHandlingOptions And eehoLog) <> 0 Then
        'bubble up the error
        ReRaiseError
    End If
    
End Sub


