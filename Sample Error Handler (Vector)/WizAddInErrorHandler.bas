Attribute VB_Name = "modWizAddinErrorHandler"
Option Explicit


Public Enum enmWizAddInErrorHandlingOptions
    ewehReRaise = 1
    ewehDisplay = 2
    ewehLog = 4
    ewehSaveErrorInfo = 8
    ewehRestoreErrorInfo = 16
    'NOTE: New error handling option value must be in sequence: 1,2,4,8,16,32,...
End Enum

Public Sub RecordError(ByVal ModuleName As String, ByVal MethodName As String, _
                    ErrorHandlingOptions As enmWizAddInErrorHandlingOptions, _
                    Optional ByVal ErrorNumberOveride As Long = 0, Optional ByVal ErrorSourceOveride As String = vbNullString, _
                    Optional ByVal ErrorDescriptionOveride As String = vbNullString)

    'Following directive tells add-in not to touch this function
    '#NO_ERROR_HANDLER

    'If override parameters are provided, use them
    Dim lErrorNumber As Long
    Dim sErrorSource As String
    Dim sErrorDescription As String
    lErrorNumber = IIf(ErrorNumberOveride = 0, Err.Number, ErrorNumberOveride)
    sErrorSource = IIf(ErrorSourceOveride = vbNullString, Err.Source, ErrorSourceOveride)
    sErrorDescription = IIf(ErrorDescriptionOveride = vbNullString, Err.Description, ErrorDescriptionOveride)

    'If we need to save error info, it will be saved in these static variables
    Static lSavedErrorNumber As Long
    Static sSavedErrorSource As String
    Static sSavedErrorDescription As String
    If (ErrorHandlingOptions And ewehSaveErrorInfo) <> 0 Then
        lSavedErrorNumber = lErrorNumber
        sSavedErrorSource = sErrorSource
        sSavedErrorDescription = sErrorDescription
    End If
    
    'if we need to restore error info, restore it back from static variables
    If (ErrorHandlingOptions And ewehRestoreErrorInfo) <> 0 Then
        lErrorNumber = lSavedErrorNumber
        sErrorSource = sSavedErrorSource
        sErrorDescription = sSavedErrorDescription
    End If
    
    'Log the error
    If (ErrorHandlingOptions And ewehLog) <> 0 Then
        'NOTE: anything that writes to log storage can be used here instead of KLog object
        If Not g_objKLog Is Nothing Then
            ' Report error
            Call g_objKLog.Report(lErrorNumber, App.EXEName, ModuleName, MethodName, Erl, sErrorDescription)
        Else
            ' Don't report
        End If
    End If
    
    'Display the error
    If (ErrorHandlingOptions And ewehDisplay) <> 0 Then
        'Bypass KLog dialog as it hides error
        Call MsgBox("Following error was occured during the operation. Error " & lErrorNumber & _
            " : " & Err.Description & vbCrLf & "Source: " & Err.Source & IIf(Erl <> 0, " Line: " & Erl, vbNullString), vbInformation Or vbOKOnly, "Vector Error")
    End If
    
    'Bubble back the error in caller chain
    If (ErrorHandlingOptions And ewehReRaise) <> 0 Then
        'The Err.Raise 0 will produce Invalid Procedure call error, so avoide that
        If lErrorNumber <> 0 Then
            Err.Raise lErrorNumber, sErrorSource, sErrorDescription
        Else
            Err.Raise -1, sErrorSource, "(original err# is 0) " & sErrorDescription
        End If
    End If
    
End Sub



