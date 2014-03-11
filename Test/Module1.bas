Attribute VB_Name = "Module1"
Option Explicit

Public Sub f1()

    On Error GoTo WizAddInErrorTrap

    Dim x
    x = 2
    x = 2 / x
    Err.Raise Err.Number

Exit Sub
WizAddInErrorTrap:
    Call RecordError("Module1", "f1", ewehLog Or ewehReRaise)
End Sub

Public Sub f2()
    '#NO_ERROR_HANDLER
    
    Dim x
    x = 2
    x = 2 / x
End Sub

'dodo
Public Sub f3()
    Rem *** FailSafe SKIP
    
    Dim x
    x = 2
    x = 2 / x
End Sub

Public Sub f4()

    On Error GoTo WizAddInErrorTrap

    Dim x
    x = 2
    x = 2 / x

Exit Sub
WizAddInErrorTrap:
    Call RecordError("Module1", "f4", ewehLog Or ewehReRaise)
End Sub

'lara
Public Sub DoSomeThing()

    On Error GoTo WizAddInErrorTrap

    Dim x
    Dim y
    x = 0
    y = 0
    Dim z
    z = x / y

Exit Sub
WizAddInErrorTrap:
    Call RecordError("Module1", "DoSomeThing", ewehLog Or ewehReRaise)
End Sub
