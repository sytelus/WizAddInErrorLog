VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


#Const debugging = True

Public Sub s1()

    On Error GoTo WizAddInErrorTrap


    Dim a As String
    
    MsgBox "kl"


Exit Sub
WizAddInErrorTrap:
    Call RecordError("Class1", "s1", ewehLog Or ewehReRaise)
End Sub

Public Sub DoSomeThing()

    On Error GoTo WizAddInErrorTrap


    Dim o2 As New Class2
    o2.DoSomeThing


Exit Sub
WizAddInErrorTrap:
    Call RecordError("Class1", "DoSomeThing", ewehLog Or ewehReRaise)
End Sub
