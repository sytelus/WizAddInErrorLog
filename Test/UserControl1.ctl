VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton Command1 
      Caption         =   "Do Something in user control!"
      Height          =   1230
      Left            =   675
      TabIndex        =   0
      Top             =   720
      Width           =   2445
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Property Get p1()

    On Error GoTo WizAddInErrorTrap


    Dim x
    'hi
    MsgBox "1 +1"
    Dim z
    
    'jojo
101:
    Call f1: Call f2


Exit Property
WizAddInErrorTrap:
    Call RecordError("UserControl1", "Get_p1", ewehLog Or ewehReRaise)
End Property

Sub f1()

End Sub

Sub f2()

End Sub

Private Sub Command1_Click()

    On Error GoTo WizAddInErrorTrap


    Dim o1 As New Class1
    o1.DoSomeThing


Exit Sub
WizAddInErrorTrap:
    Call RecordError("UserControl1", "Command1_Click", ewehLog Or ewehDisplay)
End Sub
