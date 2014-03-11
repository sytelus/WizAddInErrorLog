VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   2580
      Left            =   360
      TabIndex        =   0
      Top             =   225
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4551
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a

#If db Then
    w = 3
#End If

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()

    On Error GoTo WizAddInErrorTrap


    wo = 5
    
    #If dfb Then    'woh
        'comma
       d = 3
    #End If


Exit Sub
WizAddInErrorTrap:
    Call RecordError("Form1", "Form_Load", ewehLog Or ewehDisplay)
End Sub
