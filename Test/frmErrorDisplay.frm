VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmErrorDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Error"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmErrorDisplay.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlgSave 
      Left            =   3060
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "doc"
      DialogTitle     =   "Save Error Report"
      FileName        =   "*.doc"
      Filter          =   "*.doc|Word Files|*.*|All Files"
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   9
      Top             =   1890
      Width           =   9195
   End
   Begin VB.Frame fraFormDownLimit 
      Height          =   60
      Left            =   -180
      TabIndex        =   7
      Top             =   5085
      Visible         =   0   'False
      Width           =   9195
   End
   Begin VB.Frame fraFormNoDetailLowerLimit 
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   2610
      Width           =   9195
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "More &Details >>"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   2070
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   5085
      TabIndex        =   3
      Top             =   2070
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateErrorReport 
      Caption         =   "&Create Error Report..."
      Height          =   420
      Left            =   3150
      TabIndex        =   2
      Top             =   2070
      Width           =   1815
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "frmErrorDisplay.frx":0442
      Height          =   1905
      Left            =   90
      TabIndex        =   0
      Top             =   3060
      Width           =   6405
      _cx             =   11298
      _cy             =   3360
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   1
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   4950
      Top             =   -540
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   90
      Picture         =   "frmErrorDisplay.frx":0457
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Error Trace And Line Numbers:"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   2790
      Width           =   2190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "If the error is unknow, please report this to our support team to help us find the solution as soon as possible."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   1710
      TabIndex        =   8
      Top             =   495
      Width           =   4770
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblErrorMessage 
      Height          =   420
      Left            =   1710
      TabIndex        =   5
      Top             =   1350
      Width           =   4740
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "An error has occured in your application."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   1710
      TabIndex        =   1
      Top             =   90
      Width           =   3795
   End
End
Attribute VB_Name = "frmErrorDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbIsDetailShown As Boolean
Private Const MAX_PATH = 260
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private moErrorTrace As ADODB.Recordset

Public Sub ShowError(ByVal voErrorTrace As ADODB.Recordset)
    '#NO_ERROR_HANDLER
    Dim sErrorInformation As String
    sErrorInformation = "Basic Error information: " & vbCrLf & "Error Number: " & Err.Number & " - " & Err.Description _
    & " [Source: " & Err.Source & " - Line: " & Erl & "]"
    lblErrorMessage.Caption = sErrorInformation
    
    Set Adodc1.Recordset = voErrorTrace
    Call VSFlexGrid1.DataRefresh
    mbIsDetailShown = True:    Call cmdDetails_Click
    Set moErrorTrace = voErrorTrace
    Me.Show vbModal
    Set moErrorTrace = Nothing
    Unload Me
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo WizAddInErrorTrap

    Me.Hide

Exit Sub
WizAddInErrorTrap:
    Call RecordError("frmErrorDisplay", "cmdCancel_Click", ewehLog Or ewehDisplay)
End Sub

Private Sub cmdCreateErrorReport_Click()

    On Error GoTo WizAddInErrorTrap


    Dim sTempDir As String
    sTempDir = GetPathWithSlash(GetTempDirPath)
    
    'Take desktop screenshot
    Call DeleteFileSilent(sTempDir & "wizaddin_error_screenshot.bmp")
    Call fSaveGuiToFile(sTempDir & "wizaddin_error_screenshot.bmp")
    
    Call DeleteFileSilent(sTempDir & "wizaddin_error_log.xml")
    Call moErrorTrace.Save(sTempDir & "wizaddin_error_log.xml", adPersistXML)
    
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    
    Set oWordApp = New Word.Application
    oWordApp.Visible = True
    oWordApp.Activate
    Set oWordDoc = oWordApp.Documents.Add
    Call oWordDoc.Activate
    Call oWordDoc.Range.SetRange(0, 0)
    oWordDoc.Range.Style = oWordDoc.Styles("Heading 2")
    oWordDoc.Range.InsertAfter "Error Report:" & vbCrLf
    oWordDoc.Range.Style = oWordDoc.Styles("Normal")
    oWordDoc.Range.InsertAfter "User: Xyz" & vbCrLf
    oWordDoc.Range.InsertAfter "Date Created: " & Format(Now, "dd mmm yyyy Hh:Mm") & vbCrLf & vbCrLf
    oWordDoc.Range.InsertAfter "Please type here any additional information which can help us to solve this problem." & vbCrLf
    
    oWordDoc.Range.InsertParagraphAfter
    oWordDoc.Range.Style = oWordDoc.Styles("Heading 3")
    oWordDoc.Range.InsertAfter Text:="Support Documents:" & vbCrLf
    Call oWordDoc.Range.MoveEnd(wdParagraph, 1)
    oWordDoc.Range.InlineShapes.AddOLEObject ClassType:="Package", FileName:= _
        sTempDir & "wizaddin_error_screenshot.bmp", LinkToFile:=False, DisplayAsIcon:=True, IconLabel:="Error Screenshot", Range:=oWordDoc.Range
    oWordDoc.Range.InlineShapes.AddOLEObject ClassType:="Package", FileName:= _
        sTempDir & "wizaddin_error_log.xml", LinkToFile:=False, DisplayAsIcon:=True, IconLabel:="Error log", Range:=oWordDoc.Range
    
    Set oWordDoc = Nothing
    Set oWordApp = Nothing

Exit Sub
WizAddInErrorTrap:
    Call RecordError("frmErrorDisplay", "cmdCreateErrorReport_Click", ewehLog Or ewehDisplay)
End Sub

Private Sub cmdDetails_Click()
    '#NO_ERROR_HANDLER
    mbIsDetailShown = Not mbIsDetailShown
    If mbIsDetailShown = False Then
        Me.Height = 2985
        cmdDetails.Caption = "More &Details >>"
    Else
        Me.Height = 5475
        cmdDetails.Caption = "<< Less &Details"
    End If
End Sub


Public Sub DeleteFileSilent(ByVal vsFileName As String)
    On Error Resume Next
    Kill vsFileName
End Sub

Private Sub fSaveGuiToFile(ByVal theFile As String)

    '#NO_ERROR_HANDLER

    ' Name: fSaveGuiToFile
    ' Author: Dalin Nie
    ' Written: 4/2/99
    ' Purpose:
    ' This procedure will Capture the Screen
    '     or the active window of your Computer an
    '     d Save it as
    ' a .bmp file
    ' Input:
    ' theFile file Name with path, where you
    '     want the .bmp to be saved
    '
    ' Output:
    ' True if successful
    '
    Dim lString As String
    
    'Check if the File Exist
    If Dir(theFile) <> "" Then Exit Sub
    'To get the Entire Screen
    Call keybd_event(vbKeySnapshot, 0, 0, 0)
    'To get the Active Window
    'Call keybd_event(vbKeySnapshot, 0, 0, 0
    '     )
    DoEvents
    SavePicture Clipboard.GetData(vbCFBitmap), theFile
End Sub


Public Function GetTempDirPath() As String

    '#NO_ERROR_HANDLER

    Dim sBuffer As String
    Dim lTempDirLen As Long
    
    sBuffer = String$(MAX_PATH, 0)
    lTempDirLen = GetTempPath(MAX_PATH, sBuffer)
    If lTempDirLen > 1 Then
        GetTempDirPath = RemoveSlashAtEnd(Left$(sBuffer, lTempDirLen - 1))
    Else
        GetTempDirPath = RemoveSlashAtEnd(App.Path)
    End If
End Function


'----------------------------------------------------------------------------
'Method:            GetPathWithSlash
'Author/Date:       Shital/Jul 30, 2001
'Description:       Check path if it ends with slash. If not, add a slash.
'Input:             sPath: Path to be checked
'Output:            None
'Modifications:
'----------------------------------------------------------------------------

Public Function GetPathWithSlash(ByVal sPath As String) As String


    '#NO_ERROR_HANDLER


    Dim sPathWithSlash As String
    If sPath <> vbNullString Then
        If Right(sPath, 1) = "\" Then
            sPathWithSlash = sPath
        Else
            sPathWithSlash = sPath & "\"
        End If
    Else
        sPathWithSlash = vbNullString
    End If
    GetPathWithSlash = sPathWithSlash
End Function


Private Sub Form_Unload(Cancel As Integer)
    '#NO_ERROR_HANDLER
    Set moErrorTrace = Nothing
End Sub

Public Function RemoveSlashAtEnd(ByVal vsPath As String, Optional ByVal vsSlashChar As String = "\") As String
    '#NO_ERROR_HANDLER
    Dim lPathLen As Long
    
    lPathLen = Len(vsPath)
    
    If Right$(vsPath, 1) <> vsSlashChar Then
        RemoveSlashAtEnd = vsPath
    Else
        If lPathLen > 1 Then
            RemoveSlashAtEnd = Left(vsPath, lPathLen - 1)
        Else
            RemoveSlashAtEnd = vbNullString
        End If
    End If

End Function
