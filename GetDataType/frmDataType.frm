VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataType 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmDataType"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9340
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File"
         Object.Tag             =   "File"
         Text            =   "File"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Type"
         Object.Tag             =   "Type"
         Text            =   "Type"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abort"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   420
      Hidden          =   -1  'True
      Left            =   4560
      System          =   -1  'True
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Path: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GetDataType
'Determines the datatype in any given file, regardless of its extension.
'
'Design & Coding by Ben Klaster
'(C)Copyright OmegaWare.nl 2003-2004

Option Explicit

Private Const FDT_FILE_INVALID = -2
Private Const FDT_FILE_EMPTY = -1
Private Const FDT_FILE_TEXT = 0
Private Const FDT_FILE_BINARY = 1
Private Const FDT_FILE_UNICODE = 2

Private sPath As String
Private bAborted As Boolean

Private Sub Form_Load()
    sPath = "c:\"
    
    File1.Path = sPath
    Label1.Caption = "Path: " & UCase(sPath)
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = True
    
    bAborted = False
    GetDataTypes
    Label2.Caption = ""

    Command2.Enabled = False
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    bAborted = True
End Sub

Private Sub GetDataTypes()
    Dim lCounter As Long, lReturn As Long, lvw1 As Variant
        
    ListView1.ListItems.Clear
    
    For lCounter = 0 To File1.ListCount - 1
        Set lvw1 = ListView1.ListItems.Add
            lvw1.Text = UCase(File1.List(lCounter))
        
        Label2.Caption = "File: " & lvw1.Text
        
        lReturn = GetDataType(sPath & "\" & File1.List(lCounter))
            If bAborted = True Then Exit For
            
            If lReturn = FDT_FILE_INVALID Then
                lvw1.SubItems(1) = "Error Opening"
            End If
            
            If lReturn = FDT_FILE_EMPTY Then
                lvw1.SubItems(1) = "Empty file"
            End If
            
            If lReturn = FDT_FILE_TEXT Then
                lvw1.SubItems(1) = "Text"
            End If
            
            If lReturn = FDT_FILE_UNICODE Then
                lvw1.SubItems(1) = "UniCode Text"
            End If
            
            If lReturn = FDT_FILE_BINARY Then
                lvw1.SubItems(1) = "Binary"
            End If
    Next lCounter
End Sub

Private Function GetDataType(ByVal lpstrFile As String) As Long
    Dim lCounter As Long, lDataSize As Long
    Dim lFileHandle As Long, bByteArray() As Byte
    
    On Error Resume Next
    
    lFileHandle = FreeFile
    
    Open lpstrFile For Binary Access Read As lFileHandle
        If Err > 0 Then
            GetDataType = FDT_FILE_INVALID
            Exit Function
        End If
        
        lDataSize = LOF(lFileHandle) - 1
            If lDataSize < 0 Then
                GetDataType = FDT_FILE_EMPTY
                Close lFileHandle
                Exit Function
            End If
            
        ReDim bByteArray(0 To lDataSize)
        Get #lFileHandle, , bByteArray()
    Close lFileHandle
    
    For lCounter = 0 To lDataSize
        If bAborted = True Then Exit Function
                
        If (lDataSize Mod lCounter) = 0 Then DoEvents
        
        If bByteArray(lCounter) = 5 Or _
           bByteArray(lCounter) = 6 Or _
           bByteArray(lCounter) = 11 Or _
           bByteArray(lCounter) = 14 Or _
           bByteArray(lCounter) = 15 Or _
           bByteArray(lCounter) = 18 Or _
           bByteArray(lCounter) = 19 Or _
           bByteArray(lCounter) = 20 Or _
           bByteArray(lCounter) = 21 Or _
           bByteArray(lCounter) = 23 Or _
           bByteArray(lCounter) = 24 Or _
           bByteArray(lCounter) = 25 Or _
           bByteArray(lCounter) = 29 Then
            GetDataType = FDT_FILE_BINARY
            Exit Function
        End If
    Next lCounter
    
    If (bByteArray(0) = 255) And (bByteArray(1) = 254) Then
        GetDataType = FDT_FILE_UNICODE
    Else
        GetDataType = FDT_FILE_TEXT
    End If
End Function

