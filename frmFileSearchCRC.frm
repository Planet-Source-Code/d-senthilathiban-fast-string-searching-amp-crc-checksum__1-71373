VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileSearchCRC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File String Searching & CRC32 Checksum Calculation Testing"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Methods To Be Benchmarked"
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   3135
      Begin VB.CommandButton cmdCRC 
         Caption         =   "Get CRC"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "Whole File Buffering"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "MMF concept"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "Others1"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "Others2"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "4KB File Buffering"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.Frame fraMethods 
      Caption         =   "Methods To Be Benchmarked"
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Asm_ C language StrStr String searching"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Visual basic Instr function With MMF"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "VB_Boyer-Moore horspool string searching"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Asm_Byte by byte string searching"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Asm_Boyer-Moore horspool string searching"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   4335
      End
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   1680
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Method"
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Run Time (µs)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size (Bytes)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CRC/Loc"
         Object.Width           =   2028
      EndProperty
   End
   Begin VB.Frame fraConditions 
      Caption         =   "File and Search String"
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6840
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtPathFile 
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "String to be Searched:"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Selected file"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   3720
      ScaleHeight     =   75
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   6000
      Width           =   3375
   End
End
Attribute VB_Name = "frmFileSearchCRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The test form Written by Chris Lucas for String Concatenation Methods
' is used with modification for time saving.
' Thanks to Chris Lucas

Private cCRCSearch As cFileSearchCRC
Private Timer As cPrecisionTimer

'Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdBrowse_Click()
frmBrowse.Show

End Sub

Private Sub cmdCRC_Click()
    Dim lngCount As Long
    Dim strBuffer As String
    Dim lngResult As Long
    Dim itmListItem As ListItem
    
    
    If Len(Me.txtPathFile) = 0 Then
    MsgBox "File name not given or empty file selected"
    Exit Sub
    End If
    
    lvwResults.ListItems.Clear
    
    Screen.MousePointer = vbHourglass
    
    ' Whole File Buffering
    If (chkCRC(0) = vbChecked) Then
        
        Timer.ResetTimer
        'StartTime1 = GetTickCount
        'StartTime2 = timeGetTime
        strBuffer = Hex(cCRCSearch.CalculateFile(Me.txtPathFile))
        strBuffer = Right("00000000" & strBuffer, 8)
        Timer.StopTimer
        lngResult = Timer.Elapsed
        'EndTime2 = timeGetTime
        'EndTime1 = GetTickCount
        'lngResult = EndTime1 - StartTime1
        'lngResult = EndTime2 - StartTime2
        
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Whole File Buffering"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
    End If
    
    ' 4KB File Buffering
    If (chkCRC(1) = vbChecked) Then
        Timer.ResetTimer
        strBuffer = cCRCSearch.CalculateFileCRC(Me.txtPathFile)
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "4KB File Buffering"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        
        ' Clean up

    End If
    
    ' Standard Buffering Concatenation
    If (chkCRC(2) = vbChecked) Then
        Timer.ResetTimer
        strBuffer = cCRCSearch.FileMapCRC(Me.txtPathFile)
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "MMF concept"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")

        ' Clean up
        'Concat.ReInit
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    'Dim lngOutPuts(0 To 5, 1 To 26) As Long
    'Dim lngCount As Long
    'Dim lngStart As Long
    'Dim strTest As String
    Dim strBuffer As String
    Dim lngResult As Long
    Dim itmListItem As ListItem
    
    If Len(Me.txtPathFile) = 0 Then
    MsgBox "File name not given or empty file selected"
    Exit Sub
    End If
    
    'strTest = "BxgÕŒ3Ö~ÄxçíÈ¿µC„Ø#”öÎ1"
    'È¿µC„Ø#”öÎ1
    'BxgÕŒ3Ö~ÄxçíÈ¿µC„Ø#”öÎ1
    
    'lngStart = 15
    'For lngCount = 3 To Len(strTest)
    
    lvwResults.ListItems.Clear
    
    Screen.MousePointer = vbHourglass
    'If lngCount = 13 Then lngStart = 1
    'txtSearch.Text = Mid$(strTest, lngStart, lngCount)
    ' Asm_Boyer-Moore horspool string searching
    If (chkMethods(0) = vbChecked) Then
        Timer.ResetTimer
        cCRCSearch.SearchAlgorithm = Asm_BMHA
        strBuffer = Str$(cCRCSearch.FileMapSearch(Me.txtPathFile, txtSearch.Text))
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Asm_Boyer-Moore horspool string search"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        'lngOutPuts(1, Len(txtSearch.Text)) = lngResult
    End If
    
    ' Asm_C language StrStr String searching
    If (chkMethods(1) = vbChecked) Then
        Timer.ResetTimer
        cCRCSearch.SearchAlgorithm = Asm_STRC
        strBuffer = Str$(cCRCSearch.FileMapSearch(Me.txtPathFile, txtSearch.Text))
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Asm_C language StrStr String search"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        'lngOutPuts(2, Len(txtSearch.Text)) = lngResult
    End If
    
    ' Asm_Byte by byte string searching
    If (chkMethods(2) = vbChecked) Then
        Timer.ResetTimer
        cCRCSearch.SearchAlgorithm = Asm_BYTE
        strBuffer = Str$(cCRCSearch.FileMapSearch(Me.txtPathFile, txtSearch.Text))
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Asm_Byte by byte string search"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        'lngOutPuts(3, Len(txtSearch.Text)) = lngResult
    End If
    
    ' VB_Boyer-Moore horspool string searching
    If (chkMethods(3) = vbChecked) Then
        Timer.ResetTimer
        cCRCSearch.SearchAlgorithm = Vb_BMHA
        strBuffer = Str$(cCRCSearch.FileMapSearch(Me.txtPathFile, txtSearch.Text))
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "VB_Boyer-Moore horspool string search"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        'lngOutPuts(4, Len(txtSearch.Text)) = lngResult
    End If
    
    ' Visual basic Instr function With MMF
    If (chkMethods(4) = vbChecked) Then
        Timer.ResetTimer
        cCRCSearch.SearchAlgorithm = Vb_InStr
        strBuffer = Str$(cCRCSearch.FileMapSearch(Me.txtPathFile, txtSearch.Text))
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Visual basic Instr function With MMF"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(FileLen(Me.txtPathFile), "#,###")
        itmListItem.SubItems(3) = strBuffer 'Format$(Len(strBuffer), "#,###")
        'lngOutPuts(5, Len(txtSearch.Text)) = lngResult
    End If
    'lngOutPuts(0, lngCount) = Val(strBuffer)
    Screen.MousePointer = vbDefault
    'Next
    
    'Open App.Path & "\out.txt" For Output Access Write As #1
    'For lngCount = 3 To Len(strTest)
    'Print #1, lngCount; lngOutPuts(0, lngCount); lngOutPuts(1, lngCount); lngOutPuts(2, lngCount); lngOutPuts(3, lngCount); lngOutPuts(4, lngCount); lngOutPuts(5, lngCount)
    'Next
    'Close #1
    
End Sub

Private Sub Form_Load()
    Set Timer = New cPrecisionTimer
    Set cCRCSearch = New cFileSearchCRC
    
    ' Beautify the listview
    lvwResults.FullRowSelect = True
    SetListViewLedger lvwResults, vbLedgerLightBlue, vbLedgerYellow, sizeNone
    'display memory status
    'cCRCSearch.GetMemStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Timer = Nothing
    Set cCRCSearch = Nothing
End Sub

