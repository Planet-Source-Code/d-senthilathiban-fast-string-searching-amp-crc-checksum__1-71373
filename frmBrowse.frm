VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Browse File"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConditions 
      Caption         =   "Select File to Search for the String and CRC32"
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   8160
         TabIndex        =   7
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   495
         Left            =   7440
         TabIndex        =   6
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1200
         TabIndex        =   4
         Top             =   4680
         Width           =   6135
      End
      Begin VB.FileListBox File1 
         Height          =   4185
         Left            =   5040
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.DirListBox Dir1 
         Height          =   3690
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   4695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Selected file"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   4800
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DrivePathFile As String

Private Sub cmdCancel_Click()
DrivePathFile = ""
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()

On Error GoTo Err_Line

Dir1.Path = Drive1.Drive
Exit Sub

Err_Line:
    
    If Err Then
        MsgBox Err.Description, vbCritical, Me.Caption
        Dir1.Path = App.Path
        Drive1.Drive = Dir1.Path
    End If

End Sub

Private Sub File1_Click()

If Right(Dir1.Path, 1) <> "\" Then
    Text1 = Dir1.Path & "\" & File1.FileName
Else
    Text1 = Dir1.Path & File1.FileName
End If

DrivePathFile = Text1
End Sub

Private Sub Form_Load()

Dir1.Path = App.Path

End Sub

Private Sub cmdOK_Click()

On Error GoTo NotFound

        If FileLen(DrivePathFile) = 0 Then
        'If Dir(DrivePathFile) = "" Then
            
            MsgBox "Empty File not allowed.", vbExclamation, Me.Caption
            Dir1.Path = App.Path
            Drive1.Drive = Dir1.Path
            Exit Sub
        
        End If
        
frmFileSearchCRC.txtPathFile = DrivePathFile
Unload Me
Exit Sub

NotFound:

MsgBox "File Not Found", vbCritical, "Error"
Unload Me

End Sub


