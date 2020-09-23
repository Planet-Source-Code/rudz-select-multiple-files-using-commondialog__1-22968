VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Multi FileSelect w. CommonDialog"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Files"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form1.frx":000C
      Left            =   120
      List            =   "Form1.frx":000E
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2520
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    With dlg

    ' Set Flags
    .Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNLongNames

    ' Max Size
    .MaxFileSize = 20000

    ' Reset FileName
    .FileName = ""

    ' Show the Open Dialog
    .ShowOpen

    ' Check to see if the user selected a file, if not - exit
    If .FileName = "" Then Exit Sub

    ' Counter var
    Dim i As Integer

    ' Go through all files selected
    For i = 1 To CountFilesInList(.FileName)

        ' Check the file size
        Select Case FileLen(GetFileFromList(.FileName, i))
            Case Is > 0
                ' Now add the file to the list boxes
                List1.AddItem GetFileFromList(.FileName, i)
            Case Else
                ' If filelenght is 0 (zero) - ask if it should be added anyway
                If MsgBox("The file " & GetFileFromList(.FileName, i) & " is zero bytes in length!" _
                         & vbCr & "Are you Sure you want to add it?", vbYesNo, "Error") = vbYes Then
                    List1.AddItem GetFileFromList(.FileName, i)
                End If
        End Select

    Next
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Command3_Click()
    Call MsgBox("Example by Rudy Alex Kohn." & vbCr & _
           "This could come in handy i guess. =)" & vbCr & _
           "Contact me at rudyalexkohn@hotmail.com", 64, Me.Caption)
End Sub

Private Sub Command4_Click()
    List1.Clear
End Sub
