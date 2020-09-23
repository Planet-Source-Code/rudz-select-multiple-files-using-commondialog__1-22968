Attribute VB_Name = "Module1"
' By : Rudy Alex Kohn
' Use as you like.
' rudyalexkohn@hotmail.com
Option Explicit

Function CountFilesInList(ByVal FileList As String) As Integer
' Counts files. Returns result as integer
    Dim iCount As Integer
    Dim iPos As Integer

    iCount = 0
    For iPos = 1 To Len(FileList)
        If Mid$(FileList, iPos, 1) = Chr$(0) Then iCount = iCount + 1
    Next
    If iCount = 0 Then iCount = 1
    CountFilesInList = iCount
End Function

Function GetFileFromList(ByVal FileList As String, FileNumber As Integer) As String
' Get filename of FileNumber from FileList
    Dim iPos                As Integer
    Dim iCount              As Integer
    Dim iFileNumberStart    As Integer
    Dim iFileNumberLen      As Integer
    Dim sPath               As String

    If InStr(FileList, Chr(0)) = 0 Then
        GetFileFromList = FileList
    Else
        iCount = 0
        sPath = Left(FileList, InStr(FileList, Chr(0)) - 1)
        If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
        FileList = FileList + Chr(0)
        For iPos = 1 To Len(FileList)
            If Mid$(FileList, iPos, 1) = Chr(0) Then
                iCount = iCount + 1
                Select Case iCount
                    Case FileNumber
                        iFileNumberStart = iPos + 1
                    Case FileNumber + 1
                        iFileNumberLen = iPos - iFileNumberStart
                        Exit For
                End Select
            End If
        Next
        GetFileFromList = sPath + Mid(FileList, iFileNumberStart, iFileNumberLen)
    End If
End Function
