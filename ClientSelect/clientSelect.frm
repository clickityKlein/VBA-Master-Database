VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clientSelect 
   Caption         =   "Client Select"
   ClientHeight    =   1728
   ClientLeft      =   -72
   ClientTop       =   -312
   ClientWidth     =   6816
   OleObjectBlob   =   "clientSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clientSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim clients As Range
    Dim directLen As Long
    Dim sht As Worksheet
    
    Sheets("DirectoryExternal").Activate
    directLen = Range("A1").End(xlDown).Row
    
    With clientSelect.clientList
        .RowSource = Range(Cells(2, 1), Cells(directLen, 1)).Address
        .ColumnHeads = True
        .ColumnCount = 1
        .MultiSelect = 2
    End With
    Sheets("Controls").Activate
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim j As Integer
    Dim count As Integer
    count = 0
    Range("A14").Select
    For j = 0 To Me.clientList.ListCount - 1
        If Me.clientList.Selected(j) = True Then
            If count = 0 Then
                ActiveCell = Me.clientList.List(j)
                count = count + 1
            Else
                ActiveCell = ActiveCell & "; " & Me.clientList.List(j)
            End If
        End If
    Next j
    Unload Me
End Sub


