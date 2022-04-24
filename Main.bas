Attribute VB_Name = "Main"
Option Explicit

Sub MasterDB_Abstract_External()
'this code will use a directory to link to input pages on different workbooks,
'compiling a master database page, which will be extracted to another location
'as a read-only file

    Dim locationList() As String 'location of data for client
    Dim planYears() As Long 'store number of plan years for each client
    Dim lastRow As Long 'last row of data for the directory page
    Const dataLines As Long = 4 'number of rows of data to retrieve (variable across each time period, CHANGE with amount of rows)
    Const infoLines As Long = 4 'number of rows of basic information to retrieve (constant across each time period, CHANGE with amount of rows)
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim numPY As Long 'number of plan years being added to the database
    Dim lastDBRow As Long 'number of rows of data currently in the total database
    
    Application.ScreenUpdating = False
    
    'clear entire database
    Sheets("OutputExternal").Activate
    lastDBRow = Cells(Rows.count, 1).End(xlUp).Row
    If lastDBRow <> 1 Then
        Range("A2:" & "A" & lastDBRow).EntireRow.Clear
    End If
    
    'get data about which clients to consolidate data for and where they're located
    Sheets("DirectoryExternal").Activate
    lastRow = Cells(Rows.count, 1).End(xlUp).Row
    Call getPlanYears 'updates number of plan years of data available for each client
    ReDim locationList(2 To lastRow) 'data array of information storage location
    ReDim planYears(2 To lastRow) 'data array of number of plan years for each client
    If lastRow > 1 Then
        For j = 2 To lastRow
            locationList(j) = Range("C" & j).Value
            planYears(j) = Range("D" & j)
        Next j
    End If
    
    'Reminder: complete refresh of the database each time (hence the next variable being set to 2)
    lastDBRow = 2
    Sheets("OutputExternal").Activate
    For j = 2 To lastRow
        numPY = planYears(j) 'the plan years of data available for each client, j
        If numPY <> 0 Then 'if directory error occurs, numPY will be 0
            For k = 1 To infoLines 'Recall this is pulling the basic information (constant across plan years)
                For l = 1 To numPY
                    Cells(lastDBRow + l - 1, k).Value = "='" & locationList(j) & "'!" & Cells(k, 2).Address
                Next l
            Next k
            'next double for-loop will pull variable information from each plan year
            For k = 1 To numPY
                For l = 1 To dataLines 'Recall dataLines is the amount of data points, and this number can be altered in the variables above
                    Cells(lastDBRow, infoLines + l).Value = "='" & locationList(j) & "'!" & Cells(l, 4 + k).Address
                Next l
                lastDBRow = lastDBRow + 1
            Next k
        End If
    Next j
    'After some experimenting, decided to go with 2 separate loops, one for the basic informatoin, one for the variable information
    'The basic information is only available once, whereas the variable information is different for each plan year available
    'Additionally, the loops themselves run on different range parameters
    
    'This next step will copy and paste to get just the values (i.e. removes the links)
    'NOTE: THIS WILL NEED TO BE UPDATED IF NUMBER OF ENTRIES CHANGES
    Range("A2:H" & lastDBRow).Copy
    Range("A2:H" & lastDBRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'This will format the Date type variables (easiest if you already know what columns are the date format going in)
    Range("E2:E" & lastDBRow).NumberFormat = "mm/dd/yyyy"
    Range("A1").Select
    
    'FullPullUpdate will update the PullHistory tab, and send an updated excel file to the data folder
    'directoryError will update the Issues Found section on the Control tab, indicating if there are any directory errors
    Call FullPullUpdate
    Call directoryError
    
    MsgBox "The entire database has been updated."
    
    Application.ScreenUpdating = True
    
End Sub
  �s1�N<Pu  �R<Pu  ����8U�0R<Pu  PM<Pu  �����O�@Q<Pu  0N<Pu  �������T<Pu  0T<Pu  ����em��O<Pu  �O<Pu  �   ���O<Pu  PP<Pu  �     l��U<Pu  �Y<Pu  ����1 �Ap�Pu  ��Pu  �   ,0gCp�Pu  ��Pu  ��������0�Pu  ��Pu  �������                �   �                �   l��                �   1JWW                �   u                  �   u3Qc                �   O+Ny