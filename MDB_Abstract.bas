Attribute VB_Name = "MDB_Abstract"
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

Sub getPlanYears()
'this code will go into the individual files and abstract the number of years of data to use
'number of plan years will be updated on the client files, themselves

    Dim locationList() As String 'location of data for client
    Dim lastRow As Long
    Dim j As Long
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("DirectoryExternal").Activate
    lastRow = Cells(Rows.count, 1).End(xlUp).Row
    ReDim locationList(2 To lastRow)
    If lastRow > 1 Then
        For j = 2 To lastRow
            locationList(j) = Range("C" & j).Value
            'NOTE: IF BASIC DATAPOINTS ARE ALTERED, VET LOCATION BELOW (CURRENTLY B5)
            Range("D" & j).Value = "='" & locationList(j) & "'!$B$5"
        Next j
    End If
    Range("D2:D" & lastRow).Copy
    Range("D2:D" & lastRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'error handling section: VarType 10 is error type
    For j = 2 To lastRow
        If VarType(Cells(j, 4)) = 10 Then
            Cells(j, 4).Value = 0
        End If
    Next j
    
End Sub

Sub FullPullUpdate()
    'This will create a copy of the current database after performing a pull of all the linked sheets.
    'It will save the copy in predetermined folder, and populate the PullHistory tab with information about
    'the pull, as well as a hyperlink which will provide a read only version of that pull.
    
    Dim pullDate As String
    Dim pullPath As String
    Dim pullName As String
    Dim pullType As String
    Dim pullWB As Workbook
    Dim dataPath As String
    
    Application.ScreenUpdating = False
    
    'initialization
    dataPath = "MDB Controls - Link.xlsm" 'name of file, update accordingly
    pullDate = CStr(Date)
    pullDate = Replace(pullDate, "/", "_")
    pullPath = "C:\Users\carlj\OneDrive\Documents\Projects\Work\Master Data Base\PullHistory" 'name of path, update accordingly
    pullName = pullPath & "\" & pullDate & "_FullPull"

    'create database workbook
    Workbooks.Add
    Set pullWB = ActiveWorkbook
    Workbooks(dataPath).Sheets("OutputExternal").Copy Before:=pullWB.Sheets(1)
    
    Application.DisplayAlerts = False 'this will overwrite existing files with the same name
    ActiveWorkbook.SaveAs Filename:=pullName, WriteResPassword:="datatime"
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    'Updating the pulling history, and providing a link (information placed at the top)
    Workbooks(dataPath).Sheets("PullHistory").Activate
    If Range("A2").Value = "" Then
        Range("A2").Value = Date
        Range("B2").Hyperlinks.Add Anchor:=Range("B2"), Address:=pullName & ".xlsx", TextToDisplay:="Open Data Set"
        Range("C2").Value = "Full"
    Else
        Range("A2").EntireRow.Insert
        Range("A2").Value = Date
        Range("B2").Hyperlinks.Add Anchor:=Range("B2"), Address:=pullName & ".xlsx", TextToDisplay:="Open Data Set"
        Range("C2").Value = "Full"
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub SubPullUpdate()
    'This will create a copy of the current database after performing a pull of a subset of the linked sheets.
    'It will save the copy in predetermined folder, and populate the PullHistory tab with information about
    'the pull, as well as a hyperlink which will provide a read only version of that pull.
    'See FullPullUpdate for more detailed notes.
    
    Dim pullDate As String
    Dim pullPath As String
    Dim pullName As String
    Dim pullType As String
    Dim pullWB As Workbook
    Dim dataPath As String
    
    dataPath = "MDB Controls - Link.xlsm"
    Workbooks(dataPath).Sheets("PullHistory").Activate
    pullDate = CStr(Date)
    pullDate = Replace(pullDate, "/", "_")
    pullPath = "C:\Users\carlj\OneDrive\Documents\Projects\Work\Master Data Base\PullHistory"
    pullName = pullPath & "\" & pullDate & "_SubPull"

    Workbooks.Add
    Set pullWB = ActiveWorkbook
    Workbooks(dataPath).Sheets("OutputExternal").Copy Before:=pullWB.Sheets(1)
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=pullName, WriteResPassword:="datatime"
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Workbooks(dataPath).Sheets("PullHistory").Activate
    If Range("A2").Value = "" Then
        Range("A2").Value = Date
        Range("B2").Hyperlinks.Add Anchor:=Range("B2"), Address:=pullName & ".xlsx", TextToDisplay:="Open Data Set"
        Range("C2").Value = "Sub"
    Else
        Range("A2").EntireRow.Insert
        Range("A2").Value = Date
        Range("B2").Hyperlinks.Add Anchor:=Range("B2"), Address:=pullName & ".xlsx", TextToDisplay:="Open Data Set"
        Range("C2").Value = "Sub"
    End If
    
End Sub

Sub selectClientButton()
'brings up the client selection form
    clientSelect.Show
End Sub

Sub subsetPull()
'This will update only the specified subset of clients
'Code roughly follows MasterDB_Abstract_External, which will likely have more detailed notes on the implemented processes
'Works by deleting current data for specified subset, then gathering all new data and adding to OutputExternal page

    Dim subStr As String
    Dim subList() As String
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim lastDBRow As Long 'number of rows in the database tab
    Const dataLines As Long = 4 'number of rows of data to retrieve (variable across each plan year)
    Const infoLines As Long = 4
    Dim directRows As Long
    Dim numPY As Long 'number of plan years being added to the database
    Dim locationList() As String 'location of data for client
    Dim planYears() As Long
    
    Application.ScreenUpdating = False
    
    'delete current data
    Sheets("Controls").Activate
    subStr = Range("A14").Value 'Change this range if the destination of the subset form is altered
    If subStr = "" Then
        MsgBox "No subset specified. Exiting procedure."
        Range("E7:H10").Select
        Exit Sub
    End If
    subList = Split(subStr, "; ")
    Sheets("OutputExternal").Activate
    lastDBRow = Cells(Rows.count, 1).End(xlUp).Row
    For j = 2 To lastDBRow
        For k = 0 To UBound(subList)
            If subList(k) = Range("A" & j).Value Then
                Range("A" & j).EntireRow.Delete
                j = j - 1 'update j, otherwise next loop will skip a row if the previous row was deleted
            End If
        Next k
    Next j
    lastDBRow = Cells(Rows.count, 1).End(xlUp).Row + 1 'update dataRows for new count (adding 1 to not overwrite current last row of data)
    
    'populate with updated data
    ReDim locationList(1 To (UBound(subList) + 1))
    ReDim planYears(1 To (UBound(subList) + 1))
    Sheets("DirectoryExternal").Activate
    directRows = Cells(Rows.count, 1).End(xlUp).Row
    For k = 0 To UBound(subList)
        For j = 2 To directRows
            If Range("A" & j).Value = subList(k) Then
                locationList(k + 1) = Range("C" & j).Value
                planYears(k + 1) = Range("D" & j)
            End If
        Next j
    Next k
    
    'Subset Pull performs a complete reset of only the selected clients
    Sheets("OutputExternal").Activate
    For j = 1 To UBound(locationList)
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
    
    'This next step will copy and paste to get just the values (i.e. removes the links)
    'NOTE: THIS WILL NEED TO BE UPDATED IF ADDITIONAL DATA ENTRIES ARE ADDED
    Range("A2:H" & lastDBRow).Copy
    Range("A2:H" & lastDBRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'This will format the Date type variables (easiest if you already know what columns are the date format going in)
    Range("E2:E" & lastDBRow).NumberFormat = "mm/dd/yyyy"
    Range("A1").Select
    
    'FullPullUpdate will update the PullHistory tab, and send an updated excel file to the data folder
    'directoryError will update the Issues Found section on the Control tab, indicating if there are any directory errors
    
    Call SubPullUpdate
    Call directoryError
    
    MsgBox "Specified subset updated."
    
    Application.ScreenUpdating = True
    
End Sub

Sub directoryError()
'Directory errors are mitigated in the main code.
'This code will check to see if there was a directory error. This is accomplished by if the number of plan years is 0,
'it is assumed there was a directory error associated with that client.

    Dim directRows As Long
    Dim errorsRows As Long
    Dim j As Long
    Dim clientList() As String
    Dim clientError() As Integer
    Dim errorCount As Long
    
    'get list of clients that should have data pulled from
    'the clientError list has 0 for data is missing, 1 for data is present
    'clientError is initialized as 0, possibly changed later on
    Sheets("DirectoryExternal").Activate
    directRows = Cells(Rows.count, 1).End(xlUp).Row
    ReDim clientList(2 To directRows)
    ReDim clientError(2 To directRows)
    For j = 2 To directRows
        clientList(j) = Range("A" & j).Value
        clientError(j) = 0
    Next j
    
    'if a directory error occured, there will be a 0 for the number of plan years
    For j = 2 To directRows
        If Range("D" & j).Value <> 0 Then
            clientError(j) = 1
        End If
    Next j
    
    'if clients remains a 0 (i.e. data not present -> directoy error assumed), they are added individually to error section
    Sheets("Controls").Activate
    If Range("A18") <> "" Then
        errorsRows = Cells(Rows.count, 1).End(xlUp).Row
        Range("A18:A" & errorsRows).Clear
    End If
    
    'errorCount initialized at 18 due to error section location on page
    errorCount = 18
    For j = 2 To directRows
        If clientError(j) = 0 Then
            Range("A" & errorCount).Value = clientList(j)
            errorCount = errorCount + 1
        End If
    Next j
    
End Sub

Vdu  `+ÃGu  –ªWPu  –ªWPu  ‡ªWPu  ‡ªWPu  ªWPu  ªWPu   ºWPu   ºWPu  ºWPu  ºWPu   ºWPu   ºWPu  0ºWPu  0ºWPu  @ºWPu  @ºWPu  PºWPu  PºWPu  `ºWPu  `ºWPu  pºWPu  pºWPu  ÄºWPu  ÄºWPu  êºWPu  êºWPu  †ºWPu  †ºWPu  ∞ºWPu  ∞ºWPu  ¿ºWPu  ¿ºWPu  –ºWPu  –ºWPu  ‡ºWPu  ‡ºWPu  ºWPu  ºWPu   ΩWPu   ΩWPu  ΩWPu  ΩWPu   ΩWPu   ΩWPu  0ΩWPu  0ΩWPu  @ΩWPu  @ΩWPu  PΩWPu  PΩWPu  `ΩWPu  `ΩWPu  pΩWPu  pΩWPu  ÄΩWPu  ÄΩWPu  êΩWPu  êΩWPu  †ΩWPu  †ΩWPu  