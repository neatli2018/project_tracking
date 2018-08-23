Attribute VB_Name = "Module11"
Sub ProjectRanking()
'
' Module11 collects the project names from the top of the Prioritization Matrix page and the scores associated with them
' and puts them in a dictionary. It then uses the project names from the Portfolio page to pull those scores out of the
' dictionary. Once it populates the scores on the Portfolio page it sorts them and replaces the score with the numerical
' ranking of the projects on the portfolio page.
'
    Dim numOfProjects As Integer
    Dim lastRow As Integer
    Dim expectedNonProjects As Integer
    Dim rank As Integer
    Dim k As Integer
    
    Dim projectsDict As Object
    Set projectsDict = CreateObject("Scripting.Dictionary")
    Dim includeScore As Boolean
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = Worksheets("Portfolio")
    Set ws2 = Worksheets("Prioritization Matrix")
        
    Dim tbl As ListObject
    Dim sortcolumn As Range
'    Set sortcolumn = Range("Table3[Project Name]")
'    With tbl.Sort
'       .SortFields.Clear
'       .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
'       .Header = xlYes
'       .Apply
'    End With
    
'   What are the names of the table and the column you want the ranks to go in?
    tblName = "Table3"
    rankColName = "Column1"
    
    Set tbl = ws1.ListObjects(tblName)
    Set sortcolumn = Range("Table3[Column1]")
    
    
'   Set to false if you do not want the score listed on Sheet1, set to true if you do.
    includeScore = False
        
    'Count non blanks in the first row, subtract the ones that are not project names (expectedNonProjects)
    expectedNonProjects = 1
    ws2.Select
    numOfProjects = WorksheetFunction.CountA(Rows("1:1")) - expectedNonProjects
    
    
    '   Capture the names of the projects and their scores in the dictionary
    With ws2
        
        For i = 1 To numOfProjects
            projectsDict.Add .Cells(1, 4 + 3 * (i - 1)).Value(), .Cells(7, 3 + 3 * i).Value()
            Debug.Print i & ": " & projectsDict(.Cells(1, 4 + 3 * (i - 1)).Value())
        Next i
    
    End With
    
    '   Go to Sheet1 and get the score for each name in the list of projects. Generate an error if a name is not found.
    ws1.Select
    With ws1
    
        For i = 1 To numOfProjects
            tempStr = .Cells(i + 4, 1).Value()
            If projectsDict.Exists(tempStr) Then
                .Cells(i + 4, 2).Value() = projectsDict(tempStr)
            Else: MsgBox (tempStr & " is not entered correctly on Sheet2")
                .Cells(i + 4, 2).Value() = "Score is not found"
            
            End If
        Next i
        Range("A4:i" & numOfProjects + 4).Sort key1:=Range("B4:B" & numOfProjects + 4), order1:=xlAscending, Header:=xlYes
        Set sortcolumn = Range("Table3[Column1]")
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
            
'   Here, it enters a rank based on the scores. If you want to eliminate the scores, swap commented line
        rank = 1
        k = 2
        Set sortcolumn = Range("Table3[Column1]")
        If includeScore Then
            k = 3
            Set sortcolumn = Range("Table3[Column2]")
        End If
            
        For i = 1 To numOfProjects
            If .Cells(i + 4, 2).Value() = "Score is not found" Then
                .Cells(i + 4, k).Value() = "Project Name not found"
                
            Else
                 .Cells(i + 4, k).Value() = rank
                
                rank = rank + 1
            End If
            
        Next i
        
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End With

End Sub

