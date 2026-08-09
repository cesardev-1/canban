Global G_ManagedItemType As String
Global G_TextShortcut As String
Global G_UsedShortcut As Boolean
Global binFrozenFocus As Boolean
Global oSeleccionListener As Object
Global oSheetHeard As Object

Sub ToggleInterface
    Dim oDoc As Object
    Dim oLayout As Object
    
    oDoc = ThisComponent
    oLayout = oDoc.CurrentController.Frame.LayoutManager
    
    ' If the interface is visible, it hides it. If it is hidden, it shows it.'
    If oLayout.isVisible() Then
        oLayout.setVisible(False)
    Else
        oLayout.setVisible(True)
    End If
End Sub


REM =========================================================================
REM DATA-DRIVEN RENDERING ENGINE (MATRIX REPLACEMENT)
REM =========================================================================
Sub DrawingBoard()
    Dim oDoc As Object, oVisor As Object, oData As Object
    Dim sActivity As String, sStateA As String, sStateB As String
    Dim TotalRowsData As Long, i As Long
    Dim CurrentRowA As Long, CurrentRowB As Long
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Disable screen refresh for maximum speed
    oDoc.addActionLock()
    
    ' 2. Clean old display ranges and hidden ID ranges (Rows 4 to 300 -> indices 3 to 299)
    ' We clean columns A, B, C, D (Tasks, Subtasks, Descriptions) and E, F, G (Hidden IDs)
    For i = 3 To 299
        oVisor.getCellByPosition(0, i).String = "" ' Col A
        oVisor.getCellByPosition(1, i).String = "" ' Col B
        oVisor.getCellByPosition(2, i).String = "" ' Col C
        oVisor.getCellByPosition(3, i).String = "" ' Col D
        oVisor.getCellByPosition(4, i).Value = 0   ' Col E (ID hidden Tasks Col A)
        oVisor.getCellByPosition(5, i).Value = 0   ' Col F (ID Oculto Tasks Col B)
    Next i
    
    ' 3. Read View Filters
    sActivity = Trim(oVisor.getCellByPosition(0, 1).String) ' A2 (combined cell A2:D2)
    sStateA = Trim(oVisor.getCellByPosition(0, 2).String)   ' A3 (Head State 1)
    sStateB = Trim(oVisor.getCellByPosition(1, 2).String)   ' B3 (Head State 2)
    
    If sActivity = "" Then
        oDoc.removeActionLock()
        Exit Sub
    End If
    
    ' 4. Determine the size of the database
    Dim oCursor As Object
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRowsData = oCursor.RangeAddress.EndRow + 1
    
    ' Initialize write pointers for the Viewer (Start at row 4 -> index 3)
    CurrentRowA = 3
    CurrentRowB = 3
    
    ' 5. Traverse the Database and Inject into the corresponding Columns
    For i = 1 To TotalRowsData - 1
        Dim sRegActivity As String, sRegTask As String, sRegState As String
        Dim IdTask As Long
        
        sRegActivity = Trim(oData.getCellByPosition(0, i).String) ' Col A
        sRegTask = Trim(oData.getCellByPosition(1, i).String)     ' Col B
        sRegState = Trim(oData.getCellByPosition(2, i).String)    ' Col C
        IdTask = CLng(oData.getCellByPosition(5, i).Value)       ' Col F (Unique Immutable ID)
        
	' If it coincides with the activity under consultation
        If sRegActivity = sActivity Then
            
	    ' Does it belong to the first State (Column A of the View)?
            If sRegState = sStateA And CurrentRowA <= 299 Then
                oVisor.getCellByPosition(0, CurrentRowA).String = sRegTask
                oVisor.getCellByPosition(4, CurrentRowA).Value = IdTask ' Safe ID in Column E
                CurrentRowA = CurrentRowA + 1
            End If
                
	    ' Does it belong to the second State (Column B of the View)?
            If sRegState = sStateB And CurrentRowB <= 299 Then
                oVisor.getCellByPosition(1, CurrentRowB).String = sRegTask
                oVisor.getCellByPosition(5, CurrentRowB).Value = IdTask ' Safe ID in Column E
                CurrentRowB = CurrentRowB + 1
            End If
            
        End If
    Next i
    
    ' 6. Release screen lock and refresh
    oDoc.removeActionLock()
    oDoc.calculateAll()
    
' ---- SYNCHRONIZE DESCRIPTION WHEN REDRAWING ----
    Dim IdCurrentFocus As Long
    Dim oCurrentSel As Object
    
    oCurrentSel = oDoc.CurrentSelection
    If oCurrentSel.supportsService("com.sun.star.sheet.SheetCell") Then
	' STAR REUSE #6: Read the ID of the current physical position
        IdCurrentFocus = GetIdFromPosition(oVisor, oCurrentSel.CellAddress.Row, oCurrentSel.CellAddress.Column)
        
	' STAR REUSE #4: Inject the legitimate description of that ID
        RunInjectionDescription(oVisor, IdCurrentFocus)
        
	' Also update debug cell E1
        oVisor.getCellByPosition(4, 0).Value = IdCurrentFocus
    End If
    ' -----------------------------------------------------------------------
    
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM --- OPTIMIZED DATA SORTING ENGINE ---
REM =========================================================================
Sub CustomDataSort()
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRowsData As Long, i As Long, j As Long
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' =========================================================================
    ' STEP 1: DATABASE SECURITY AND HEALING
    ' Before moving a single row, we fix any missing IDs or Orders.
    ' =========================================================================
    AssignMissingIdsBD()
    AssignMissingOrdersBD()
    
    ' 2. Find the actual row limit in the post-curing data sheet
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRowsData = oCursor.RangeAddress.EndRow + 1
    
    ' If only the headers are there or it's empty, we can leave safely.
    If TotalRowsData <= 2 Then Exit Sub
    
    ' 3. Load priority reference lists from columns H and J
    Dim listctivities() As Variant, ListStates() As Variant
    listctivities = oData.getCellRangeByName("H2:H100").getDataArray()
    ListStates = oData.getCellRangeByName("J2:J20").getDataArray()
    
    ' 4. Extract the entire current data matrix (Columns A to F, starting from row 2)
    ' Expanded strictly to column F (index 5) to drag the immutable IDs
    Dim oFullRange As Object
    Dim DataMatrix() As Variant
    oFullRange = oData.getCellRangeByPosition(0, 1, 5, TotalRowsData - 1)
    DataMatrix = oFullRange.getDataArray()
    
    Dim nPatients As Long
    nPatients = UBound(DataMatrix)
    
    Dim sActivity As String, sState As String
    Dim temporal As Variant
    Dim pAct1 As Long, pAct2 As Long
    Dim pSta1 As Long, pSta2 As Long
    Dim ord1 As Double, ord2 As Double
    
    ' 5. Sort algorithm by memory occurrence weight
    For i = 0 To nPatients - 1
        For j = i + 1 To nPatients
            ' --- TASK 1 ---
            sActivity = DataMatrix(i)(0)
            sState = DataMatrix(i)(2)
            ord1 = CDbl(DataMatrix(i)(4))
            
            pAct1 = GetListIndex(sActivity, listctivities)
            pSta1 = GetListIndex(sState, ListStates)
            
            ' --- TASK 2 ---
            sActivity = DataMatrix(j)(0)
            sState = DataMatrix(j)(2)
            ord2 = CDbl(DataMatrix(j)(4))
            
            pAct2 = GetListIndex(sActivity, listctivities)
            pSta2 = GetListIndex(sState, ListStates)
            
	    ' Ordering criteria for your design philosophy:
            ' 1° Activity Hierarchy (Column H)
	    ' 2° State Hierarchy (Column J)
            ' 3° Internal Position / Interpolation (Column E)
            If (pAct1 > pAct2) Or _
               (pAct1 = pAct2 And pSta1 > pSta2) Or _
               (pAct1 = pAct2 And pSta1 = pSta2 And ord1 > ord2) Then
                ' Full row switch in memory
                temporal = DataMatrix(i)
                DataMatrix(i) = DataMatrix(j)
                DataMatrix(j) = temporal
            End If
        Next j
    Next i
    
    ' =========================================================================
    ' CORRECTED LOCATION: GLOBAL SEQUENTIAL INDEX WITH ANTI-GHOST FILTER (F2 PRO)
    ' Normalizes the entire database, skipping empty residual rows
    ' =========================================================================
    For i = 0 To nPatients
        ' We check if the current row of the array actually contains a valid task
        If Trim(DataMatrix(i)(0)) <> "" And Trim(DataMatrix(i)(1)) <> "" Then
            ' If it is a real task, it receives its global sequential index run.
            DataMatrix(i)(4) = CDbl(i + 1)
        Else
            ' IF THE ROW IS EMPTY (GHOST): We clear any residual numeric data from memory
            DataMatrix(i)(4) = "" ' Delete the Order (Col E)
            DataMatrix(i)(5) = "" ' Delete the ID (Col F)
        End If
    Next i
    ' =========================================================================
    
    ' 6. Pour the perfectly ordered and complete matrix back into the cells
    oFullRange.setDataArray(DataMatrix)
    
    ' 7. Automatic synchronization: Redraw the Kanban canvas of the Viewer
    DrawingBoard()
End Sub

REM =========================================================================
REM AUXILIARY FUNCTION: SEARCH FOR THE WEIGHT/INDEX WITHIN THE REFERENCE LISTS
REM =========================================================================
Function GetListIndex(sValue As String, ByRef list() As Variant) As Long
    Dim i As Long
    sValue = Trim(LCase(sValue))
    
    For i = 0 To UBound(list)
        If Trim(list(i)(0)) <> "" Then
            If Trim(LCase(list(i)(0))) = sValue Then
                GetListIndex = i ' Returns the physical position (0, 1, 2...) as its priority level
                Exit Function
            End If
        End If
    Next i
    
    ' If for any reason the activity/status is not in list H or J, it is assigned the lowest priority.
    GetListIndex = 999 
End Function



REM =========================================================================
REM UNIFIED CASCADE ERASURE ENGINE
REM =========================================================================
Sub RunCascadeDelete(sItem As String, sTypeItem As String)
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim TaskCounter As Long: TaskCounter = 0
    Dim answer As Integer
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Security validation for Base States
    If sTypeItem = "State" Then
        If sItem = "In progress" Or sItem = "To do" Then
            answer = MsgBox(" Warning: You are trying to delete a system base state ('" & sItem & "')." & Chr(13) & _
                               "This could alter the default display of the columns." & Chr(13) & _
                               "Do you want to continue anyway?", 4 + 32, "Baseline State Validation")
            If answer <> 6 Then Exit Sub
        End If
    End If
    
    ' 2. Finding database limits
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    ' 3. Count affected tasks
    For i = 1 To TotalRows - 1
        Dim sValueCompare As String
        If sTypeItem = "Activity" Then
            sValueCompare = oData.getCellByPosition(0, i).String ' Column A
        Else
            sValueCompare = oData.getCellByPosition(2, i).String ' Column C
        End If
        
        If sValueCompare = sItem Then
            TaskCounter = TaskCounter + 1
        End If
    Next i
    
    ' 4. Critical and irreversible safety confirmation
    answer = MsgBox("⚠️CRITICAL ELIMINATION WARNING ⚠️" & Chr(13) & Chr(13) & _
                       "You're going to delete " & sTypeItem & ": """ & sItem & """" & Chr(13) & _
                       "They will be permanently destroyed [" & TaskCounter & "] associated tasks." & Chr(13) & Chr(13) & _
                       "¿You confirm the operation?", 4 + 16 + 256, "Confirm Data Cascade")
                       
    If answer = 6 Then ' YES
        oDoc.addActionLock()
        
        ' A. Delete tasks from the bottom up to maintain stable row indexes
        For i = TotalRows - 1 To 1 Step -1
            Dim sValueCompare2 As String
            If sTypeItem = "Activity" Then
                sValueCompare2 = oData.getCellByPosition(0, i).String
            Else
                sValueCompare2 = oData.getCellByPosition(2, i).String
            End If
            
            If LCase(Trim(sValueCompare2)) = LCase(Trim(sItem)) Then
                oData.Rows.removeByIndex(i, 1)
            End If
        Next i
        
        ' B. Delete the item from the configuration catalog
        If sTypeItem = "Activity" Then
            Dim mCatAct() As Variant
            mCatAct = oData.getCellRangeByName("H2:H100").getDataArray()
            For i = 0 To UBound(mCatAct)
                If LCase(Trim(mCatAct(i)(0))) = LCase(Trim(sItem)) Then
                    oData.getCellByPosition(7, i + 1).String = "" ' Column H (índice 7)
                    Exit For
                End If
            Next i
        Else
            Dim mCatEst() As Variant
            mCatEst = oData.getCellRangeByName("J2:J20").getDataArray()
            For i = 0 To UBound(mCatEst)
                If LCase(Trim(mCatEst(i)(0))) = LCase(Trim(sItem)) Then
                    oData.getCellByPosition(9, i + 1).String = "" ' Column J (índice 9)
                    Exit For
                End If
            Next i
        End If
        
        ' C. Reorder and redraw the Kanban board
        CustomDataSort()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        MsgBox sTypeItem & " and their tasks have been purged from the database.", 64, "Purge Completed"
    End If
End Sub






' --- 1. ACTIVATE THE LISTENER (This should run when opening the file) ---
Sub StartListenerSelection()
    Dim oDoc As Object
    oDoc = ThisComponent
    oSheetHeard = oDoc.Sheets.getByName("view")
    
    ' Create the selection event listener object
    oSeleccionListener = CreateUnoListener("VisorSeleccion_", "com.sun.star.view.XSelectionChangeListener")
    oDoc.CurrentController.addSelectionChangeListener(oSeleccionListener)
End Sub

Sub RemoveListenerSelection()
    On Error Resume Next
    ThisComponent.CurrentController.removeSelectionChangeListener(oSeleccionListener)
End Sub


REM =========================================================================
REM ISOLATED COMPONENT: PURE EXTRACTOR OF ID DESCE HIDDEN COLUMNS (E=4, F=5)
REM =========================================================================
Function GetIdFromPosition(oVisor As Object, nRow As Long, nColumn As Long) As Long
    Dim IdFound As Long
    IdFound = 0
    
    ' The valid range of the Kanban is Rows 4 to 300 (indices 3 to 299)
    If nRow >= 3 And nRow <= 299 Then
        If nColumn = 0 Then
            ' Column A (0) -> Read ID Column E (4)
            IdFound = CLng(oVisor.getCellByPosition(4, nRow).Value)
        ElseIf nColumn = 1 Then
            ' Column B (1) -> Read ID Column F (5)
            IdFound = CLng(oVisor.getCellByPosition(5, nRow).Value)
        End If
    End If
    
    GetIdFromPosition = IdFound
End Function

REM =========================================================================
REM ULTRA-SIMPLIFIED LISTENER FOR ID LINK TESTING
REM =========================================================================
Sub VisorSeleccion_selectionChanged(oEvent)
    Dim oDoc As Object, oVisor As Object, oSel As Object
    Dim nRow As Long, nColumn As Long
    Dim sActiveSheet As String
    Dim IdDetected As Long
    
    On Error Resume Next
    
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    
    ' 1. ESCAPE: Validate that it is a single cell
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    
    ' 2. ESCAPE: Operate strictly within the "view" sheet
    If oSel.Spreadsheet.Name <> "view" Then Exit Sub
    
    nRow = oSel.CellAddress.Row
    nColumn = oSel.CellAddress.Column
    oVisor = oDoc.Sheets.getByName("view")
    
    ' 3. OBTAIN THE ID USING OUR SEPARATE COMPONENT
    IdDetected = GetIdFromPosition(oVisor, nRow, nColumn)
    
    ' 4. DEPOSIT THE ID IN E1 FOR VISUAL VERIFICATION
    ' Cell E1 in zero base is Column 4, Row 0
    oVisor.getCellByPosition(4, 0).Value = IdDetected
    
    ' Launch Process 1: Description Management
    RunInjectionDescription(oVisor, IdDetected)
    
    ' Refresh screen subtly
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM SHEET EVENT: CONTENT MODIFIED ON THE VIEWER SHEET (KANBAN REFRESH)
REM =========================================================================
Sub Visor_onContentChanged(oEvent As Object)
    Dim nRow As Long, nCol As Long
    
    On Error Resume Next
    
    ' 1. ESCAPE: Validate that the change comes from a single physical cell
    If Not oEvent.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    
    nRow = oEvent.CellAddress.Row
    nCol = oEvent.CellAddress.Column
    
    ' 2. EVALUATE CRITICAL FILTERING COORDINATES:
    ' - A2 (Row 1, Col 0) o C2 (Row 1, Col 2) -> Active Activity Filter
    ' - A3 (Row 2, Col 0) -> Left State Filter (Column A)
    ' - B3 (Row 2, Col 1) -> Right-of-State Filter (Column B)
    If (nRow = 1 And (nCol = 0 Or nCol = 2)) Or _
       (nRow = 2 And nCol = 0) Or _
       (nRow = 2 And nCol = 1) Then
       
        ' Temporarily freeze the selection listener to avoid visual loops
        ' while DrawingBoard rewrites the viewer cells
        binFrozenFocus = True
        
        ' Synchronously redraw the dashboard to reflect the new filter
        Call CustomDataSort()
                
        ' Reactivate the selection listener once the redraw is complete.
        binFrozenFocus = False
    End If
End Sub

REM =========================================================================
REM PHASE 2 - COMPONENT 1: GAP FINDER AND AVAILABLE ID GENERATOR
REM =========================================================================
Function GetNextAvailableId() As Long
    Dim oDoc As Object, oData As Object
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim IdCandidate As Long
    Dim IdExists As Boolean
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' Finding the actual limit of the data sheet
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    ' If there are no data rows (header only), the first ID is 1
    If TotalRows <= 2 Then
        GetNextAvailableId = 1
        Exit Function
    End If
    
    ' LOGIC FOR REUSING GAPS:
    ' We start testing from ID = 1 upwards.
    ' The first number that we do NOT find in column F, that will be our chosen ID.
    IdCandidate = 1
    Do
        IdExists = False
        'Traverse column F (index 5) looking to see if the IdCandidate is already taken 
        For i = 1 To TotalRows - 1
            If CLng(oData.getCellByPosition(5, i).Value) = IdCandidate Then
                IdExists = True
                Exit For ' It's already taken, there's no need to keep looking this time.
            End If
        Next i
        
        ' If it didn't exist, we've found a gap or the end of the sequence!
        If Not IdExists Then
            Exit Do
        End If
        
        ' If it existed, we tried the next number
        IdCandidate = IdCandidate + 1
    Loop
    
    GetNextAvailableId = IdCandidate
End Function


REM =========================================================================
REM PHASE 2 - COMPONENT 2: COMPLETE SWEEPER OF REGISTERS WITHOUT ID
REM =========================================================================
Sub AssignMissingIdsBD()
    Dim oDoc As Object, oData As Object
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sActivity As String, sTask As String, sState As String
    Dim idCurrent As Long, NewId As Long
    Dim AssignedCount As Long
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    AssignedCount = 0
    
    ' Disable screen updating for millisecond processing.
    oDoc.addActionLock()
    
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    ' Iterate from row 2 (index 1) onwards.
    For i = 1 To TotalRows - 1
        ' Read mandatory validation fields
        sActivity = Trim(oData.getCellByPosition(0, i).String) ' Col A
        sTask = Trim(oData.getCellByPosition(1, i).String)     ' Col B
        sState = Trim(oData.getCellByPosition(2, i).String)    ' Col C
        idCurrent = oData.getCellByPosition(5, i).Value          ' Col F
        
        ' VALIDATION RULE:
        ' If Activity, Task, and Status are present (fields populated) AND the ID is 0 or empty.
        If sActivity <> "" And sTask <> "" And sState <> "" And idCurrent = 0 Then
            
            ' Check the next available ID (reusing gaps if any exist)
            NewId = GetNextAvailableId()
            
            ' Inject the ID into Column F
            oData.getCellByPosition(5, i).Value = NewId
            AssignedCount = AssignedCount + 1
        End If
    Next i
    
    oDoc.removeActionLock()
    
    ' Subtle confirmation message if executed manually and something was assigned.
    If AssignedCount > 0 Then
        MsgBox "The records have been audited. IDs successfully assigned.: " & AssignedCount, 64, "ID Maintenance"
    End If
End Sub

REM =========================================================================
REM ISOLATED PROCESS: INJECT TITLE (D3) AND DESCRIPTION (D4) INTO THE VIEWER
REM =========================================================================
Sub RunInjectionDescription(oVisor As Object, idBusqueda As Long)
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRowsData As Long, i As Long
    Dim sDescription As String
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Clear column D completely before painting (only cell D4, as it is a merged range).
    ' This ensures that no visual "ghosts" of previous cells remain.
    oVisor.getCellByPosition(3, 3).String = ""
    oVisor.getCellByPosition(3, 2).String = "" 
    
    ' 2. If the ID is 0 (empty cell), clear the field and exit.
    If idBusqueda = 0 Then Exit Sub
    
    ' 3. Search for the description in the database by ID.
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRowsData = oCursor.RangeAddress.EndRow + 1
    
    sTitle = ""
    sDescription = ""
    For i = 1 To TotalRowsData - 1
        If CLng(oData.getCellByPosition(5, i).Value) = idBusqueda Then
            sDescription = oData.getCellByPosition(3, i).String ' Column D
            sTitle = oData.getCellByPosition(1, i).String      ' Column B
            Exit For
        End If
    Next i
    
    ' 4. Write the result in the Viewer.
    oVisor.getCellByPosition(3, 2).String = sTitle
    oVisor.getCellByPosition(3, 3).String = sDescription
End Sub

REM =========================================================================
REM UNIQUE PRIORITY ENGINE: GLOBAL SEQUENTIAL INDEX VERSION (F2 PRO)
REM =========================================================================
Function GetNextOrder(sActivity As String, sState As String, Optional nPosWanted As Long) As Double
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim CurrentTotalRows As Long, i As Long
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    CurrentTotalRows = oCursor.RangeAddress.EndRow + 1
    
    If IsMissing(nPosWanted) Then nPosWanted = 0
    
    ' =========================================================================
    ' SCENARIO A: GLOBAL INTERPOLATION WITHIN KANBAN
    ' =========================================================================
    If nPosWanted > 0 Then
        Dim mHistoric() As Variant, oTaskRange As Object
        oTaskRange = oData.getCellRangeByPosition(0, 1, 4, CurrentTotalRows - 1)
        mHistoric = oTaskRange.getDataArray()
        
        ' We filter to locate the global indices for the tasks of this specific group
        Dim mGlobalEndices() As Long, cFiltered As Long : cFiltered = 0
        For i = 0 To UBound(mHistoric)
            If Trim(mHistoric(i)(0)) = sActivity And Trim(mHistoric(i)(2)) = sState Then
                ReDim Preserve mGlobalEndices(cFiltered)
                mGlobalEndices(cFiltered) = i ' We store the physical position of the row
                cFiltered = cFiltered + 1
            End If
        Next i
        
        If cFiltered > 0 Then
            ' Determine the global weights in column E based on the desired position
            If nPosWanted <= 1 Then
                ' At the start of the group: Weight of the first group task minus 0.1
                GetNextOrder = CDbl(mHistoric(mGlobalEndices(0))(4)) - 0.1
            ElseIf nPosWanted > cFiltered Then
                ' At the end of the group: Weight of the last group task plus 0.1
                GetNextOrder = CDbl(mHistoric(mGlobalEndices(cFiltered - 1))(4)) + 0.1
            Else
                ' In between: Decimal midpoint between the chosen task and the previous one
                Dim PreviousOrder As Double, NextOrder As Double
                PreviousOrder = CDbl(mHistoric(mGlobalEndices(nPosWanted - 2))(4))
                NextOrder = CDbl(mHistoric(mGlobalEndices(nPosWanted - 1))(4))
                GetNextOrder = PreviousOrder + ((NextOrder - PreviousOrder) / 2)
            End If
            Exit Function
        End If
    End If
    
    ' =========================================================================
    ' SCENARIO B: CONSECUTIVE AT THE VERY END OF THE TABLE
    ' =========================================================================
    ' If the table is empty (header only), start at 1; otherwise, go to the end of the list
    If CurrentTotalRows <= 2 Then
        GetNextOrder = 1.0
    Else
        GetNextOrder = CDbl(CurrentTotalRows)
    End If
End Function

REM =========================================================================
REM DEFENSIVE COMPONENT: REPAIRS AND ASSIGNS MISSING SEQUENCE NUMBERS (F2)
REM =========================================================================
Sub AssignMissingOrdersBD()
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRowsData As Long, i As Long
    Dim sActivity As String, sState As String
    Dim oOrderCell As Object
    Dim nAssignedOrder As Double
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Find the actual limit of rows in the "datos" sheet.
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRowsData = oCursor.RangeAddress.EndRow + 1
    
    ' If only the headers are present, there is nothing to audit
    If TotalRowsData <= 2 Then Exit Sub
    
    ' 2. Scanning row by row, analyzing column E (index 4)
    For i = 1 To TotalRowsData - 1
        sActivity = Trim(oData.getCellByPosition(0, i).String) ' Col A
        sState = Trim(oData.getCellByPosition(2, i).String)    ' Col C
        oOrderCell = oData.getCellByPosition(4, i)             ' Col E
        
        ' A row is valid if it has at least an Activity and a Task Name
        If sActivity <> "" And Trim(oData.getCellByPosition(1, i).String) <> "" Then
            
            ' Detect whether the order is null, empty, or zero
            If oOrderCell.Type = com.sun.star.table.CellContentType.EMPTY Or oOrderCell.Value = 0 Then
                
                ' We ask the expert component for the next integer for this group
                ' Note: We do not pass the third parameter because we want it to go to the end of its list.
                nAssignedOrder = GetNextOrder(sActivity, sState)
                
                ' We inject the curated value into the database
                oOrderCell.Value = nAssignedOrder
                
            End If
        End If
    Next i
End Sub

REM =========================================================================
REM AUTONOMOUS CENTRAL SERVICE HUB FOR DIALOGUES (ATOMIC STRUCTURE)
REM =========================================================================
Sub InitializeDialogServices(oDialogue As Object, IdActive As Long)
    Dim oDoc As Object, oVisor As Object, oData As Object
    Dim oModel As Object, oSel As Object
    
    ' 1. GLOBAL ENVIRONMENT LOADING (Always used)
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oModel = oDialogue.Model
    
    REM ---------------------------------------------------------------------
    REM DYNAMIC VISIBILITY SERVICE: SHOW/HIDE ACTION BUTTONS
    REM ---------------------------------------------------------------------
    ' If it is a new record (IdActive = 0), we hide the physical and logical deletion buttons
    If oModel.hasByName("BtnDelete") Then
        oDialogue.getControl("BtnDelete").setVisible(IdActive > 0)
    End If
    
    If oModel.hasByName("BtnTrash") Then
        oDialogue.getControl("BtnTrash").setVisible(IdActive > 0)
    End If
    
    ' Disable screen updates for general performance
    oDoc.addActionLock()
    
    REM ---------------------------------------------------------------------
    REM SERVICE 1: POPULATE ACTIVITIES COMBO BOX
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbActivity") Then
        ' We pass the model and the view separately to the external macro
        Call ServPopulateCmbActivity(oModel.getByName("CmbActivity"), oDialogue.getControl("CmbActivity"), oData, IdActive)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 2: POPULATE STATES COMBO BOX
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbState") Then
        ' We pass the model and the view separately to the external macro
        Call ServPopulateCmbState(oModel.getByName("CmbState"), oDialogue.getControl("CmbState"), oData, IdActive)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 3: INJECT TITLE (Contextual to Active ID or Selection)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitle") Then
        Dim oCtrlTitle As Object
        oCtrlTitle = oModel.getByName("TxtTitle")
        ' Clean and atomic: It only handles the title text.
        Call ServManageTxtTitle(oCtrlTitle, IdActive, oData)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 4: INSERT DESCRIPTION
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescription") Then
        Dim oCtrlDesc As Object
        oCtrlDesc = oModel.getByName("TxtDescription")
        Call ServManageTxtDescription(oCtrlDesc, IdActive, oData)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 5: CALCULATE AND ASSIGN RELATIVE POSITION
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosition") Then
        ' We are sending the model and the view separately
        Call ServCalculateNumPosition(oModel.getByName("NumPosition"), oDialogue.getControl("NumPosition"), IdActive, oData, oModel)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 6 (UNIFIED): POPULATE ITEM DELETION COMBO BOX (CmbItem)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbItem") Then
        Dim sTypeItem As String : sTypeItem = ""
        ' Dynamically detect the context based on the window title.
        If InStr(LCase(oDialogue.Title), "activity") > 0 Then
            sTypeItem = "Activity"
        ElseIf InStr(LCase(oDialogue.Title), "state") > 0 Then
            sTypeItem = "State"
        End If
        
        Call ServPopulateCmbItem(oModel.getByName("CmbItem"), oDialogue.getControl("CmbItem"), oData, sTypeItem)
    End If
    
    ' Reactivate screen update
    oDoc.removeActionLock()
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 1: POPULATE AND SELECT activity (VIEW)
REM =========================================================================
Sub ServPopulateCmbActivity(oModelCombo As Object, oControlView As Object, oData As Object, IdActive As Long)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sValue As String, sActivitySearch As String
    Dim vlist() As String, nCounter As Long : nCounter = 0
    Dim idxActivity As Long : idxActivity = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Load data into the array
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    
    For i = 1 To TotalRows - 1
        sValue = Trim(oData.getCellByPosition(7, i).String)
        If sValue <> "" Then
            vlist(nCounter) = sValue
            nCounter = nCounter + 1
        End If
    Next i
    
    ' 2. Inject the list into the model
    If nCounter > 0 Then
        ReDim Preserve vlist(nCounter - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vEmpty(0) As String
        oModelCombo.StringItemList = vEmpty()
        Exit Sub
    End If
    
    ' 3. Pre-selection via this proposal (.selectItemPos)
    If IdActive > 0 Then
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = IdActive Then
                sActivitySearch = oData.getCellByPosition(0, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sActivitySearch Then
                idxActivity = i
                Exit For
            End If
        Next i
        
        If idxActivity <> -1 Then
            ' Native and foolproof method to force the visual text in a ComboBox's view
            oControlView.setText(sActivitySearch)
        End If
    End If
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 2: POPULATE AND SELECT STATE (VIEW)
REM =========================================================================
Sub ServPopulateCmbState(oModelCombo As Object, oControlView As Object, oData As Object, IdActive As Long)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sValue As String, sStateSearch As String
    Dim vlist() As String, nCounter As Long : nCounter = 0
    Dim idxState As Long : idxState = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Load data into the array
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    
    For i = 1 To TotalRows - 1
        sValue = Trim(oData.getCellByPosition(9, i).String)
        If sValue <> "" And LCase(sValue) <> "Trash" Then
            vlist(nCounter) = sValue
            nCounter = nCounter + 1
        End If
    Next i
    
    ' 2. Inject the list into the model
    If nCounter > 0 Then
        ReDim Preserve vlist(nCounter - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vEmpty(0) As String
        oModelCombo.StringItemList = vEmpty()
        Exit Sub
    End If
    
    ' 3. Pre-selection via your view proposal (.selectItemPos)
    If IdActive > 0 Then
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = IdActive Then
                sStateSearch = oData.getCellByPosition(2, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sStateSearch Then
                idxState = i
                Exit For
            End If
        Next i
        
        If idxState <> -1 Then
            ' Native and foolproof method to force the visual text in a ComboBox's view.
            oControlView.setText(sStateSearch)
        End If
    End If
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 3: MANAGE TASK TITLE (ATOMIC)
REM =========================================================================
Sub ServManageTxtTitle(oModelTitle As Object, IdActive As Long, oData As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    
    ' If the ID is 0, it is a new entry; we clear the title and exit
    If IdActive = 0 Then
        oModelTitle.Text = ""
        Exit Sub
    End If
    
    ' We look for the task row on the data sheet
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        ' Column F (Index 5) contains the task ID.
        If oData.getCellByPosition(5, i).Value = IdActive Then
            ' Inject only the current title (Column B, Index 1)
            oModelTitle.Text = oData.getCellByPosition(1, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 4: GESTIONAR DESCRIPCIÓN DE LA TAREA
REM =========================================================================
Sub ServManageTxtDescription(oModelDesc As Object, IdActive As Long, oData As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    
    ' If the ID is 0, it is a new task; we leave the description blank
    If IdActive = 0 Then
        oModelDesc.Text = ""
        Exit Sub
    End If
    
    ' We are looking for the task row to extract the description
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        ' Column F (Index 5) contains the task ID.
        If oData.getCellByPosition(5, i).Value = IdActive Then
            ' Column D (Index 3) contains the description
            oModelDesc.Text = oData.getCellByPosition(3, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 5: GENERATE VISUAL POSITION MAP (CORRECTED)
REM =========================================================================
Sub ServCalculateNumPosition(oModelCombo As Object, oControlView As Object, IdActive As Long, oData As Object, oModelDialogo As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sStateTask As String, sActivityTask As String
    Dim nTotalInState As Long : nTotalInState = 0
    Dim nPositionCurrent As Long : nPositionCurrent = 0
    Dim vlist() As String
    
    ' 1. Determine the activity and STATE context in which the task operates
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    If IdActive = 0 Then
        ' If it is a new task, we read the dialog options
        If oModelDialogo.hasByName("CmbActivity") Then sActivityTask = Trim(oModelDialogo.getByName("CmbActivity").Text)
        If oModelDialogo.hasByName("CmbState") Then sStateTask = Trim(oModelDialogo.getByName("CmbState").Text)
        
        ' Default values if empty
        If sActivityTask = "" Then sActivityTask = "General"
        If sStateTask = "" Then sStateTask = "To do"
    Else
        ' If the task exists, we extract its actual Activity (Col A) and State (Col C) from the row
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = IdActive Then
                sActivityTask = oData.getCellByPosition(0, i).String ' Col A
                sStateTask    = oData.getCellByPosition(2, i).String ' Col C
                Exit For
            End If
        Next i
    End If
    
    ' 2. Count accurately (filtering by Activity + State and omitting junk rows)
    Dim sEstRow As String, sActRow As String
    
    For i = 1 To TotalRows - 1
        sActRow = Trim(oData.getCellByPosition(0, i).String)
        sEstRow = Trim(oData.getCellByPosition(2, i).String)
        
        ' GOLDEN RULE: Activity and State must match, and must NOT be 'Trash' or empty
        If sActRow = sActivityTask And sEstRow = sStateTask And LCase(sEstRow) <> "Trash" And sActRow <> "" Then
            nTotalInState = nTotalInState + 1
            
            ' If we find the ID we are looking for, we store its exact position within this subset
            If IdActive > 0 Then
                If oData.getCellByPosition(5, i).Value = IdActive Then
                    nPositionCurrent = nTotalInState
                End If
            End If
        End If
    Next i
    
    ' 3. Build the combo box option list based on the actual count
    Dim nLimitOptions As Long
    nLimitOptions = nTotalInState
    
    ' We create the array for the integers (1, 2, 3...) plus the "F" option.
    ReDim vlist(nLimitOptions)
    
    For i = 1 To nLimitOptions
        vlist(i - 1) = CStr(i)
    Next i
    vlist(nLimitOptions) = "F"
    
    ' We inject the clean list of positions
    oModelCombo.StringItemList = vlist()
    
    ' 4. Visual pre-selection using the setText method on the View
    If IdActive = 0 Then
        oControlView.setText("F")
    Else
        If nPositionCurrent > 0 Then
            oControlView.setText(CStr(nPositionCurrent))
        Else
            oControlView.setText("F")
        End If
    End If
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 6: POPULATE UNIFIED ITEM COMBO (VIEW)
REM =========================================================================
Sub ServPopulateCmbItem(oModelCombo As Object, oControlView As Object, oData As Object, sTypeItem As String)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long, nCounter As Long
    Dim sValue As String
    Dim vlist() As String
    
    oModelCombo = oControlView.Model
    
    ' Get limits on the data sheet
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    nCounter = 0
    
    ' Extract texts from the corresponding catalog
    For i = 1 To TotalRows - 1
        If sTypeItem = "Activity" Then
            sValue = Trim(oData.getCellByPosition(7, i).String) ' Column H (index 7)
        Else
            sValue = Trim(oData.getCellByPosition(9, i).String) ' Column J (index 9)
        End If
        
        If sValue <> "" Then
            vlist(nCounter) = sValue
            nCounter = nCounter + 1
        End If
    Next i
    
    ' Inject the catalog into the ComboBox
    If nCounter > 0 Then
        ReDim Preserve vlist(nCounter - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vEmpty(0) As String
        oModelCombo.StringItemList = vEmpty()
    End If
End Sub





'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM STANDALONE CENTRAL WRITING HUB (ATOMIC STRUCTURE)
REM =========================================================================
Sub SaveDataFromDialog(oDialogue As Object, IdTask As Long)
    Dim oDoc As Object, oData As Object
    Dim oModel As Object, oCursor As Object
    Dim nRowDestiny As Long
    Dim TotalRows As Long, i As Long
    Dim NewId As Long
    Dim BinIsNew As Boolean
    
    ' 1. Environment loading
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    oModel = oDialogue.Model
    
    ' Disable screen updating for maximum speed and to prevent flickering
    oDoc.addActionLock()
    
    ' =========================================================================
    ' BLOCK 1: DETERMINATION OF PARAMETERS IN MEMORY (WITHOUT MODIFYING THE DATABASE)
    ' =========================================================================
    If IdTask = 0 Then
        ' --- SCENARIO A: NEW TASK REGISTRATION ---
        BinIsNew = True
        
        oCursor = oData.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        nRowDestiny = oCursor.RangeAddress.EndRow + 1
        
        ' Calculate the new ID available only in memory
        NewId = GetNextAvailableId()
        IdTask = NewId
    Else
        ' --- SCENARIO B: EDITING AN EXISTING TASK ---
        BinIsNew = False
        
        oCursor = oData.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        TotalRows = oCursor.RangeAddress.EndRow + 1
        
        nRowDestiny = -1
        ' Database scan to locate the row by ID (Column F / index 5)
        For i = 1 To TotalRows - 1
            If CLng(oData.getCellByPosition(5, i).Value) = IdTask Then
                nRowDestiny = i
                Exit For
            End If
        Next i
        
        ' Security check if the ID is not found in the database
        If nRowDestiny = -1 Then
            oDoc.removeActionLock()
            MsgBox " Error: Could not locate the task with ID" & IdTask & " .in the database", 48, "Synchronization Error"
            Exit Sub
        End If
    End If
    
    ' =========================================================================
    ' BLOCK 2: DATA PROCESSING AND WRITING TO THE DATABASE
    ' =========================================================================
    
    ' Unique ID Injection (Column F / Index 5)
    ' It is centralized in this block, along with the rest of the physical title deeds
    oData.getCellByPosition(5, nRowDestiny).Value = IdTask
    
    REM ---------------------------------------------------------------------
    REM SERVICE 1: SAVE activity
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbActivity") Then
        ' Extract the text from the combo box and inject it into Column A (0)
        Call ServSaveCmbActivity(oDialogue.getControl("CmbActivity"), oData, nRowDestiny)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 2: SAVE STATE
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbState") Then
        Call ServSaveCmbState(oDialogue.getControl("CmbState"), oData, nRowDestiny)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 3: SAVE TITLE
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitle") Then
        Call ServSaveTxtTitle(oDialogue.getControl("TxtTitle"), oData, nRowDestiny)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 4: SAVE DESCRIPTION
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescription") Then
        Call ServSaveTxtDescription(oDialogue.getControl("TxtDescription"), oData, nRowDestiny)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICE 5: SAVE POSITION / RELATIVE ORDER
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosition") Then
        Call ServSaveNumPosition(oDialogue.getControl("NumPosition"), oData, nRowDestiny, oModel, BinIsNew)
    ElseIf BinIsNew Then
        ' If it is a new task and the dialog does not specify a position control, it is assigned to the end by default.
        Call ServSaveNumPositionByDefault(oData, nRowDestiny, oModel)
    End If
    
    ' 3. Unlock screen
    oDoc.removeActionLock()
    
    ' 4. Automatic maintenance and updates
    ' Physically sort the database and redraw the screen
    CustomDataSort()
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 1: SAVE activity TO THE DATABASE
REM =========================================================================
Sub ServSaveCmbActivity(oControlView As Object, oData As Object, nRowDestiny As Long)
    Dim sActivity As String
    
    ' Extract the clean text directly from the ComboBox view
    sActivity = Trim(oControlView.Text)
    
    ' Write the value in Column A (index 0) of the corresponding row
    oData.getCellByPosition(0, nRowDestiny).String = sActivity
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 2: SAVE STATE TO THE DATABASE
REM =========================================================================
Sub ServSaveCmbState(oControlView As Object, oData As Object, nRowDestiny As Long)
    Dim sState As String
    
    ' Extract the clean text directly from the status ComboBox view
    sState = Trim(oControlView.Text)
    
    ' Write the value in Column C (index 2) of the corresponding row.
    oData.getCellByPosition(2, nRowDestiny).String = sState
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 3: SAVE TITLE TO THE DATABASE
REM =========================================================================
Sub ServSaveTxtTitle(oControlView As Object, oData As Object, nRowDestiny As Long)
    Dim sTitle As String
    
    ' Extract the clean text from the text field (TextField)
    sTitle = Trim(oControlView.Text)
    
    ' Write the value in Column B (index 1) of the corresponding row
    oData.getCellByPosition(1, nRowDestiny).String = sTitle
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 4: SAVE DESCRIPTION TO THE DATABASE
REM =========================================================================
Sub ServSaveTxtDescription(oControlView As Object, oData As Object, nRowDestiny As Long)
    Dim sDescription As String
    
    ' Extract the text from the description box (supports multi-line formatting)
    sDescription = Trim(oControlView.Text)
    
    ' Write the value in Column D (index 3) of the corresponding row.
    oData.getCellByPosition(3, nRowDestiny).String = sDescription
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 5: SAVE POSITION / RELATIVE ORDER IN THE DATABASE
REM =========================================================================
Sub ServSaveNumPosition(oControlView As Object, oData As Object, nRowDestiny As Long, oModelDialogo As Object, BinIsNew As Boolean)
    Dim sActivity As String, sState As String
    Dim sPosText As String
    Dim nCalculatedOrder As Double
    
    ' 1. Read the actual activity and status already physically written in the destination row
    sActivity = oData.getCellByPosition(0, nRowDestiny).String
    sState = oData.getCellByPosition(2, nRowDestiny).String
    
    ' 2. Extract the text from the NumPosition ComboBox view
    sPosText = Trim(oControlView.Text)
    
    ' 3. Evaluate the input and request the weight from the priority engine
    If LCase(sPosText) = "f" Or sPosText = "" Or Not IsNumeric(sPosText) Then
        ' "F" case, empty or invalid input: Place at the absolute end of the group.
        nCalculatedOrder = GetNextOrder(sActivity, sState, 0)
    Else
        ' Numerical case: Interpolate the position in the middle of the task list
        nCalculatedOrder = GetNextOrder(sActivity, sState, CLng(sPosText))
    End If
    
    ' 4. Save the decimal weight in Column E (index 4)
    oData.getCellByPosition(4, nRowDestiny).Value = nCalculatedOrder
End Sub

REM =========================================================================
REM AUXILIARY SERVICE 5: ASSIGN DEFAULT RELATIVE ORDER (AT THE END)
REM =========================================================================
Sub ServSaveNumPositionByDefault(oData As Object, nRowDestiny As Long, oModelDialogo As Object)
    Dim sActivity As String, sState As String
    Dim nCalculatedOrder As Double
    
    ' 1. Read the row's current group context.
    sActivity = oData.getCellByPosition(0, nRowDestiny).String
    sState = oData.getCellByPosition(2, nRowDestiny).String
    
    ' 2. Request an order at the end of the respective list
    nCalculatedOrder = GetNextOrder(sActivity, sState, 0)
    
    ' 3. Write in Column E (index 4)
    oData.getCellByPosition(4, nRowDestiny).Value = nCalculatedOrder
End Sub




'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM DIALOGS AND BUTTONS
REM =========================================================================
REM =========================================================================
REM UNIFIED MACRO: SEND TO TRASH (SUPPORTS DIALOG AND COMMAND)
REM =========================================================================
Sub MoveTaskToTrash(Optional oEvent As Object)
    Dim oDoc As Object, oData As Object, oVisor As Object, oSel As Object
    Dim oDialogue As Object
    Dim IdActive As Long, TotalRows As Long, i As Long
    Dim DataRowIdx As Long : DataRowIdx = -1
    
    ' 1. Determine the context of the call (Does it stem from a dialogue?)
    oDialogue = Nothing
    If Not IsMissing(oEvent) Then
        On Error Resume Next
        ' Attempts to retrieve the containing dialog if it originates from a control
        oDialogue = oEvent.Source.Context
        On Error GoTo 0
    End If
    
    ' 2. Load environment and get the active ID of the focused cell
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oSel = oDoc.CurrentSelection
    
    ' Safety exit if no single cell is selected
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    IdActive = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If IdActive <= 0 Then
        MsgBox "No active task was detected to move to the Trash.", 48, "Attention"
        Exit Sub
    End If
    
    ' 3. Locate the physical row in the database (Column F / index 5)
    Dim oCursor As Object
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        If CLng(oData.getCellByPosition(5, i).Value) = IdActive Then
            DataRowIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Execution of the logical state change
    If DataRowIdx <> -1 Then
        ' Block screen update
        oDoc.addActionLock()
        
        ' Change the State to "Trash" in Column C (index 2)
        oData.getCellByPosition(2, DataRowIdx).String = "Trash"
        
        ' Assign a very high order to remove it from active priorities
        oData.getCellByPosition(4, DataRowIdx).Value = 99999
        
        ' Clean up descriptions in the Viewer
        RunInjectionDescription(oVisor, 0)
        
        ' Reorder the database and redraw the board synchronously
        CustomDataSort()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        ' 5. Dialogue Management: If the object was detected in memory, we close it.
        If Not IsNull(oDialogue) And Not IsEmpty(oDialogue) Then
            oDialogue.endExecute()
        End If
        
        MsgBox "The task has been moved to the Trash..", 64, "Process Completed"
    Else
        MsgBox "Error: The task could not be located in the database.", 48, "Synchronization Error"
    End If
End Sub

REM =========================================================================
REM Synchronization Error
REM =========================================================================
Sub DeleteActiveTask(Optional oEvent As Object)
    Dim oDoc As Object, oData As Object, oVisor As Object, oSel As Object
    Dim oDialogue As Object
    Dim IdActive As Long, TotalRows As Long, i As Long
    Dim DataRowIdx As Long : DataRowIdx = -1
    Dim answer As Integer
    Dim sDeleteTaskName As String, sDeleteActivity As String
    
    ' 1. Determine the context of the call (Does it stem from a dialogue?)
    oDialogue = Nothing
    If Not IsMissing(oEvent) Then
        On Error Resume Next
        ' Attempts to retrieve the containing dialog if it originates from a control
        oDialogue = oEvent.Source.Context
        On Error GoTo 0
    End If
    
    ' 2. Load environment and get the active ID of the focused cell
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oSel = oDoc.CurrentSelection
    
    ' Emergency exit
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    IdActive = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If IdActive <= 0 Then
        MsgBox "No active task to delete was detected.", 48, "Attention"
        Exit Sub
    End If
    
    ' 3. Locate the physical row in the database (Column F / index 5)
    Dim oCursor As Object
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        If CLng(oData.getCellByPosition(5, i).Value) = IdActive Then
            DataRowIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Critical confirmation and physical elimination
    If DataRowIdx <> -1 Then
        sDeleteActivity = oData.getCellByPosition(0, DataRowIdx).String
        sDeleteTaskName = oData.getCellByPosition(1, DataRowIdx).String
        
        ' Critical confirmatory clinical picture
        answer = MsgBox("⚠️ ¿Are you sure you want to permanently delete this task?? ⚠️" & Chr(13) & Chr(13) & _
                           "📌 Activity: " & sDeleteActivity & Chr(13) & _
                           "📝 Task: """ & sDeleteTaskName & """" & Chr(13) & _
                           "🆔 Unique ID: " & IdActive & Chr(13) & Chr(13) & _
                           "This action will completely remove the row from the database. It is irreversible.", 4 + 16 + 256, "Confirm Permanent Deletion")
                           
        If answer <> 6 Then Exit Sub ' If cancel, we're going out.
        
        ' Block screen update
        oDoc.addActionLock()
        
        ' Delete row from the database
        oData.Rows.removeByIndex(DataRowIdx, 1)
        
        ' Clear residual descriptions in the viewer (Cell D4)
        RunInjectionDescription(oVisor, 0)
        
        ' Sequentially reorder and redraw the Kanban board
        CustomDataSort()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        ' 5. Dialogue Management: If the object was detected in memory, we close it
        If Not IsNull(oDialogue) And Not IsEmpty(oDialogue) Then
            oDialogue.endExecute()
        End If
        
        MsgBox "The task has been successfully deleted.", 64, "Process Completed"
    Else
        MsgBox "Error: The task could not be located in the database.", 48, "Synchronization Error"
    End If
End Sub



REM =========================================================================
REM SUMMONERS
REM =========================================================================

'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM EDITING MACRO: EDIT SELECTED TASK (USING UNIFIED DIALOG)
REM =========================================================================
Sub EditActiveTask()
    Dim oDoc As Object, oVisor As Object, oSel As Object
    Dim oDialogue As Object
    Dim IdActive As Long
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oSel = oDoc.CurrentSelection
    
    ' 1. Detect the ID of the task selected in the viewer
    IdActive = 0
    If oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        IdActive = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    End If
    
    ' If the selected cell is empty or a header, we exit immediately
    If IdActive <= 0 Then
        MsgBox "Please select a valid task on the board first.", 48, "Attention"
        Exit Sub
    End If
    
    ' 2. Load and instantiate the unified "DialogPanelTask" dialog.
    DialogLibraries.LoadLibrary("Standard")
    oDialogue = CreateUnoDialog(DialogLibraries.Standard.DialogPanelTask)
    
    ' Freeze listener for visual stability
    binFrozenFocus = True
    
    ' 3. Invoke the Reading Hub with the selected ID (This will display BtnDelete and BtnTrash)
    Call InitializeDialogServices(oDialogue, IdActive)
    
    ' 4. Execution of the dialogue
    If oDialogue.execute() = 1 Then
        ' Save the changes to the existing record.
        Call SaveDataFromDialog(oDialogue, IdActive)
    End If
    
    ' Safely release resources
    binFrozenFocus = False
    oDialogue.dispose()
End Sub

REM =========================================================================
REM CREATION MACRO: REGISTER NEW TASK (USING UNIFIED DIALOG)
REM =========================================================================
Sub AddNewTask()
    Dim oDoc As Object, oVisor As Object, oData As Object
    Dim oDialogue As Object
    Dim sActivityFilter As String, sStateFilter As String
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Load and instantiate the unified "DialogPanelTask" dialog.
    DialogLibraries.LoadLibrary("Standard")
    oDialogue = CreateUnoDialog(DialogLibraries.Standard.DialogPanelTask)
    
    ' Temporarily freeze the listener to prevent visual flickering
    binFrozenFocus = True
    
    ' 2. Invoke the Reading Hub with ID = 0 (This will automatically hide BtnDelete and BtnTrash)
    Call InitializeDialogServices(oDialogue, 0)
    
    ' 3. Smart pre-selection based on the Viewer's active filters
    sActivityFilter = Trim(oVisor.getCellByPosition(0, 1).String) ' Cell A2
    sStateFilter = Trim(oVisor.getCellByPosition(0, 2).String)    ' Cell A3
    
    If sActivityFilter <> "" Then
        oDialogue.getControl("CmbActivity").setText(sActivityFilter)
    End If
    
    If sStateFilter <> "" Then
        oDialogue.getControl("CmbState").setText(sStateFilter)
    Else
        oDialogue.getControl("CmbState").setText("In progress")
    End If
    
    ' 4. Execution of the dialogue
    If oDialogue.execute() = 1 Then
        ' Validation of mandatory fields
        If Trim(oDialogue.getControl("TxtTitle").Text) = "" Then
            MsgBox "The 'Activity' and the 'Task' name are mandatory..", 48, "Incomplete Fields"
        Else
            ' Save the new task to the database (ID = 0)
            Call SaveDataFromDialog(oDialogue, 0)
        End If
    End If
    
    ' Safely release environment
    binFrozenFocus = False
    oDialogue.dispose()
End Sub

REM =========================================================================
REM ENTRY POINT: MANAGE PROJECTS / ACTIVITIES
REM =========================================================================
Sub ManageActivityes()
    G_ManagedItemType = "Activity"
    Call OpenItemPanel()
End Sub

REM =========================================================================
REM ENTRY POINT: MANAGE COLUMNS / STATUSES
REM =========================================================================
Sub ManageStates()
    G_ManagedItemType = "State"
    Call OpenItemPanel()
End Sub

REM =========================================================================
REM UNIFIED Summoner: OPEN ITEM PANEL
REM =========================================================================
Sub OpenItemPanel()
    Dim oDialogue As Object
    
    ' 1. Instantiate the unified dialog "DialogPanelItem"
    DialogLibraries.LoadLibrary("Standard")
    oDialogue = CreateUnoDialog(DialogLibraries.Standard.DialogPanelItem)
    
    ' Configure the title dynamically based on the global context
    oDialogue.Title = " Management of" & G_ManagedItemType & "is"
    
    ' 2. Call the Reading Hub (ID = 0) to populate the "CmbItem" ComboBox.
    Call InitializeDialogServices(oDialogue, 0)
    
    ' Initially disable the Accept button until the user types or selects something.
    oDialogue.getControl("BtnAccept").Model.Enabled = False
    
    ' 3. Expand on-screen panel
    oDialogue.execute()
    oDialogue.dispose()
End Sub

' --------------------------------------------------------------------------

REM =========================================================================
REM INTERACTIVE: DYNAMIC BUTTON CHANGE AT RUNTIME
REM =========================================================================
Sub CmbItemTextModified(oEvent As Object)
    Dim oDialogue As Object, oBtnModel As Object, oCmbItem As Object
    Dim sText As String
    Dim vlist() As String
    Dim BinExist As Boolean
    
    ' Obtain references
    oCmbItem = oEvent.Source
    oDialogue = oCmbItem.Context
    oBtnModel = oDialogue.getControl("BtnAccept").Model
    
    sText = Trim(oCmbItem.Text)
    
    ' 1. Empty State: Disable the button and reset it to neutral.
    If sText = "" Then
        oBtnModel.Enabled = False
        oBtnModel.Label = "Accept"
        oBtnModel.BackgroundColor = RGB(240, 240, 240)
        oBtnModel.TextColor = RGB(0, 0, 0)
        Exit Sub
    End If
    
    ' 2. Check if the text already exists in the combo box list.
    vlist() = oCmbItem.Model.StringItemList
    BinExist = ElementExistInList(sText, vlist())
    
    ' 3. Apply dynamic visual mutation to the button
    If BinExist Then
        ' --- DELETION MODE (Available in the list) ---
        oBtnModel.Enabled = True
        oBtnModel.Label = "⮿ Delete"
        oBtnModel.BackgroundColor = RGB(220, 53, 69)
        oBtnModel.TextColor = RGB(255, 255, 255)
    Else
        ' --- ADDITION MODE (New item) ---
        oBtnModel.Enabled = True
        oBtnModel.Label = "[+] Add"
        oBtnModel.BackgroundColor = RGB(40, 167, 69)
        oBtnModel.TextColor = RGB(255, 255, 255)
    End If
End Sub

REM =========================================================================
REM CATALOG WRITING HUB: ADD OR REMOVE BASED ON CONTEXT (MUTATED)
REM =========================================================================
Sub SaveCatalogFromDialog(oEvent As Object)
    Dim oDialogue As Object, oData As Object, oCmbItem As Object
    Dim sText As String
    Dim vlist() As String
    Dim BinExist As Boolean
    Dim answer As Integer
    
    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    oCmbItem = oDialogue.getControl("CmbItem")
    sText = Trim(oCmbItem.Text)
    
    If sText = "" Then Exit Sub
    
    vlist() = oCmbItem.Model.StringItemList
    BinExist = ElementExistInList(sText, vlist())
    
    If BinExist Then
        ' OPERATION: CASCADE ELIMINATION
        oDialogue.endExecute()
        Call RunCascadeDelete(sText, G_ManagedItemType)
    Else
        ' OPERATION: ADD NEW ELEMENT
        answer = MsgBox("Do you want to add the new one? " & G_ManagedItemType & ": """ & sText & """?", 4 + 32, "Confirm Addition")
        
        If answer = 6 Then
            Dim mData() As Variant, i As Long, EmptyRow As Long
            Dim nColIdx As Long
            
            EmptyRow = -1
            If G_ManagedItemType = "Activity" Then
                nColIdx = 7
                mData = oData.getCellRangeByName("H2:H100").getDataArray()
            Else
                nColIdx = 9
                mData = oData.getCellRangeByName("J2:J20").getDataArray()
            End If
            
            For i = 0 To UBound(mData)
                If Trim(mData(i)(0)) = "" Then
                    EmptyRow = i + 2
                    Exit For
                End If
            Next i
            
            If EmptyRow <> -1 Then
                oData.getCellByPosition(nColIdx, EmptyRow - 1).String = sText
                oDialogue.endExecute()
                
                CustomDataSort()
                MsgBox G_ManagedItemType & " successfully added to the catalog.", 64, "Success"
            Else
                MsgBox "Error: The catalog of " & G_ManagedItemType & "It's full. ", 48, "Limit Exceeded"
            End If
        End If
    End If
End Sub

REM =========================================================================
REM AUXILIARY FUNCTION: CASE-INSENSITIVE COMPARISON IN ARRAYS
REM =========================================================================
Function ElementExistInList(sSearch As String, ByRef vlist() As String) As Boolean
    Dim i As Long
    sSearch = LCase(Trim(sSearch))
    For i = 0 To UBound(vlist)
        If LCase(Trim(vlist(i))) = sSearch Then
            ElementExistInList = True
            Exit Function
        End If
    Next i
    ElementExistInList = False
End Function
