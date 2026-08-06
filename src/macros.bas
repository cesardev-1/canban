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
Sub StartListenerSeleccion()
    Dim oDoc As Object
    oDoc = ThisComponent
    oSheetHeard = oDoc.Sheets.getByName("view")
    
    ' Create the selection event listener object
    oSeleccionListener = CreateUnoListener("VisorSeleccion_", "com.sun.star.view.XSelectionChangeListener")
    oDoc.CurrentController.addSelectionChangeListener(oSeleccionListener)
End Sub

Sub RemoveListenerSeleccion()
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
    ' - A2 (Fila 1, Col 0) o C2 (Fila 1, Col 2) -> Active Activity Filter
    ' - A3 (Fila 2, Col 0) -> Left State Filter (Column A)
    ' - B3 (Fila 2, Col 1) -> Right-of-State Filter (Column B)
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
    Dim idActual As Long, NewId As Long
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
        idActual = oData.getCellByPosition(5, i).Value          ' Col F
        
        ' VALIDATION RULE:
        ' If Activity, Task, and Status are present (fields populated) AND the ID is 0 or empty.
        If sActivity <> "" And sTask <> "" And sState <> "" And idActual = 0 Then
            
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
        Dim mHistorico() As Variant, oRangoTareas As Object
        oRangoTareas = oData.getCellRangeByPosition(0, 1, 4, CurrentTotalRows - 1)
        mHistorico = oRangoTareas.getDataArray()
        
        ' Filtramos para ubicar los índices globales de las tareas de este grupo específico
        Dim mIndicesGlobales() As Long, cFiltrados As Long : cFiltrados = 0
        For i = 0 To UBound(mHistorico)
            If Trim(mHistorico(i)(0)) = sActivity And Trim(mHistorico(i)(2)) = sState Then
                ReDim Preserve mIndicesGlobales(cFiltrados)
                mIndicesGlobales(cFiltrados) = i ' Guardamos la posición física de la fila
                cFiltrados = cFiltrados + 1
            End If
        Next i
        
        If cFiltrados > 0 Then
            ' Determinar los pesos globales en la columna E basados en la posición deseada
            If nPosWanted <= 1 Then
                ' Al principio del grupo: Peso de la primera tarea del grupo menos 0.1
                GetNextOrder = CDbl(mHistorico(mIndicesGlobales(0))(4)) - 0.1
            ElseIf nPosWanted > cFiltrados Then
                ' Al final del grupo: Peso de la última tarea del grupo más 0.1
                GetNextOrder = CDbl(mHistorico(mIndicesGlobales(cFiltrados - 1))(4)) + 0.1
            Else
                ' En medio: Punto medio decimal entre la tarea elegida y la anterior
                Dim ordenAnterior As Double, ordenSiguiente As Double
                ordenAnterior = CDbl(mHistorico(mIndicesGlobales(nPosWanted - 2))(4))
                ordenSiguiente = CDbl(mHistorico(mIndicesGlobales(nPosWanted - 1))(4))
                GetNextOrder = ordenAnterior + ((ordenSiguiente - ordenAnterior) / 2)
            End If
            Exit Function
        End If
    End If
    
    ' =========================================================================
    ' ESCENARIO B: CONSECUTIVO AL FINAL ABSOLUTO DE LA TABLA
    ' =========================================================================
    ' Si la tabla está vacía (solo cabecera), inicia en 1, sino va al final de la list corrida
    If CurrentTotalRows <= 2 Then
        GetNextOrder = 1.0
    Else
        GetNextOrder = CDbl(CurrentTotalRows)
    End If
End Function

REM =========================================================================
REM COMPONENTE DEFENSIVO: REPARA Y ASIGNA NÚMEROS DE ORDEN FALTANTES (F2)
REM =========================================================================
Sub AssignMissingOrdersBD()
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRowsData As Long, i As Long
    Dim sActivity As String, sState As String
    Dim oCeldaOrden As Object
    Dim nOrdenAsignado As Double
    
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Encontrar el límite real de filas en la hoja datos
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRowsData = oCursor.RangeAddress.EndRow + 1
    
    ' Si sólo están las cabeceras, no hay nada que auditar
    If TotalRowsData <= 2 Then Exit Sub
    
    ' 2. Barrido fila por fila analizando la columna E (índice 4)
    For i = 1 To TotalRowsData - 1
        sActivity = Trim(oData.getCellByPosition(0, i).String) ' Col A
        sState = Trim(oData.getCellByPosition(2, i).String)    ' Col C
        oCeldaOrden = oData.getCellByPosition(4, i)             ' Col E
        
        ' Una fila es válida si al menos tiene Actividad y Nombre de Tarea
        If sActivity <> "" And Trim(oData.getCellByPosition(1, i).String) <> "" Then
            
            ' Detectar si el orden es nulo, vacío o cero
            If oCeldaOrden.Type = com.sun.star.table.CellContentType.EMPTY Or oCeldaOrden.Value = 0 Then
                
                ' Solicitamos al componente experto el siguiente número entero para este grupo
                ' Nota: No pasamos el tercer parámetro porque queremos que vaya al final de su list
                nOrdenAsignado = GetNextOrder(sActivity, sState)
                
                ' Inyectamos el valor curado en la base de datos
                oCeldaOrden.Value = nOrdenAsignado
                
            End If
        End If
    Next i
End Sub

REM =========================================================================
REM HUB CENTRAL DE SERVICIOS AUTÓNOMO PARA DIÁLOGOS (ESTRUCTURA ATÓMICA)
REM =========================================================================
Sub InicializarServiciosDialogo(oDialogo As Object, idActivo As Long)
    Dim oDoc As Object, oVisor As Object, oData As Object
    Dim oModel As Object, oSel As Object
    
    ' 1. CARGA DE ENTORNO GLOBAL (Siempre se usará)
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oModel = oDialogo.Model
    
    REM ---------------------------------------------------------------------
    REM SERVICIO DE VISIBILIDAD DINÁMICA: MOSTRAR/OCULTAR BOTONES DE ACCIÓN
    REM ---------------------------------------------------------------------
    ' Si es alta nueva (idActivo = 0), ocultamos los botones de destrucción física y lógica
    If oModel.hasByName("BtnEliminar") Then
        oDialogo.getControl("BtnEliminar").setVisible(idActivo > 0)
    End If
    
    If oModel.hasByName("BtnPapelera") Then
        oDialogo.getControl("BtnPapelera").setVisible(idActivo > 0)
    End If
    
    ' Desactivar actualización de pantalla por rendimiento general
    oDoc.addActionLock()
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 1: POBLAR COMBO DE ACTIVIDADES
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbActividad") Then
        ' Pasamos el modelo y la vista por separado a la macro externa
        Call ServPoblarCmbActividad(oModel.getByName("CmbActividad"), oDialogo.getControl("CmbActividad"), oData, idActivo)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 2: POBLAR COMBO DE ESTADOS
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbEstado") Then
        ' Pasamos el modelo y la vista por separado a la macro externa
        Call ServPoblarCmbEstado(oModel.getByName("CmbEstado"), oDialogo.getControl("CmbEstado"), oData, idActivo)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 3: INYECTAR TÍTULO (Contextual al ID Activo o Selección)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitulo") Then
        Dim oCtrlTitulo As Object
        oCtrlTitulo = oModel.getByName("TxtTitulo")
        ' Limpio y atómico: Solo se encarga del texto del título
        Call ServGestionarTxtTitulo(oCtrlTitulo, idActivo, oData)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 4: INYECTAR DESCRIPCIÓN
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescripcion") Then
        Dim oCtrlDesc As Object
        oCtrlDesc = oModel.getByName("TxtDescripcion")
        Call ServGestionarTxtDescripcion(oCtrlDesc, idActivo, oData)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 5: CALCULAR Y ASIGNAR POSICIÓN RELATIVA
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosicion") Then
        ' Enviamos modelo y vista por separado
        Call ServCalcularNumPosicion(oModel.getByName("NumPosicion"), oDialogo.getControl("NumPosicion"), idActivo, oData, oModel)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 6 (UNIFICADO): POBLAR COMBO DE BORRADO DE ITEMS (CmbItem)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbItem") Then
        Dim sTypeItem As String : sTypeItem = ""
        ' Detectar dinámicamente el contexto por el título de la ventana
        If InStr(LCase(oDialogo.Title), "actividad") > 0 Then
            sTypeItem = "Activity"
        ElseIf InStr(LCase(oDialogo.Title), "estado") > 0 Then
            sTypeItem = "State"
        End If
        
        Call ServPoblarCmbItem(oModel.getByName("CmbItem"), oDialogo.getControl("CmbItem"), oData, sTypeItem)
    End If
    
    ' Reactivar actualización de pantalla
    oDoc.removeActionLock()
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 1: POBLAR Y SELECCIONAR ACTIVIDAD (VISTA)
REM =========================================================================
Sub ServPoblarCmbActividad(oModelCombo As Object, oControlView As Object, oData As Object, idActivo As Long)
    Dim  oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sValue As String, sActivityBuscar As String
    Dim vlist() As String, nContador As Long : nContador = 0
    Dim idxActividad As Long : idxActividad = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Cargar datos en el Array
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    
    For i = 1 To TotalRows - 1
        sValue = Trim(oData.getCellByPosition(7, i).String)
        If sValue <> "" Then
            vlist(nContador) = sValue
            nContador = nContador + 1
        End If
    Next i
    
    ' 2. Inyectar la list al modelo
    If nContador > 0 Then
        ReDim Preserve vlist(nContador - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
        Exit Sub
    End If
    
    ' 3. PRE-SELECCIÓN MEDIANTE TU PROPUESTA DE VISTA (.selectItemPos)
    If idActivo > 0 Then
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = idActivo Then
                sActivityBuscar = oData.getCellByPosition(0, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sActivityBuscar Then
                idxActividad = i
                Exit For
            End If
        Next i
        
        If idxActividad <> -1 Then
            ' Método nativo e infalible para forzar el texto visual en la Vista de un ComboBox
            oControlView.setText(sActivityBuscar)
        End If
    End If
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 2: POBLAR Y SELECCIONAR ESTADO (VISTA)
REM =========================================================================
Sub ServPoblarCmbEstado(oModelCombo As Object, oControlView As Object, oData As Object, idActivo As Long)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sValue As String, sStateBuscar As String
    Dim vlist() As String, nContador As Long : nContador = 0
    Dim idxEstado As Long : idxEstado = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Cargar datos en el Array
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    
    For i = 1 To TotalRows - 1
        sValue = Trim(oData.getCellByPosition(9, i).String)
        If sValue <> "" And LCase(sValue) <> "Papelera" Then
            vlist(nContador) = sValue
            nContador = nContador + 1
        End If
    Next i
    
    ' 2. Inyectar la list al modelo
    If nContador > 0 Then
        ReDim Preserve vlist(nContador - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
        Exit Sub
    End If
    
    ' 3. PRE-SELECCIÓN MEDIANTE TU PROPUESTA DE VISTA (.selectItemPos)
    If idActivo > 0 Then
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = idActivo Then
                sStateBuscar = oData.getCellByPosition(2, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sStateBuscar Then
                idxEstado = i
                Exit For
            End If
        Next i
        
        If idxEstado <> -1 Then
            ' Método nativo e infalible para forzar el texto visual en la Vista de un ComboBox
            oControlView.setText(sStateBuscar)
        End If
    End If
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 3: GESTIONAR TÍTULO DE LA TAREA (ATÓMICO)
REM =========================================================================
Sub ServGestionarTxtTitulo(oModelTitulo As Object, idActivo As Long, oData As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    
    ' Si el ID es 0, es un alta nueva; limpiamos el título y salimos
    If idActivo = 0 Then
        oModelTitulo.Text = ""
        Exit Sub
    End If
    
    ' Buscamos la fila de la tarea en la hoja datos
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        ' Columna F (Índice 5) contiene el ID de la tarea
        If oData.getCellByPosition(5, i).Value = idActivo Then
            ' Inyectar únicamente el título actual (Columna B, Índice 1)
            oModelTitulo.Text = oData.getCellByPosition(1, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 4: GESTIONAR DESCRIPCIÓN DE LA TAREA
REM =========================================================================
Sub ServGestionarTxtDescripcion(oModelDesc As Object, idActivo As Long, oData As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    
    ' Si el ID es 0, es una tarea nueva; dejamos la descripción limpia
    If idActivo = 0 Then
        oModelDesc.Text = ""
        Exit Sub
    End If
    
    ' Buscamos la fila de la tarea para extraer la descripción
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        ' Columna F (Índice 5) contiene el ID de la tarea
        If oData.getCellByPosition(5, i).Value = idActivo Then
            ' Columna D (Índice 3) contiene la Descripción
            oModelDesc.Text = oData.getCellByPosition(3, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 5: GENERAR MAPA DE POSICIONES VISUALES (CORREGIDO)
REM =========================================================================
Sub ServCalcularNumPosicion(oModelCombo As Object, oControlView As Object, idActivo As Long, oData As Object, oModelDialogo As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sStateTarea As String, sActivityTarea As String
    Dim nTotalEnEstado As Long : nTotalEnEstado = 0
    Dim nPosicionActual As Long : nPosicionActual = 0
    Dim vlist() As String
    
    ' 1. Determinar el contexto de ACTIVIDAD y ESTADO en el que opera la tarea
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    If idActivo = 0 Then
        ' Si es nueva tarea, leemos los combos del diálogo
        If oModelDialogo.hasByName("CmbActividad") Then sActivityTarea = Trim(oModelDialogo.getByName("CmbActividad").Text)
        If oModelDialogo.hasByName("CmbEstado") Then sStateTarea = Trim(oModelDialogo.getByName("CmbEstado").Text)
        
        ' Valuees por defecto si están vacíos
        If sActivityTarea = "" Then sActivityTarea = "General"
        If sStateTarea = "" Then sStateTarea = "To do"
    Else
        ' Si la tarea existe, extraemos su Actividad (Col A) y Estado (Col C) reales de la fila
        For i = 1 To TotalRows - 1
            If oData.getCellByPosition(5, i).Value = idActivo Then
                sActivityTarea = oData.getCellByPosition(0, i).String ' Col A
                sStateTarea    = oData.getCellByPosition(2, i).String ' Col C
                Exit For
            End If
        Next i
    End If
    
    ' 2. Contar de forma precisa (Filtrando por Actividad + Estado y omitiendo filas basura)
    Dim sEstFila As String, sActFila As String
    
    For i = 1 To TotalRows - 1
        sActFila = Trim(oData.getCellByPosition(0, i).String)
        sEstFila = Trim(oData.getCellByPosition(2, i).String)
        
        ' REGLA DE ORO: Deben coincidir Actividad Y Estado, y NO ser papelera ni estar vacíos
        If sActFila = sActivityTarea And sEstFila = sStateTarea And LCase(sEstFila) <> "Papelera" And sActFila <> "" Then
            nTotalEnEstado = nTotalEnEstado + 1
            
            ' Si encontramos el ID que estamos buscando, guardamos su posición exacta en este subconjunto
            If idActivo > 0 Then
                If oData.getCellByPosition(5, i).Value = idActivo Then
                    nPosicionActual = nTotalEnEstado
                End If
            End If
        End If
    Next i
    
    ' 3. Construir la list de opciones del Combo basadas en el conteo real
    Dim nLimiteOpciones As Long
    nLimiteOpciones = nTotalEnEstado
    
    ' Creamos el array para los enteros (1, 2, 3...) más la opción "F"
    ReDim vlist(nLimiteOpciones)
    
    For i = 1 To nLimiteOpciones
        vlist(i - 1) = CStr(i)
    Next i
    vlist(nLimiteOpciones) = "F"
    
    ' Inyectamos la list limpia de posiciones
    oModelCombo.StringItemList = vlist()
    
    ' 4. Preselección visual con el método setText sobre la Vista
    If idActivo = 0 Then
        oControlView.setText("F")
    Else
        If nPosicionActual > 0 Then
            oControlView.setText(CStr(nPosicionActual))
        Else
            oControlView.setText("F")
        End If
    End If
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 6: POBLAR COMBO UNIFICADO DE ITEMS (VISTA)
REM =========================================================================
Sub ServPoblarCmbItem(oModelCombo As Object, oControlView As Object, oData As Object, sTypeItem As String)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long, nContador As Long
    Dim sValue As String
    Dim vlist() As String
    
    oModelCombo = oControlView.Model
    
    ' Obtener límites en la hoja datos
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    nContador = 0
    
    ' Extraer textos del catálogo correspondiente
    For i = 1 To TotalRows - 1
        If sTypeItem = "Activity" Then
            sValue = Trim(oData.getCellByPosition(7, i).String) ' Columna H (índice 7)
        Else
            sValue = Trim(oData.getCellByPosition(9, i).String) ' Columna J (índice 9)
        End If
        
        If sValue <> "" Then
            vlist(nContador) = sValue
            nContador = nContador + 1
        End If
    Next i
    
    ' Inyectar el catálogo en el ComboBox
    If nContador > 0 Then
        ReDim Preserve vlist(nContador - 1)
        oModelCombo.StringItemList = vlist()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
    End If
End Sub





'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM HUB CENTRAL DE ESCRITURA AUTÓNOMO (ESTRUCTURA ATÓMICA)
REM =========================================================================
Sub GuardarDatosDesdeDialogo(oDialogo As Object, IdTask As Long)
    Dim oDoc As Object, oData As Object
    Dim oModel As Object, oCursor As Object
    Dim nRowDestino As Long
    Dim TotalRows As Long, i As Long
    Dim NewId As Long
    Dim binEsNuevo As Boolean
    
    ' 1. Carga de entorno
    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    oModel = oDialogo.Model
    
    ' Desactivar actualización de pantalla para máxima velocidad y evitar parpadeos
    oDoc.addActionLock()
    
    ' =========================================================================
    ' BLOQUE 1: DETERMINACIÓN DE PARÁMETROS EN MEMORIA (SIN MODIFICAR LA BD)
    ' =========================================================================
    If IdTask = 0 Then
        ' --- ESCENARIO A: REGISTRO DE NUEVA TAREA ---
        binEsNuevo = True
        
        oCursor = oData.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        nRowDestino = oCursor.RangeAddress.EndRow + 1
        
        ' Calcular el nuevo ID disponible únicamente en memoria
        NewId = GetNextAvailableId()
        IdTask = NewId
    Else
        ' --- ESCENARIO B: EDICIÓN DE TAREA EXISTENTE ---
        binEsNuevo = False
        
        oCursor = oData.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        TotalRows = oCursor.RangeAddress.EndRow + 1
        
        nRowDestino = -1
        ' Escaneo de base de datos para localizar la fila por ID (Columna F / índice 5)
        For i = 1 To TotalRows - 1
            If CLng(oData.getCellByPosition(5, i).Value) = IdTask Then
                nRowDestino = i
                Exit For
            End If
        Next i
        
        ' Control de seguridad si el ID no se localiza en la BD
        If nRowDestino = -1 Then
            oDoc.removeActionLock()
            MsgBox "Error: No se pudo localizar la tarea con ID " & IdTask & " en la base de datos.", 48, "Error de Sincronización"
            Exit Sub
        End If
    End If
    
    ' =========================================================================
    ' BLOQUE 2: PROCESAMIENTO Y ESCRITURA DE DATOS EN LA BD
    ' =========================================================================
    
    ' Inyección del ID Único (Columna F / índice 5)
    ' Se centraliza en este bloque junto con el resto de escrituras físicas
    oData.getCellByPosition(5, nRowDestino).Value = IdTask
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 1: GUARDAR ACTIVIDAD
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbActividad") Then
        ' Extrae el texto del combo e inyecta en Col A (0)
        Call ServGuardarCmbActividad(oDialogo.getControl("CmbActividad"), oData, nRowDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 2: GUARDAR ESTADO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbEstado") Then
        ' [Pendiente de implementar]: Extrae el texto del combo e inyecta en Col C (2)
        Call ServGuardarCmbEstado(oDialogo.getControl("CmbEstado"), oData, nRowDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 3: GUARDAR TÍTULO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitulo") Then
        ' [Pendiente de implementar]: Extrae el texto del título e inyecta en Col B (1)
        Call ServGuardarTxtTitulo(oDialogo.getControl("TxtTitulo"), oData, nRowDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 4: GUARDAR DESCRIPCIÓN
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescripcion") Then
        ' [Pendiente de implementar]: Extrae la descripción e inyecta en Col D (3)
        Call ServGuardarTxtDescripcion(oDialogo.getControl("TxtDescripcion"), oData, nRowDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 5: GUARDAR POSICIÓN / ORDEN RELATIVO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosicion") Then
        ' [Pendiente de implementar]: Procesa la posición ("F" o valor numérico) y calcula el peso en Col E (4)
        Call ServGuardarNumPosicion(oDialogo.getControl("NumPosicion"), oData, nRowDestino, oModel, binEsNuevo)
    ElseIf binEsNuevo Then
        ' Si es una nueva tarea y el diálogo no expone control de posición, se asigna al final por defecto
        Call ServGuardarNumPosicionPorDefecto(oData, nRowDestino, oModel)
    End If
    
    ' 3. Liberar bloqueo de pantalla
    oDoc.removeActionLock()
    
    ' 4. Mantenimiento y actualización automática
    ' Ordena la base de datos físicamente y redibuja la pantalla
    CustomDataSort()
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 1: GUARDAR ACTIVIDAD EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarCmbActividad(oControlView As Object, oData As Object, nRowDestino As Long)
    Dim sActivity As String
    
    ' Extraer el texto limpio directamente de la vista del ComboBox
    sActivity = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna A (índice 0) de la fila correspondiente
    oData.getCellByPosition(0, nRowDestino).String = sActivity
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 2: GUARDAR ESTADO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarCmbEstado(oControlView As Object, oData As Object, nRowDestino As Long)
    Dim sState As String
    
    ' Extraer el texto limpio directamente de la vista del ComboBox de estados
    sState = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna C (índice 2) de la fila correspondiente
    oData.getCellByPosition(2, nRowDestino).String = sState
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 3: GUARDAR TÍTULO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarTxtTitulo(oControlView As Object, oData As Object, nRowDestino As Long)
    Dim sTitle As String
    
    ' Extraer el texto limpio del cuadro de texto (TextField)
    sTitle = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna B (índice 1) de la fila correspondiente
    oData.getCellByPosition(1, nRowDestino).String = sTitle
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 4: GUARDAR DESCRIPCIÓN EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarTxtDescripcion(oControlView As Object, oData As Object, nRowDestino As Long)
    Dim sDescription As String
    
    ' Extraer el texto del cuadro de descripción (admite formato multilínea)
    sDescription = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna D (índice 3) de la fila correspondiente
    oData.getCellByPosition(3, nRowDestino).String = sDescription
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 5: GUARDAR POSICIÓN / ORDEN RELATIVO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarNumPosicion(oControlView As Object, oData As Object, nRowDestino As Long, oModelDialogo As Object, binEsNuevo As Boolean)
    Dim sActivity As String, sState As String
    Dim sPosText As String
    Dim nOrdenCalculado As Double
    
    ' 1. Leer la actividad y el estado reales ya escritos físicamente en la fila destino
    sActivity = oData.getCellByPosition(0, nRowDestino).String
    sState = oData.getCellByPosition(2, nRowDestino).String
    
    ' 2. Extraer el texto de la vista del ComboBox NumPosicion
    sPosText = Trim(oControlView.Text)
    
    ' 3. Evaluar la entrada y solicitar el peso al motor de prioridades
    If LCase(sPosText) = "f" Or sPosText = "" Or Not IsNumeric(sPosText) Then
        ' Caso "F", vacío o entrada inválida: Colocar al final absoluto del grupo
        nOrdenCalculado = GetNextOrder(sActivity, sState, 0)
    Else
        ' Caso numérico: Interpolar la posición en medio de la list de tareas
        nOrdenCalculado = GetNextOrder(sActivity, sState, CLng(sPosText))
    End If
    
    ' 4. Guardar el peso decimal en la Columna E (índice 4)
    oData.getCellByPosition(4, nRowDestino).Value = nOrdenCalculado
End Sub

REM =========================================================================
REM SERVICIO AUXILIAR 5: ASIGNAR ORDEN RELATIVO POR DEFECTO (AL FINAL)
REM =========================================================================
Sub ServGuardarNumPosicionPorDefecto(oData As Object, nRowDestino As Long, oModelDialogo As Object)
    Dim sActivity As String, sState As String
    Dim nOrdenCalculado As Double
    
    ' 1. Leer el contexto de grupo actual de la fila
    sActivity = oData.getCellByPosition(0, nRowDestino).String
    sState = oData.getCellByPosition(2, nRowDestino).String
    
    ' 2. Solicitar un orden al final de su respectiva list
    nOrdenCalculado = GetNextOrder(sActivity, sState, 0)
    
    ' 3. Escribir en la Columna E (índice 4)
    oData.getCellByPosition(4, nRowDestino).Value = nOrdenCalculado
End Sub




'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM DIALOGOS Y BOTONES
REM =========================================================================
REM =========================================================================
REM MACRO UNIFICADA: ENVIAR A PAPELERA (SOPORTA DIÁLOGO Y COMANDO)
REM =========================================================================
Sub MoverTareaAPapelera(Optional oEvent As Object)
    Dim oDoc As Object, oData As Object, oVisor As Object, oSel As Object
    Dim oDialogo As Object
    Dim idActivo As Long, TotalRows As Long, i As Long
    Dim filaDatosIdx As Long : filaDatosIdx = -1
    
    ' 1. Determinar el contexto de la llamada (¿Proviene de un diálogo?)
    oDialogo = Nothing
    If Not IsMissing(oEvent) Then
        On Error Resume Next
        ' Intenta recuperar el diálogo contenedor si proviene de un control
        oDialogo = oEvent.Source.Context
        On Error GoTo 0
    End If
    
    ' 2. Cargar entorno y obtener ID activo de la celda enfocada
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oSel = oDoc.CurrentSelection
    
    ' Salida de seguridad si no hay una sola celda seleccionada
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    idActivo = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If idActivo <= 0 Then
        MsgBox "No se detectó ninguna tarea activa para mover a la papelera.", 48, "Atención"
        Exit Sub
    End If
    
    ' 3. Localizar la fila física en la base de datos (Columna F / índice 5)
    Dim oCursor As Object
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        If CLng(oData.getCellByPosition(5, i).Value) = idActivo Then
            filaDatosIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Ejecución del cambio de estado lógico
    If filaDatosIdx <> -1 Then
        ' Bloquear actualización de pantalla
        oDoc.addActionLock()
        
        ' Cambiar el Estado a "Papelera" en la Columna C (índice 2)
        oData.getCellByPosition(2, filaDatosIdx).String = "Papelera"
        
        ' Asignar un orden muy alto para sacarla de las prioridades activas
        oData.getCellByPosition(4, filaDatosIdx).Value = 99999
        
        ' Limpiar descripciones en el Visor
        RunInjectionDescription(oVisor, 0)
        
        ' Reordenar la base de datos y redibujar el tablero síncronamente
        CustomDataSort()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        ' 5. Gestión del Diálogo: Si se detectó el objeto en memoria, lo cerramos
        If Not IsNull(oDialogo) And Not IsEmpty(oDialogo) Then
            oDialogo.endExecute()
        End If
        
        MsgBox "La tarea ha sido movida a la papelera.", 64, "Proceso Completado"
    Else
        MsgBox "Error: No se pudo localizar la tarea en la base de datos.", 48, "Error de Sincronización"
    End If
End Sub

REM =========================================================================
REM MACRO UNIFICADA: ELIMINACIÓN PERMANENTE (SOPORTA DIÁLOGO Y COMANDO)
REM =========================================================================
Sub EliminarTareaActiva(Optional oEvent As Object)
    Dim oDoc As Object, oData As Object, oVisor As Object, oSel As Object
    Dim oDialogo As Object
    Dim idActivo As Long, TotalRows As Long, i As Long
    Dim filaDatosIdx As Long : filaDatosIdx = -1
    Dim answer As Integer
    Dim sNombreTareaBorrar As String, sActivityBorrar As String
    
    ' 1. Determinar el contexto de la llamada (¿Proviene de un diálogo?)
    oDialogo = Nothing
    If Not IsMissing(oEvent) Then
        On Error Resume Next
        ' Intenta recuperar el diálogo contenedor si proviene de un control
        oDialogo = oEvent.Source.Context
        On Error GoTo 0
    End If
    
    ' 2. Cargar entorno y obtener ID activo de la celda enfocada
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    oSel = oDoc.CurrentSelection
    
    ' Salida de seguridad
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    idActivo = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If idActivo <= 0 Then
        MsgBox "No se detectó ninguna tarea activa para eliminar.", 48, "Atención"
        Exit Sub
    End If
    
    ' 3. Localizar la fila física en la base de datos (Columna F / índice 5)
    Dim oCursor As Object
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To TotalRows - 1
        If CLng(oData.getCellByPosition(5, i).Value) = idActivo Then
            filaDatosIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Confirmación crítica y eliminación física
    If filaDatosIdx <> -1 Then
        sActivityBorrar = oData.getCellByPosition(0, filaDatosIdx).String
        sNombreTareaBorrar = oData.getCellByPosition(1, filaDatosIdx).String
        
        ' Cuadro confirmatorio crítico
        answer = MsgBox("⚠️ ¿Estás seguro de eliminar permanentemente esta tarea? ⚠️" & Chr(13) & Chr(13) & _
                           "📌 Actividad: " & sActivityBorrar & Chr(13) & _
                           "📝 Tarea: """ & sNombreTareaBorrar & """" & Chr(13) & _
                           "🆔 ID Único: " & idActivo & Chr(13) & Chr(13) & _
                           "Esta acción eliminará la fila por completo de la base de datos. Es irreversible.", 4 + 16 + 256, "Confirmar Eliminación Permanente")
                           
        If answer <> 6 Then Exit Sub ' Si cancela, salimos
        
        ' Bloquear actualización de pantalla
        oDoc.addActionLock()
        
        ' Eliminar fila de la base de datos
        oData.Rows.removeByIndex(filaDatosIdx, 1)
        
        ' Limpiar descripciones residuales en el visor (Celda D4)
        RunInjectionDescription(oVisor, 0)
        
        ' Reordenar secuencialmente y redibujar el lienzo Kanban
        CustomDataSort()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        ' 5. Gestión del Diálogo: Si se detectó el objeto en memoria, lo cerramos
        If Not IsNull(oDialogo) And Not IsEmpty(oDialogo) Then
            oDialogo.endExecute()
        End If
        
        MsgBox "La tarea ha sido eliminada con éxito.", 64, "Proceso Completado"
    Else
        MsgBox "Error: No se pudo localizar la tarea en la base de datos.", 48, "Error de Sincronización"
    End If
End Sub



REM =========================================================================
REM INVOCADORES
REM =========================================================================


'------------------------------------------------------------------------------------------------------
REM =========================================================================
REM MACRO DE EDICIÓN: EDITAR TAREA SELECCIONADA (USANDO DIÁLOGO UNIFICADO)
REM =========================================================================
Sub EditarTareaActiva()
    Dim oDoc As Object, oVisor As Object, oSel As Object
    Dim oDialogo As Object
    Dim idActivo As Long
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oSel = oDoc.CurrentSelection
    
    ' 1. Detectar el ID de la tarea seleccionada en el visor
    idActivo = 0
    If oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        idActivo = GetIdFromPosition(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    End If
    
    ' Si la celda seleccionada está vacía o es una cabecera, salimos de inmediato
    If idActivo <= 0 Then
        MsgBox "Por favor, seleccione primero una tarea válida en el tablero.", 48, "Atención"
        Exit Sub
    End If
    
    ' 2. Cargar e instanciar el diálogo unificado "DialogoPanelTarea"
    DialogLibraries.LoadLibrary("Standard")
    oDialogo = CreateUnoDialog(DialogLibraries.Standard.DialogoPanelTarea)
    
    ' Congelar listener para estabilidad visual
    binFrozenFocus = True
    
    ' 3. Invocar al Hub de Lectura con el ID seleccionado (Esto mostrará BtnEliminar y BtnPapelera)
    Call InicializarServiciosDialogo(oDialogo, idActivo)
    
    ' 4. Ejecución del diálogo
    If oDialogo.execute() = 1 Then
        ' Guardar las modificaciones sobre el registro existente
        Call GuardarDatosDesdeDialogo(oDialogo, idActivo)
    End If
    
    ' Liberar recursos de forma segura
    binFrozenFocus = False
    oDialogo.dispose()
End Sub

REM =========================================================================
REM MACRO DE CREACIÓN: REGISTRAR NUEVA TAREA (USANDO DIÁLOGO UNIFICADO)
REM =========================================================================
Sub RegistrarNuevaTarea()
    Dim oDoc As Object, oVisor As Object, oData As Object
    Dim oDialogo As Object
    Dim sActivityFiltro As String, sStateFiltro As String
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("view")
    oData = oDoc.Sheets.getByName("data")
    
    ' 1. Cargar e instanciar el diálogo unificado "DialogoPanelTarea"
    DialogLibraries.LoadLibrary("Standard")
    oDialogo = CreateUnoDialog(DialogLibraries.Standard.DialogoPanelTarea)
    
    ' Congelar el listener temporalmente para evitar parpadeos visuales
    binFrozenFocus = True
    
    ' 2. Invocar al Hub de Lectura con ID = 0 (Esto ocultará automáticamente BtnEliminar y BtnPapelera)
    Call InicializarServiciosDialogo(oDialogo, 0)
    
    ' 3. Pre-selección inteligente basada en los filtros activos del Visor
    sActivityFiltro = Trim(oVisor.getCellByPosition(0, 1).String) ' Celda A2
    sStateFiltro = Trim(oVisor.getCellByPosition(0, 2).String)    ' Celda A3
    
    If sActivityFiltro <> "" Then
        oDialogo.getControl("CmbActividad").setText(sActivityFiltro)
    End If
    
    If sStateFiltro <> "" Then
        oDialogo.getControl("CmbEstado").setText(sStateFiltro)
    Else
        oDialogo.getControl("CmbEstado").setText("In progress")
    End If
    
    ' 4. Ejecución del diálogo
    If oDialogo.execute() = 1 Then
        ' Validación de campos obligatorios
        If Trim(oDialogo.getControl("TxtTitulo").Text) = "" Then
            MsgBox "La 'Actividad' y el nombre de la 'Tarea' son obligatorios.", 48, "Campos Incompletos"
        Else
            ' Guardar la nueva tarea en la base de datos (ID = 0)
            Call GuardarDatosDesdeDialogo(oDialogo, 0)
        End If
    End If
    
    ' Liberar entorno de forma segura
    binFrozenFocus = False
    oDialogo.dispose()
End Sub

REM =========================================================================
REM PUNTO DE ENTRADA: GESTIONAR PROYECTOS / ACTIVIDADES
REM =========================================================================
Sub GestionarActividades()
    G_ManagedItemType = "Activity"
    Call AbrirPanelItem()
End Sub

REM =========================================================================
REM PUNTO DE ENTRADA: GESTIONAR COLUMNAS / ESTADOS
REM =========================================================================
Sub GestionarEstados()
    G_ManagedItemType = "State"
    Call AbrirPanelItem()
End Sub

REM =========================================================================
REM INVOCADOR UNIFICADO: ABRIR PANEL DE ITEM
REM =========================================================================
Sub AbrirPanelItem()
    Dim oDialogo As Object
    
    ' 1. Instanciar el diálogo unificado "DialogoPanelItem"
    DialogLibraries.LoadLibrary("Standard")
    oDialogo = CreateUnoDialog(DialogLibraries.Standard.DialogoPanelItem)
    
    ' Configurar el título dinámicamente según el contexto global
    oDialogo.Title = "Gestión de " & G_ManagedItemType & "es"
    
    ' 2. Invocar al Hub de Lectura (ID = 0) para poblar el ComboBox "CmbItem"
    Call InicializarServiciosDialogo(oDialogo, 0)
    
    ' Deshabilitar inicialmente el botón aceptar hasta que el usuario escriba o seleccione
    oDialogo.getControl("BtnAceptar").Model.Enabled = False
    
    ' 3. Desplegar panel en pantalla
    oDialogo.execute()
    oDialogo.dispose()
End Sub






REM =========================================================================
REM INTERACTIVO: CAMBIO DINÁMICO DE BOTÓN EN TIEMPO DE EJECUCIÓN
REM =========================================================================
Sub CmbItemTextoModificado(oEvent As Object)
    Dim oDialogo As Object, oBtnModel As Object, oCmbItem As Object
    Dim sTexto As String
    Dim vlist() As String
    Dim binExiste As Boolean
    
    ' Obtener referencias
    oCmbItem = oEvent.Source
    oDialogo = oCmbItem.Context
    oBtnModel = oDialogo.getControl("BtnAceptar").Model
    
    sTexto = Trim(oCmbItem.Text)
    
    ' 1. Estado Vacío: Deshabilitar el botón y restablecerlo a neutral
    If sTexto = "" Then
        oBtnModel.Enabled = False
        oBtnModel.Label = "Aceptar"
        oBtnModel.BackgroundColor = RGB(240, 240, 240) ' Gris neutral
        oBtnModel.TextColor = RGB(0, 0, 0)             ' Texto negro
        Exit Sub
    End If
    
    ' 2. Verificar si el texto ya existe en la list del combo
    vlist() = oCmbItem.Model.StringItemList
    binExiste = ElementoExisteEnlist(sTexto, vlist())
    
    ' 3. Aplicar mutación visual dinámica al botón
    If binExiste Then
        ' --- MODO ELIMINACIÓN (Existe en la list) ---
        oBtnModel.Enabled = True
        oBtnModel.Label = "⮿ Eliminar"
        oBtnModel.BackgroundColor = RGB(220, 53, 69)   ' Rojo de advertencia
        oBtnModel.TextColor = RGB(255, 255, 255)       ' Texto blanco para contraste
    Else
        ' --- MODO ADICIÓN (Nuevo elemento) ---
        oBtnModel.Enabled = True
        oBtnModel.Label = "[+] Agregar"
        oBtnModel.BackgroundColor = RGB(40, 167, 69)   ' Verde de inserción
        oBtnModel.TextColor = RGB(255, 255, 255)       ' Texto blanco
    End If
End Sub

REM =========================================================================
REM HUB DE ESCRITURA DE CATÁLOGOS: AGREGAR O ELIMINAR SEGÚN CONTEXTO (MUTADO)
REM =========================================================================
Sub GuardarCatalogoDesdeDialogo(oEvent As Object)
    Dim oDialogo As Object, oData As Object, oCmbItem As Object
    Dim sTexto As String
    Dim vlist() As String
    Dim binExiste As Boolean
    Dim answer As Integer
    
    oDialogo = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    oCmbItem = oDialogo.getControl("CmbItem")
    sTexto = Trim(oCmbItem.Text)
    
    If sTexto = "" Then Exit Sub
    
    vlist() = oCmbItem.Model.StringItemList
    binExiste = ElementoExisteEnlist(sTexto, vlist())
    
    If binExiste Then
        ' OPERACIÓN: ELIMINACIÓN EN CASCADA
        oDialogo.endExecute()
        Call RunCascadeDelete(sTexto, G_ManagedItemType)
    Else
        ' OPERACIÓN: AGREGAR NUEVO ELEMENTO
        answer = MsgBox("¿Deseas agregar la nueva " & G_ManagedItemType & ": """ & sTexto & """?", 4 + 32, "Confirmar Adición")
        
        If answer = 6 Then ' SÍ
            Dim mDatos() As Variant, i As Long, filaVacia As Long
            Dim nColIdx As Long
            
            filaVacia = -1
            If G_ManagedItemType = "Activity" Then
                nColIdx = 7
                mDatos = oData.getCellRangeByName("H2:H100").getDataArray()
            Else
                nColIdx = 9
                mDatos = oData.getCellRangeByName("J2:J20").getDataArray()
            End If
            
            For i = 0 To UBound(mDatos)
                If Trim(mDatos(i)(0)) = "" Then
                    filaVacia = i + 2
                    Exit For
                End If
            Next i
            
            If filaVacia <> -1 Then
                oData.getCellByPosition(nColIdx, filaVacia - 1).String = sTexto
                oDialogo.endExecute()
                
                CustomDataSort()
                MsgBox G_ManagedItemType & " agregada correctamente al catálogo.", 64, "Éxito"
            Else
                MsgBox "Error: El catálogo de " & G_ManagedItemType & "es está lleno.", 48, "Límite Excedido"
            End If
        End If
    End If
End Sub

REM =========================================================================
REM FUNCIÓN AUXILIAR: COMPARACIÓN INSENSIBLE A MAYÚSCULAS EN ARREGLOS
REM =========================================================================
Function ElementoExisteEnlist(sBuscar As String, ByRef vlist() As String) As Boolean
    Dim i As Long
    sBuscar = LCase(Trim(sBuscar))
    For i = 0 To UBound(vlist)
        If LCase(Trim(vlist(i))) = sBuscar Then
            ElementoExisteEnlist = True
            Exit Function
        End If
    Next i
    ElementoExisteEnlist = False
End Function
