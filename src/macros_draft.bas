Global G_ManagedItemType As String
Global G_TextShortcut As String
Global G_UsedShortcut As Boolean
Global binFrozenFocus As Boolean
Global oSeleccionListener As Object
Global oSheetHeard As Object
Global G_SelectedItem As String
Global G_SelectedPos As Long
Global G_SelectedRole As String
Global G_SelectedScope As String
Global G_SelectedActState As String
Global G_UpdatingPanel As Boolean
Global G_LastActivity As String
Global G_GateBusy As Boolean
Global G_BoardFilterBusy As Boolean

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
REM CATALOG MANAGEMENT ENGINE (CHANGE 1)
REM Data-driven ordering (matrix position) and state roles (column K).
REM =========================================================================

REM Returns the catalog column index for the given item type (H = 7, J = 9).
Function GetCatalogColumn(sType As String) As Long
    If sType = "Activity" Then
        GetCatalogColumn = 7
    Else
        GetCatalogColumn = 9
    End If
End Function

REM Returns the state role column index (K = 10).
Function GetRoleColumn() As Long
    GetRoleColumn = 10
End Function

REM Returns the last row index used by the catalog (H2:H100 / J2:J20).
Function GetCatalogMaxRow(sType As String) As Long
    If sType = "Activity" Then
        GetCatalogMaxRow = 99   ' H2:H100
    Else
        GetCatalogMaxRow = 19   ' J2:J20
    End If
End Function

REM Returns a Variant array with the non-empty catalog names in matrix order.
Function GetCatalogNames(sType As String) As Variant
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long, nCol As Long
    Dim vNames() As Variant

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    nCol = GetCatalogColumn(sType)

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    ReDim vNames(TotalRows - 2)
    n = 0
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(nCol, i).String) <> "" Then
            vNames(n) = Trim(oData.getCellByPosition(nCol, i).String)
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim vNames(0)
        vNames(0) = ""
    Else
        ReDim Preserve vNames(n - 1)
    End If
    GetCatalogNames = vNames()
End Function

REM Returns a Variant array with the state roles (column K) aligned to GetCatalogNames("State").
Function GetStateRoles() As Variant
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vRoles() As Variant

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    ReDim vRoles(TotalRows - 2)
    n = 0
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(9, i).String) <> "" Then
            vRoles(n) = Trim(oData.getCellByPosition(10, i).String)
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim vRoles(0)
        vRoles(0) = ""
    Else
        ReDim Preserve vRoles(n - 1)
    End If
    GetStateRoles = vRoles()
End Function

REM Returns the role ("base", "trash" or "") for a given state name.
Function GetStateRole(sState As String) As String
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sTarget As String

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    sTarget = LCase(Trim(sState))

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(9, i).String)) = sTarget Then
            GetStateRole = LCase(Trim(oData.getCellByPosition(10, i).String))
            Exit Function
        End If
    Next i
    GetStateRole = ""
End Function

REM Returns the name of the state marked with the "trash" role (or "").
Function GetTrashStateName() As String
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(10, i).String)) = "trash" Then
            GetTrashStateName = Trim(oData.getCellByPosition(9, i).String)
            Exit Function
        End If
    Next i
    GetTrashStateName = ""
End Function

REM True when the state is protected (role "base" or "trash").
Function IsBaseState(sState As String) As Boolean
    Dim sRole As String
    sRole = GetStateRole(sState)
    IsBaseState = (sRole = "base" Or sRole = "trash")
End Function

REM Returns a Variant array with the activity states (column I) aligned to GetCatalogNames("Activity").
Function GetActivityStates() As Variant
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vStates() As Variant

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    ReDim vStates(TotalRows - 2)
    n = 0
    Dim sVal As String
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(7, i).String) <> "" Then
            sVal = Trim(oData.getCellByPosition(8, i).String)
            If LCase(sVal) = "active" Then sVal = ""
            vStates(n) = sVal
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim vStates(0)
        vStates(0) = ""
    Else
        ReDim Preserve vStates(n - 1)
    End If
    GetActivityStates = vStates()
End Function

REM Returns a Variant array with the state scopes (column L) aligned to GetCatalogNames("State").
Function GetStateScopes() As Variant
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vScopes() As Variant

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    ReDim vScopes(TotalRows - 2)
    n = 0
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(9, i).String) <> "" Then
            vScopes(n) = Trim(oData.getCellByPosition(11, i).String)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim vScopes(0)
        vScopes(0) = ""
    Else
        ReDim Preserve vScopes(n - 1)
    End If
    GetStateScopes = vScopes()
End Function

REM Returns the scope (owner activity, or "") for a given state name.
Function GetStateScope(sState As String) As String
    Dim oDoc As Object, oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sTarget As String

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    sTarget = LCase(Trim(sState))
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(9, i).String)) = sTarget Then
            GetStateScope = Trim(oData.getCellByPosition(11, i).String)
            Exit Function
        End If
    Next i
    GetStateScope = ""
End Function

REM Clears and writes the activity catalog (H) plus activity states (I).
Sub WriteActivityCatalog(vNames As Variant, vStates As Variant)
    Dim oDoc As Object, oData As Object
    Dim i As Long, n As Long, nMaxRow As Long, nTop As Long

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    nMaxRow = GetCatalogMaxRow("Activity")

    For i = 1 To nMaxRow
        oData.getCellByPosition(7, i).String = ""
        oData.getCellByPosition(8, i).String = ""
    Next i

    nTop = UBound(vStates)
    n = 1
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then
            oData.getCellByPosition(7, n).String = CStr(vNames(i))
            If i <= nTop Then oData.getCellByPosition(8, n).String = CStr(vStates(i))
            n = n + 1
        End If
    Next i
End Sub

REM Clears and writes the state catalog (J), roles (K) and scopes (L).
Sub WriteStateCatalog(vNames As Variant, vRoles As Variant, vScopes As Variant)
    Dim oDoc As Object, oData As Object
    Dim i As Long, n As Long, nMaxRow As Long, nTopR As Long, nTopS As Long

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    nMaxRow = GetCatalogMaxRow("State")

    For i = 1 To nMaxRow
        oData.getCellByPosition(9, i).String = ""
        oData.getCellByPosition(10, i).String = ""
        oData.getCellByPosition(11, i).String = ""
    Next i

    nTopR = UBound(vRoles)
    nTopS = UBound(vScopes)
    n = 1
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then
            oData.getCellByPosition(9, n).String = CStr(vNames(i))
            If i <= nTopR Then oData.getCellByPosition(10, n).String = CStr(vRoles(i))
            If i <= nTopS Then oData.getCellByPosition(11, n).String = CStr(vScopes(i))
            n = n + 1
        End If
    Next i
End Sub

REM Rebuilds the catalog as a contiguous list (removes blank cells).
Sub CompactCatalog(sType As String)
    Dim vNames As Variant, vRoles As Variant, vScopes As Variant

    vNames = GetCatalogNames(sType)
    If sType = "Activity" Then
        vScopes = GetActivityStates()
        Call WriteActivityCatalog(vNames, vScopes)
    Else
        vRoles = GetStateRoles()
        vScopes = GetStateScopes()
        Call WriteStateCatalog(vNames, vRoles, vScopes)
    End If
End Sub

REM Adds a new element at the end of the catalog.
Sub AddCatalogItem(sName As String, sType As String, Optional sScope As String, Optional nPos As Long, Optional sRole As String, Optional sActState As String)
    Dim vNames As Variant, vAux1 As Variant, vAux2 As Variant
    Dim vOutNames() As Variant, vOutAux1() As Variant, vOutAux2() As Variant
    Dim i As Long, n As Long, nCount As Long, nAt As Long
    Dim bState As Boolean
    Dim sA As String

    If IsMissing(sScope) Then sScope = ""
    If IsMissing(nPos) Then nPos = 0
    If IsMissing(sRole) Then sRole = ""
    If IsMissing(sActState) Then sActState = ""

    bState = (sType = "State")
    vNames = GetCatalogNames(sType)
    If bState Then
        vAux1 = GetStateRoles()
        vAux2 = GetStateScopes()
    Else
        vAux1 = GetActivityStates()
    End If

    nCount = 0
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then nCount = nCount + 1
    Next i

    ' Determine the insertion index (0-based) for the new element.
    If nPos < 1 Or nPos > nCount + 1 Then
        nAt = nCount
    Else
        nAt = nPos - 1
    End If

    ReDim vOutNames(nCount)
    ReDim vOutAux1(nCount)
    If bState Then ReDim vOutAux2(nCount)

    n = 0
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then
            If n = nAt Then n = n + 1
            vOutNames(n) = vNames(i)
            vOutAux1(n) = vAux1(i)
            If bState Then vOutAux2(n) = vAux2(i)
            n = n + 1
        End If
    Next i

    vOutNames(nAt) = sName
    If bState Then
        vOutAux1(nAt) = LCase(Trim(sRole))
        vOutAux2(nAt) = Trim(sScope)
        Call WriteStateCatalog(vOutNames, vOutAux1, vOutAux2)
    Else
        sA = Trim(sActState)
        If LCase(sA) = "active" Then sA = ""
        vOutAux1(nAt) = sA
        Call WriteActivityCatalog(vOutNames, vOutAux1)
    End If

    If bState Then Call RefreshBoardStateFilters()
    Call CustomDataSort()
    MsgBox sType & " successfully added to the catalog.", 64, "Success"
End Sub

REM Renames a catalog element, reassigns tasks and updates the view filters.
Sub RenameCatalogItem(sOld As String, sNew As String, sType As String)
    Dim oDoc As Object, oData As Object, oVisor As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim nCol As Long, nTaskCol As Long
    Dim sTarget As String

    oDoc = ThisComponent
    oData = oDoc.Sheets.getByName("data")
    oVisor = oDoc.Sheets.getByName("view")
    nCol = GetCatalogColumn(sType)
    sTarget = LCase(Trim(sOld))

    ' 1. Update the catalog name.
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(nCol, i).String)) = sTarget Then
            oData.getCellByPosition(nCol, i).String = sNew
            Exit For
        End If
    Next i

    ' 2. Reassign the tasks (data.A for activities, data.C for states).
    If sType = "Activity" Then
        nTaskCol = 0
    Else
        nTaskCol = 2
    End If
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(nTaskCol, i).String)) = sTarget Then
            oData.getCellByPosition(nTaskCol, i).String = sNew
        End If
    Next i

    ' 3. Keep the view filters in sync with the new name.
    If sType = "Activity" Then
        ' A2 (0,1) holds the activity filter.
        If LCase(Trim(oVisor.getCellByPosition(0, 1).String)) = sTarget Then
            oVisor.getCellByPosition(0, 1).String = sNew
        End If
        ' Update the scope (column L) of per-activity states owned by this activity.
        For i = 1 To TotalRows - 1
            If LCase(Trim(oData.getCellByPosition(11, i).String)) = sTarget Then
                oData.getCellByPosition(11, i).String = sNew
            End If
        Next i
    Else
        ' A3 (0,2) and B3 (1,2) hold the two visible states.
        If LCase(Trim(oVisor.getCellByPosition(0, 2).String)) = sTarget Then
            oVisor.getCellByPosition(0, 2).String = sNew
        End If
        If LCase(Trim(oVisor.getCellByPosition(1, 2).String)) = sTarget Then
            oVisor.getCellByPosition(1, 2).String = sNew
        End If
    End If
End Sub

REM Returns the number of elements in the catalog.
Function CountCatalog(sType As String) As Long
    Dim vNames As Variant, i As Long, n As Long

    vNames = GetCatalogNames(sType)
    n = 0
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then n = n + 1
    Next i
    CountCatalog = n
End Function

REM Resolves a position text ("1".."N" or "F") into a 1-based position.
Function ResolvePosition(sText As String, nCount As Long) As Long
    Dim s As String, n As Long

    s = Trim(sText)
    If nCount < 1 Then nCount = 1
    If s = "" Or UCase(s) = "F" Or Not IsNumeric(s) Then
        ResolvePosition = nCount
    Else
        n = CLng(s)
        If n < 1 Then n = 1
        If n > nCount Then n = nCount
        ResolvePosition = n
    End If
End Function

REM Maps a role label to its stored value.
Function RoleLabelToValue(sLabel As String) As String
    Select Case LCase(Trim(sLabel))
        Case "base"
            RoleLabelToValue = "base"
        Case "trash"
            RoleLabelToValue = "trash"
        Case Else
            RoleLabelToValue = ""
    End Select
End Function

REM Populates the position combo with "1".."N" plus "F" and selects the current position.
Sub PopulateCmbPos(oDialogue As Object, sItem As String)
    Dim vNames As Variant, vList() As String
    Dim i As Long, nCount As Long, nPos As Long

    vNames = GetCatalogNames(G_ManagedItemType)
    nCount = CountCatalog(G_ManagedItemType)
    If nCount < 1 Then nCount = 1

    ReDim vList(nCount)
    For i = 1 To nCount
        vList(i - 1) = CStr(i)
    Next i
    vList(nCount) = "F"

    nPos = 1
    For i = 0 To UBound(vNames)
        If LCase(Trim(CStr(vNames(i)))) = LCase(Trim(sItem)) Then
            nPos = i + 1
            Exit For
        End If
    Next i
    G_SelectedPos = nPos

    G_UpdatingPanel = True
    oDialogue.getControl("CmbPos").Model.StringItemList = vList()
    oDialogue.getControl("CmbPos").setText(CStr(nPos))
    oDialogue.getControl("CmbPos").Model.Enabled = True
    G_UpdatingPanel = False
End Sub

REM Populates the role combo and selects the current role.
Sub PopulateCmbRole(oDialogue As Object, sItem As String)
    Dim vList(2) As String
    Dim sRole As String, sLabel As String

    vList(0) = "Normal"
    vList(1) = "Base"
    vList(2) = "Trash"

    sRole = GetStateRole(sItem)
    G_SelectedRole = sRole
    Select Case sRole
        Case "base"
            sLabel = "Base"
        Case "trash"
            sLabel = "Trash"
        Case Else
            sLabel = "Normal"
    End Select

    G_UpdatingPanel = True
    oDialogue.getControl("CmbRole").Model.StringItemList = vList()
    oDialogue.getControl("CmbRole").setText(sLabel)
    oDialogue.getControl("CmbRole").Model.Enabled = True
    G_UpdatingPanel = False
End Sub

REM Returns the activity-state names catalog (column M).
Function GetActStateNames() As Variant
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vNames() As Variant

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    ReDim vNames(TotalRows - 2)
    n = 0
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(12, i).String) <> "" Then
            vNames(n) = Trim(oData.getCellByPosition(12, i).String)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim vNames(0)
        vNames(0) = ""
    Else
        ReDim Preserve vNames(n - 1)
    End If
    GetActStateNames = vNames()
End Function

REM Maps the activity-state label to its stored value ("" = active).
Function ActivityStateLabelToValue(sLabel As String) As String
    If LCase(Trim(sLabel)) = "active" Then
        ActivityStateLabelToValue = ""
    Else
        ActivityStateLabelToValue = Trim(sLabel)
    End If
End Function

REM Fills the scope filter (General + activity names) and selects General.
Sub PopulateScopeFilter(oDialogue As Object)
    Dim vActs As Variant, vList() As String
    Dim i As Long, n As Long

    vActs = GetCatalogNames("Activity")
    ReDim vList(UBound(vActs) + 1)
    vList(0) = "General"
    n = 1
    For i = 0 To UBound(vActs)
        If Trim(CStr(vActs(i))) <> "" Then
            vList(n) = CStr(vActs(i))
            n = n + 1
        End If
    Next i
    If n > 1 Then
        ReDim Preserve vList(n - 1)
    Else
        ReDim vList(0)
        vList(0) = "General"
    End If

    G_UpdatingPanel = True
    oDialogue.getControl("CmbScope").Model.StringItemList = vList()
    oDialogue.getControl("CmbScope").setText("General")
    oDialogue.getControl("CmbScope").Model.Enabled = True
    G_UpdatingPanel = False
End Sub

REM Returns "" for Generalized scope, or the selected activity name.
Function GetSelectedScopeValue(oDialogue As Object) As String
    Dim s As String
    s = Trim(oDialogue.getControl("CmbScope").Text)
    If LCase(s) = "general" Then s = ""
    GetSelectedScopeValue = s
End Function

REM Fills CmbItem with the states that belong to the selected scope.
Sub ServPopulateCmbItemFiltered(oDialogue As Object, oData As Object, sScope As String)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vList() As String
    Dim sState As String, sL As String

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vList(TotalRows)
    n = 0
    For i = 1 To TotalRows - 1
        sState = Trim(oData.getCellByPosition(9, i).String)
        sL = Trim(oData.getCellByPosition(11, i).String)
        If sState <> "" Then
            If (LCase(sScope) = "general" And sL = "") Or _
               (LCase(sScope) <> "general" And LCase(sL) = LCase(sScope)) Then
                vList(n) = sState
                n = n + 1
            End If
        End If
    Next i

    G_UpdatingPanel = True
    If n > 0 Then
        ReDim Preserve vList(n - 1)
        oDialogue.getControl("CmbItem").Model.StringItemList = vList()
    Else
        Dim vEmpty(0) As String
        oDialogue.getControl("CmbItem").Model.StringItemList = vEmpty()
    End If
    G_UpdatingPanel = False
End Sub

REM Handles a change of the scope filter in the states panel.
Sub CmbScopeChanged(oEvent As Object)
    Dim oDialogue As Object, oData As Object
    Dim sScope As String

    If G_UpdatingPanel Then Exit Sub
    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    sScope = Trim(oDialogue.getControl("CmbScope").Text)

    Call ServPopulateCmbItemFiltered(oDialogue, oData, sScope)

    G_UpdatingPanel = True
    oDialogue.getControl("CmbItem").setText("")
    oDialogue.getControl("TxtNewName").Text = ""
    oDialogue.getControl("CmbPos").setText("")
    oDialogue.getControl("CmbRole").setText("")
    G_UpdatingPanel = False
    Call ApplyWatermarks(oDialogue)

    G_SelectedItem = ""
    G_SelectedPos = 0
    G_SelectedRole = ""
    Call SetSelectionControlsEnabled(oDialogue, False)
    Call UpdateItemPanelButton(oEvent)
End Sub

REM Populates the activity-state combo (Active + catalog) and selects the current state.
Sub PopulateCmbActState(oDialogue As Object, sItem As String)
    Dim vStates As Variant, vList() As String
    Dim sState As String, sLabel As String
    Dim i As Long, n As Long
    Dim bHasActive As Boolean

    vStates = GetActStateNames()
    ReDim vList(UBound(vStates) + 1)
    n = 0
    bHasActive = False
    For i = 0 To UBound(vStates)
        If Trim(CStr(vStates(i))) <> "" Then
            vList(n) = CStr(vStates(i))
            If LCase(Trim(CStr(vStates(i)))) = "active" Then bHasActive = True
            n = n + 1
        End If
    Next i

    ' Ensure there is an "active" option (the catalog should provide it).
    If Not bHasActive Then
        For i = n To 1 Step -1
            vList(i) = vList(i - 1)
        Next i
        vList(0) = "Active"
        n = n + 1
    End If

    If n > 0 Then
        ReDim Preserve vList(n - 1)
    Else
        ReDim vList(0)
        vList(0) = "Active"
    End If

    sState = GetActivityStateName(sItem)
    G_SelectedActState = sState
    sLabel = ""
    If Trim(sState) = "" Then
        For i = 0 To UBound(vList)
            If LCase(Trim(vList(i))) = "active" Then
                sLabel = vList(i)
                Exit For
            End If
        Next i
    Else
        sLabel = sState
    End If

    G_UpdatingPanel = True
    oDialogue.getControl("CmbActState").Model.StringItemList = vList()
    oDialogue.getControl("CmbActState").setText(sLabel)
    oDialogue.getControl("CmbActState").Model.Enabled = True
    G_UpdatingPanel = False
End Sub

REM Prepares the selector controls for adding a new element.
Sub PopulateAddDefaults(oDialogue As Object)
    Dim vList() As String
    Dim vRoles(2) As String
    Dim i As Long, nCount As Long

    If G_ManagedItemType = "State" Then
        nCount = CountCatalog("State") + 1
        If nCount < 1 Then nCount = 1
        ReDim vList(nCount)
        For i = 1 To nCount
            vList(i - 1) = CStr(i)
        Next i
        vList(nCount) = "F"

        vRoles(0) = "Normal"
        vRoles(1) = "Base"
        vRoles(2) = "Trash"

        G_UpdatingPanel = True
        oDialogue.getControl("CmbPos").Model.StringItemList = vList()
        oDialogue.getControl("CmbPos").setText("F")
        oDialogue.getControl("CmbPos").Model.Enabled = True
        oDialogue.getControl("CmbRole").Model.StringItemList = vRoles()
        oDialogue.getControl("CmbRole").setText("Normal")
        oDialogue.getControl("CmbRole").Model.Enabled = True
        G_UpdatingPanel = False
        G_SelectedPos = 0
        G_SelectedRole = ""
    Else
        nCount = CountCatalog("Activity") + 1
        If nCount < 1 Then nCount = 1
        ReDim vList(nCount)
        For i = 1 To nCount
            vList(i - 1) = CStr(i)
        Next i
        vList(nCount) = "F"

        G_UpdatingPanel = True
        oDialogue.getControl("CmbPos").Model.StringItemList = vList()
        oDialogue.getControl("CmbPos").setText("F")
        oDialogue.getControl("CmbPos").Model.Enabled = True
        G_UpdatingPanel = False
        G_SelectedPos = 0
        G_SelectedActState = ""
        Call PopulateCmbActState(oDialogue, "")
    End If
End Sub

REM Enables or disables the position and role controls.
Sub SetSelectionControlsEnabled(oDialogue As Object, bEnabled As Boolean)
    If oDialogue.Model.hasByName("CmbPos") Then
        oDialogue.getControl("CmbPos").Model.Enabled = bEnabled
    End If
    If oDialogue.Model.hasByName("CmbRole") Then
        If G_ManagedItemType = "State" Then
            oDialogue.getControl("CmbRole").Model.Enabled = bEnabled
        End If
    End If
    If oDialogue.Model.hasByName("CmbActState") Then
        If G_ManagedItemType = "Activity" Then
            oDialogue.getControl("CmbActState").Model.Enabled = bEnabled
        End If
    End If
End Sub

REM Handles a change on the catalog selector (CmbItem).
Sub CmbItemSelectionChanged(oEvent As Object)
    Dim oCtrl As Object, oDialogue As Object
    Dim sText As String
    Dim vList() As String
    Dim bExist As Boolean

    If G_UpdatingPanel Then Exit Sub

    oCtrl = oEvent.Source
    oDialogue = oCtrl.Context
    sText = Trim(oCtrl.Text)

    G_UpdatingPanel = True
    oDialogue.getControl("TxtNewName").Text = ""
    oDialogue.getControl("CmbPos").setText("")
    If G_ManagedItemType = "State" Then
        oDialogue.getControl("CmbRole").setText("")
    Else
        oDialogue.getControl("CmbActState").setText("")
    End If
    G_UpdatingPanel = False
    Call ApplyWatermarks(oDialogue)

    If sText = "" Then
        G_SelectedItem = ""
        G_SelectedPos = 0
        G_SelectedRole = ""
        G_SelectedScope = ""
        G_SelectedActState = ""
        Call SetSelectionControlsEnabled(oDialogue, False)
        Call UpdateItemPanelButton(oEvent)
        Exit Sub
    End If

    vList() = oCtrl.Model.StringItemList
    bExist = ElementExistInList(sText, vList())

    If bExist Then
        G_SelectedItem = sText
        Call PopulateCmbPos(oDialogue, sText)
        If G_ManagedItemType = "State" Then
            Call PopulateCmbRole(oDialogue, sText)
        Else
            Call PopulateCmbActState(oDialogue, sText)
        End If
    Else
        G_SelectedItem = ""
        G_SelectedPos = 0
        G_SelectedRole = ""
        G_SelectedScope = ""
        G_SelectedActState = ""
        Call PopulateAddDefaults(oDialogue)
    End If

    Call UpdateItemPanelButton(oEvent)
End Sub

REM Recomputes the dynamic action button (Add / Delete / Modify + icons).
Sub UpdateItemPanelButton(oEvent As Object)
    Dim oDialogue As Object, oBtn As Object
    Dim sItem As String, sNewName As String, sPos As String, sRoleLabel As String
    Dim sActStateLabel As String
    Dim vList() As String
    Dim bExist As Boolean, bName As Boolean, bPos As Boolean, bRole As Boolean
    Dim bActState As Boolean
    Dim nPos As Long, sRole As String, sIcons As String

    If G_UpdatingPanel Then Exit Sub

    oDialogue = oEvent.Source.Context
    oBtn = oDialogue.getControl("BtnAccept").Model
    sItem = Trim(oDialogue.getControl("CmbItem").Text)

    If sItem = "" Then
        oBtn.Enabled = False
        oBtn.Label = "Accept"
        oBtn.BackgroundColor = RGB(240, 240, 240)
        oBtn.TextColor = RGB(0, 0, 0)
        Exit Sub
    End If

    vList() = oDialogue.getControl("CmbItem").Model.StringItemList
    bExist = ElementExistInList(sItem, vList())

    If Not bExist Then
        ' ADD mode
        oBtn.Enabled = True
        oBtn.Label = "[+] Add"
        oBtn.BackgroundColor = RGB(40, 167, 69)
        oBtn.TextColor = RGB(255, 255, 255)
        Exit Sub
    End If

    ' Existing item: compute the pending changes.
    sNewName = WatermarkValue(oDialogue.getControl("TxtNewName"))
    bName = (sNewName <> "" And LCase(sNewName) <> LCase(sItem))

    bPos = False
    If oDialogue.getControl("CmbPos").Model.Enabled Then
        sPos = Trim(oDialogue.getControl("CmbPos").Text)
        nPos = ResolvePosition(sPos, CountCatalog(G_ManagedItemType))
        bPos = (nPos <> G_SelectedPos)
    End If

    bRole = False
    bActState = False
    If G_ManagedItemType = "State" Then
        If oDialogue.getControl("CmbRole").Model.Enabled Then
            sRoleLabel = Trim(oDialogue.getControl("CmbRole").Text)
            sRole = RoleLabelToValue(sRoleLabel)
            bRole = (sRole <> G_SelectedRole)
        End If
    Else
        If oDialogue.getControl("CmbActState").Model.Enabled Then
            sActStateLabel = Trim(oDialogue.getControl("CmbActState").Text)
            bActState = (LCase(ActivityStateLabelToValue(sActStateLabel)) <> LCase(G_SelectedActState))
        End If
    End If

    If Not bName And Not bPos And Not bRole And Not bActState Then
        ' DELETE mode
        oBtn.Enabled = True
        oBtn.Label = "⮿ Delete"
        oBtn.BackgroundColor = RGB(220, 53, 69)
        oBtn.TextColor = RGB(255, 255, 255)
    Else
        ' MODIFY mode with icons: name, position, role, activity state.
        sIcons = ""
        If bName Then sIcons = sIcons & "[✎]"
        If bPos Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[⇅]"
        End If
        If bRole Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[◇]"
        End If
        If bActState Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[⊙]"
        End If
        oBtn.Enabled = True
        oBtn.Label = "Modify" & sIcons
        oBtn.BackgroundColor = RGB(0, 123, 255)
        oBtn.TextColor = RGB(255, 255, 255)
    End If
End Sub

REM =========================================================================
REM WATERMARK / PLACEHOLDER HELPERS
REM =========================================================================

Function GetPlaceholder(sName As String) As String
    Select Case sName
        Case "TxtNewName"
            GetPlaceholder = "New name (optional)"
        Case "TxtParam"
            GetPlaceholder = "Date (YYYY-MM-DD)"
        Case "TxtMessage"
            GetPlaceholder = "Message shown when blocked"
        Case "TxtTitle"
            GetPlaceholder = "Task title"
        Case "TxtDescription"
            GetPlaceholder = "Description"
        Case Else
            GetPlaceholder = ""
    End Select
End Function

REM Applies grey watermark placeholders to empty text fields.
Sub ApplyWatermarks(oDialogue As Object)
    Dim vNames As Variant, i As Long
    Dim sName As String, sPh As String, oCtrl As Object

    vNames = Array("TxtNewName", "TxtParam", "TxtMessage", "TxtTitle", "TxtDescription")
    For i = 0 To UBound(vNames)
        sName = vNames(i)
        sPh = GetPlaceholder(sName)
        If sPh <> "" And oDialogue.Model.hasByName(sName) Then
            oCtrl = oDialogue.getControl(sName)
            G_UpdatingPanel = True
            If Trim(oCtrl.Text) = "" Then
                oCtrl.Text = sPh
                oCtrl.Model.TextColor = RGB(160, 160, 160)
            End If
            G_UpdatingPanel = False
        End If
    Next i
End Sub

REM Clears the placeholder when the field receives focus.
Sub WatermarkFocus(oEvent As Object)
    Dim oCtrl As Object, sName As String, sPh As String

    oCtrl = oEvent.Source
    sName = oCtrl.Model.Name
    sPh = GetPlaceholder(sName)
    If sPh <> "" Then
        If Trim(oCtrl.Text) = sPh Then
            G_UpdatingPanel = True
            oCtrl.Text = ""
            oCtrl.Model.TextColor = RGB(0, 0, 0)
            G_UpdatingPanel = False
        End If
    End If
End Sub

REM Restores the placeholder when the field loses focus and is empty.
Sub WatermarkBlur(oEvent As Object)
    Dim oCtrl As Object, sName As String, sPh As String

    oCtrl = oEvent.Source
    sName = oCtrl.Model.Name
    sPh = GetPlaceholder(sName)
    If sPh <> "" Then
        If Trim(oCtrl.Text) = "" Then
            G_UpdatingPanel = True
            oCtrl.Text = sPh
            oCtrl.Model.TextColor = RGB(160, 160, 160)
            G_UpdatingPanel = False
        End If
    End If
End Sub

REM Returns the field value, treating the placeholder as empty.
Function WatermarkValue(oCtrl As Object) As String
    Dim s As String
    s = Trim(oCtrl.Text)
    If s = GetPlaceholder(oCtrl.Model.Name) Then s = ""
    WatermarkValue = s
End Function

REM Applies the pending rename / position / role / scope / activity-state changes.
Sub ModifyCatalogItem(sOld As String, sNew As String, bNameChanged As Boolean, bPosChanged As Boolean, nNewPos As Long, bRoleChanged As Boolean, sNewRole As String, bScopeChanged As Boolean, sNewScope As String, bActStateChanged As Boolean, sNewActState As String, sType As String)
    Dim sCurrent As String

    sCurrent = sOld
    If bNameChanged Then
        Call RenameCatalogItem(sOld, sNew, sType)
        sCurrent = sNew
    End If
    If bPosChanged Then
        Call ReorderCatalogItem(sType, sCurrent, nNewPos)
    End If
    If bRoleChanged Then
        Call SetCatalogRole(sCurrent, sNewRole)
    End If
    If bScopeChanged Then
        Call SetCatalogScope(sCurrent, sNewScope)
    End If
    If bActStateChanged Then
        Call SetActivityState(sCurrent, sNewActState)
        ' If the visible activity became gated, switch to an accessible one.
        Call EnsureAccessibleActivity(True)
    End If

    Call RefreshBoardStateFilters()
    Call DrawingBoard()
    MsgBox sType & " updated successfully.", 64, "Success"
End Sub

REM Reorders one catalog element without touching task rows (no CustomDataSort).
Sub ReorderCatalogItem(sType As String, sItem As String, nNewPos As Long)
    Dim vNames As Variant, vAux1 As Variant, vAux2 As Variant
    Dim vOutNames() As Variant, vOutAux1() As Variant, vOutAux2() As Variant
    Dim i As Long, n As Long, nCount As Long, nOld As Long, nTarget As Long
    Dim bState As Boolean

    bState = (sType = "State")
    vNames = GetCatalogNames(sType)
    If bState Then
        vAux1 = GetStateRoles()
        vAux2 = GetStateScopes()
    Else
        vAux1 = GetActivityStates()
    End If

    ' Count the valid entries.
    nCount = 0
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" Then nCount = nCount + 1
    Next i
    If nCount = 0 Then Exit Sub
    If nNewPos < 1 Then nNewPos = 1
    If nNewPos > nCount Then nNewPos = nCount

    ' Locate the current position (0-based).
    nOld = -1
    For i = 0 To UBound(vNames)
        If LCase(Trim(CStr(vNames(i)))) = LCase(Trim(sItem)) Then
            nOld = i
            Exit For
        End If
    Next i
    If nOld = -1 Then Exit Sub
    If (nOld + 1) = nNewPos Then Exit Sub

    ' Build the list without the element.
    ReDim vOutNames(nCount - 1)
    ReDim vOutAux1(nCount - 1)
    If bState Then ReDim vOutAux2(nCount - 1)

    n = 0
    For i = 0 To UBound(vNames)
        If Trim(CStr(vNames(i))) <> "" And i <> nOld Then
            vOutNames(n) = vNames(i)
            vOutAux1(n) = vAux1(i)
            If bState Then vOutAux2(n) = vAux2(i)
            n = n + 1
        End If
    Next i

    ' Shift to the right to open the target slot, then insert.
    nTarget = nNewPos - 1
    For i = nCount - 1 To nTarget + 1 Step -1
        vOutNames(i) = vOutNames(i - 1)
        vOutAux1(i) = vOutAux1(i - 1)
        If bState Then vOutAux2(i) = vOutAux2(i - 1)
    Next i
    vOutNames(nTarget) = sItem
    vOutAux1(nTarget) = vAux1(nOld)
    If bState Then vOutAux2(nTarget) = vAux2(nOld)

    ' Write only the catalog columns; task rows are untouched.
    If bState Then
        Call WriteStateCatalog(vOutNames, vOutAux1, vOutAux2)
    Else
        Call WriteActivityCatalog(vOutNames, vOutAux1)
    End If
End Sub

REM Sets the role of a state (empty, "base" or "trash").
Sub SetCatalogRole(sItem As String, sNewRole As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sRole As String

    sRole = LCase(Trim(sNewRole))
    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(9, i).String)) = LCase(Trim(sItem)) Then
            oData.getCellByPosition(10, i).String = sRole
            If sRole = "trash" Then oData.getCellByPosition(11, i).String = ""
            Exit Sub
        End If
    Next i
End Sub

REM Sets the owner activity of a state ("" = general). Trash states stay global.
Sub SetCatalogScope(sItem As String, sNewScope As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(9, i).String)) = LCase(Trim(sItem)) Then
            If LCase(Trim(oData.getCellByPosition(10, i).String)) = "trash" Then Exit Sub
            oData.getCellByPosition(11, i).String = Trim(sNewScope)
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM ACTIVITY STATES (COLUMN I) AND ACCESS RULES (COLUMNS M..P)
REM =========================================================================

REM Returns the current activity state for an activity (column I).
Function GetActivityStateName(sActivity As String) As String
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sTarget As String

    oData = ThisComponent.Sheets.getByName("data")
    sTarget = LCase(Trim(sActivity))
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(7, i).String)) = sTarget Then
            Dim sVal As String
            sVal = Trim(oData.getCellByPosition(8, i).String)
            ' "active" is equivalent to an empty activity state.
            If LCase(sVal) = "active" Then sVal = ""
            GetActivityStateName = sVal
            Exit Function
        End If
    Next i
    GetActivityStateName = ""
End Function

REM Sets the current activity state for an activity (column I).
Sub SetActivityState(sActivity As String, sState As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(7, i).String)) = LCase(Trim(sActivity)) Then
            Dim sVal As String
            sVal = Trim(sState)
            If LCase(sVal) = "active" Then sVal = ""
            oData.getCellByPosition(8, i).String = sVal
            Exit Sub
        End If
    Next i
End Sub

REM Returns the rule of an activity state. Columns: M=name, N=type, O=param, P=message.
Function GetActStateRule(sState As String, ByRef sType As String, ByRef sParam As String, ByRef sMessage As String) As Boolean
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long

    sType = ""
    sParam = ""
    sMessage = ""
    GetActStateRule = False

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(12, i).String)) = LCase(Trim(sState)) Then
            sType = LCase(Trim(oData.getCellByPosition(13, i).String))
            sParam = Trim(oData.getCellByPosition(14, i).String)
            sMessage = Trim(oData.getCellByPosition(15, i).String)
            GetActStateRule = True
            Exit Function
        End If
    Next i
End Function

REM Parses a "YYYY-MM-DD" date into a date serial (0 = invalid).
Function ParseIsoDate(s As String) As Double
    Dim parts() As String

    ParseIsoDate = 0
    If Trim(s) = "" Then Exit Function
    parts = Split(Trim(s), "-")
    If UBound(parts) < 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function
    ParseIsoDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
End Function

REM Returns True when the activity access is blocked by its current state.
Function IsActivityBlocked(sActivity As String, ByRef sMessage As String, ByRef sType As String) As Boolean
    Dim sState As String, sParam As String

    sMessage = ""
    sType = ""
    IsActivityBlocked = False

    sState = GetActivityStateName(sActivity)
    If sState = "" Then Exit Function
    If Not GetActStateRule(sState, sType, sParam, sMessage) Then Exit Function

    Select Case LCase(Trim(sType))
        Case "manual"
            IsActivityBlocked = True
        Case "date-from"
            If ParseIsoDate(sParam) > 0 Then IsActivityBlocked = (Date() < ParseIsoDate(sParam))
        Case "date-until"
            If ParseIsoDate(sParam) > 0 Then IsActivityBlocked = (Date() > ParseIsoDate(sParam))
    End Select
End Function

REM Reverts the activity filter (A2) to the last valid activity.
Sub RevertActivityFilter()
    Dim oVisor As Object

    oVisor = ThisComponent.Sheets.getByName("view")
    G_GateBusy = True
    oVisor.getCellByPosition(0, 1).String = G_LastActivity
    G_GateBusy = False
End Sub

REM Sets the activity filter without firing the access gate (used by seeding).
Sub SetActivityFilterSilently(sActivity As String)
    Dim oVisor As Object

    oVisor = ThisComponent.Sheets.getByName("view")
    G_GateBusy = True
    oVisor.getCellByPosition(0, 1).String = sActivity
    G_GateBusy = False
End Sub

REM Validates access to an activity; may resolve a manual condition or revert A2.
Function ValidateActivityAccess(sActivity As String) As Boolean
    Dim sType As String, sMessage As String

    If Trim(sActivity) = "" Then
        ValidateActivityAccess = True
        Exit Function
    End If

    If Not IsActivityBlocked(sActivity, sMessage, sType) Then
        G_LastActivity = sActivity
        ValidateActivityAccess = True
        Exit Function
    End If

    If LCase(Trim(sType)) = "manual" Then
        If MsgBox(sMessage & Chr(13) & Chr(13) & _
                  "Mark the condition as fulfilled and return to active?", 4 + 32, "Activity Locked") = 6 Then
            Call SetActivityState(sActivity, "")
            G_LastActivity = sActivity
            ValidateActivityAccess = True
            Exit Function
        End If
    Else
        MsgBox sMessage, 48, "Activity Locked"
    End If

    Call RevertActivityFilter()
    ValidateActivityAccess = False
End Function

REM Returns the activity (column A) of a task identified by its ID.
Function GetTaskActivityById(oData As Object, IdTask As Long) As String
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    For i = 1 To TotalRows - 1
        If CLng(oData.getCellByPosition(5, i).Value) = IdTask Then
            GetTaskActivityById = Trim(oData.getCellByPosition(0, i).String)
            Exit Function
        End If
    Next i
    GetTaskActivityById = ""
End Function

REM =========================================================================
REM STARTUP ACCESS SERVICE + BOARD STATE FILTER
REM =========================================================================

REM True when the activity exists in the catalog.
Function ActivityExists(sActivity As String) As Boolean
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim t As String

    oData = ThisComponent.Sheets.getByName("data")
    t = LCase(Trim(sActivity))
    ActivityExists = False
    If t = "" Then Exit Function

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(7, i).String)) = t Then
            ActivityExists = True
            Exit Function
        End If
    Next i
End Function

REM Returns the first accessible activity (catalog order), or "".
Function FindAccessibleActivity() As String
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sAct As String, sType As String, sMsg As String

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    FindAccessibleActivity = ""

    For i = 1 To TotalRows - 1
        sAct = Trim(oData.getCellByPosition(7, i).String)
        If sAct <> "" Then
            If Not IsActivityBlocked(sAct, sMsg, sType) Then
                FindAccessibleActivity = sAct
                Exit Function
            End If
        End If
    Next i
End Function

REM Ensures A2 points to an accessible activity (startup and live changes).
Sub EnsureAccessibleActivity(bInteractive As Boolean)
    Dim oVisor As Object
    Dim sCur As String, sNext As String
    Dim sType As String, sMsg As String

    oVisor = ThisComponent.Sheets.getByName("view")
    sCur = Trim(oVisor.getCellByPosition(0, 1).String)

    If sCur <> "" And ActivityExists(sCur) Then
        If Not IsActivityBlocked(sCur, sMsg, sType) Then
            G_LastActivity = sCur
            Exit Sub
        End If
    End If

    sNext = FindAccessibleActivity()
    If sNext <> "" Then
        If bInteractive And sCur <> "" Then
            MsgBox "The activity """ & sCur & """ cannot be accessed in its current state." & Chr(13) & _
                   "Switching to """ & sNext & """.", 48, "Activity Access"
        End If
        G_GateBusy = True
        oVisor.getCellByPosition(0, 1).String = sNext
        G_GateBusy = False
        G_LastActivity = sNext
        Call DrawingBoard()
        Exit Sub
    End If

    MsgBox "No activity is currently accessible. Please change an activity's state in the manager.", 48, "Activity Access"
    Call ManageActivityes()
    sNext = FindAccessibleActivity()
    If sNext <> "" Then
        G_GateBusy = True
        oVisor.getCellByPosition(0, 1).String = sNext
        G_GateBusy = False
        G_LastActivity = sNext
        Call DrawingBoard()
    End If
End Sub

REM Builds the board state list: general states + the current activity's states.
Function BuildBoardStateList(sActivity As String) As Variant
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vList() As Variant
    Dim sState As String, sRole As String, sScope As String

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vList(TotalRows)
    n = 0

    For i = 1 To TotalRows - 1
        sState = Trim(oData.getCellByPosition(9, i).String)
        sRole = LCase(Trim(oData.getCellByPosition(10, i).String))
        sScope = LCase(Trim(oData.getCellByPosition(11, i).String))
        If sState <> "" And sRole <> "trash" Then
            If sScope = "" Or sScope = LCase(Trim(sActivity)) Then
                vList(n) = sState
                n = n + 1
            End If
        End If
    Next i

    If n = 0 Then
        ReDim vList(0)
        vList(0) = ""
    Else
        ReDim Preserve vList(n - 1)
    End If
    BuildBoardStateList = vList()
End Function

REM True when the value is in the list (empty is treated as valid).
Function ValueInArray(vList As Variant, s As String) As Boolean
    Dim i As Long, t As String

    t = LCase(Trim(s))
    ValueInArray = False
    If t = "" Then
        ValueInArray = True
        Exit Function
    End If
    For i = 0 To UBound(vList)
        If LCase(Trim(CStr(vList(i)))) = t Then
            ValueInArray = True
            Exit Function
        End If
    Next i
End Function



REM Sets a list validation pointing at the helper range in column Q.
Sub ApplyBoardStateList(oCell As Object)
    Dim oVal As Object

    oVal = oCell.Validation
    oVal.Type = com.sun.star.sheet.ValidationType.LIST
    oVal.Formula1 = "$data.$Q$2:$Q$20"
    oVal.ShowList = com.sun.star.sheet.TableValidationVisibility.UNSORTED
    oVal.IgnoreBlankCells = True
    oVal.ShowErrorMessage = False
    oVal.ShowInputMessage = False
    oCell.Validation = oVal
End Sub

REM Rebuilds the filtered board state list (helper column Q) and A3/B3 validations.
REM Returns the two default board states (general base states, with fallback).
Sub GetDefaultBoardStates(ByRef sFirst As String, ByRef sSecond As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim nBase As Long, nGen As Long
    Dim vBase() As String, vGen() As String
    Dim sState As String, sRole As String, sScope As String

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vBase(TotalRows)
    ReDim vGen(TotalRows)
    nBase = 0
    nGen = 0

    For i = 1 To TotalRows - 1
        sState = Trim(oData.getCellByPosition(9, i).String)
        sRole = LCase(Trim(oData.getCellByPosition(10, i).String))
        sScope = Trim(oData.getCellByPosition(11, i).String)
        If sState <> "" And sScope = "" And sRole <> "trash" Then
            vGen(nGen) = sState
            nGen = nGen + 1
            If sRole = "base" Then
                vBase(nBase) = sState
                nBase = nBase + 1
            End If
        End If
    Next i

    sFirst = ""
    sSecond = ""
    If nBase >= 2 Then
        sFirst = vBase(0)
        sSecond = vBase(1)
    ElseIf nBase = 1 Then
        sFirst = vBase(0)
        For i = 0 To nGen - 1
            If LCase(vGen(i)) <> LCase(sFirst) Then
                sSecond = vGen(i)
                Exit For
            End If
        Next i
    Else
        If nGen >= 2 Then
            sFirst = vGen(0)
            sSecond = vGen(1)
        ElseIf nGen = 1 Then
            sFirst = vGen(0)
        End If
    End If
End Sub

Sub RefreshBoardStateFilters()
    Dim oData As Object, oVisor As Object
    Dim TotalRows As Long, i As Long, nMax As Long
    Dim vList As Variant
    Dim sAct As String, sDefA As String, sDefB As String

    oData = ThisComponent.Sheets.getByName("data")
    oVisor = ThisComponent.Sheets.getByName("view")
    sAct = Trim(oVisor.getCellByPosition(0, 1).String)
    vList = BuildBoardStateList(sAct)

    nMax = GetCatalogMaxRow("State")
    For i = 1 To nMax
        oData.getCellByPosition(16, i).String = ""
    Next i
    For i = 0 To UBound(vList)
        If Trim(CStr(vList(i))) <> "" Then
            oData.getCellByPosition(16, i + 1).String = CStr(vList(i))
        End If
    Next i

    Call ApplyBoardStateList(oVisor.getCellByPosition(0, 2))
    Call ApplyBoardStateList(oVisor.getCellByPosition(1, 2))

    Call GetDefaultBoardStates(sDefA, sDefB)

    G_BoardFilterBusy = True
    If Not ValueInArray(vList, Trim(oVisor.getCellByPosition(0, 2).String)) Then
        oVisor.getCellByPosition(0, 2).String = sDefA
    End If
    If Not ValueInArray(vList, Trim(oVisor.getCellByPosition(1, 2).String)) Then
        oVisor.getCellByPosition(1, 2).String = sDefB
    End If
    ' Avoid identical columns when possible.
    If Trim(oVisor.getCellByPosition(0, 2).String) <> "" Then
        If LCase(Trim(oVisor.getCellByPosition(0, 2).String)) = LCase(Trim(oVisor.getCellByPosition(1, 2).String)) Then
            If LCase(Trim(sDefB)) <> LCase(Trim(oVisor.getCellByPosition(0, 2).String)) Then
                oVisor.getCellByPosition(1, 2).String = sDefB
            ElseIf LCase(Trim(sDefA)) <> LCase(Trim(oVisor.getCellByPosition(0, 2).String)) Then
                oVisor.getCellByPosition(1, 2).String = sDefA
            End If
        End If
    End If
    G_BoardFilterBusy = False
End Sub

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
    
    ' 1. Security validation for Base / Trash States (role read from the matrix)
    If sTypeItem = "State" Then
        If IsBaseState(sItem) Then
            MsgBox " The state """ & sItem & """ is marked as a base/trash state and cannot be deleted." & Chr(13) & _
                   "Change its role to normal first if you really need to remove it.", 48, "Protected State"
            Exit Sub
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
        
        If LCase(Trim(sValueCompare)) = LCase(Trim(sItem)) Then
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
        
        ' A. Clear the task cells instead of removing rows.
        '    Physical row removal would also shift the catalog cells stored
        '    in columns H/J on the same rows.
        Dim nCompareCol As Long
        If sTypeItem = "Activity" Then
            nCompareCol = 0
        Else
            nCompareCol = 2
        End If

        Dim nTaskCol As Long
        For i = TotalRows - 1 To 1 Step -1
            If LCase(Trim(oData.getCellByPosition(nCompareCol, i).String)) = LCase(Trim(sItem)) Then
                For nTaskCol = 0 To 5
                    oData.getCellByPosition(nTaskCol, i).String = ""
                Next nTaskCol
            End If
        Next i

        ' B. Delete the item from the configuration catalog and compact it.
        Dim nCatCol As Long, nCatMaxRow As Long
        nCatCol = GetCatalogColumn(sTypeItem)
        nCatMaxRow = GetCatalogMaxRow(sTypeItem)
        For i = 1 To nCatMaxRow
            If LCase(Trim(oData.getCellByPosition(nCatCol, i).String)) = LCase(Trim(sItem)) Then
                oData.getCellByPosition(nCatCol, i).String = ""
                If sTypeItem = "State" Then
                    oData.getCellByPosition(GetRoleColumn(), i).String = ""
                    oData.getCellByPosition(11, i).String = ""
                End If
                Exit For
            End If
        Next i
        Call CompactCatalog(sTypeItem)

        ' B2. When deleting an activity, also remove its per-activity states.
        If sTypeItem = "Activity" Then
            Dim nStateMaxRow As Long
            nStateMaxRow = GetCatalogMaxRow("State")
            For i = 1 To nStateMaxRow
                If LCase(Trim(oData.getCellByPosition(11, i).String)) = LCase(Trim(sItem)) Then
                    oData.getCellByPosition(9, i).String = ""
                    oData.getCellByPosition(10, i).String = ""
                    oData.getCellByPosition(11, i).String = ""
                End If
            Next i
            Call CompactCatalog("State")
        End If

        ' C. Reorder and redraw the Kanban board
        If sTypeItem = "Activity" Then Call EnsureAccessibleActivity(True)
        Call RefreshBoardStateFilters()
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

    ' Remember the current activity so the access gate can revert A2.
    G_LastActivity = Trim(oSheetHeard.getCellByPosition(0, 1).String)
    G_GateBusy = False
    G_BoardFilterBusy = False

    ' Startup verification: ensure the visible activity is accessible and
    ' the board state filters are in sync.
    Call EnsureAccessibleActivity(False)
    Call RefreshBoardStateFilters()

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
    Dim oVisor As Object
    Dim sNewActivity As String
    
    On Error Resume Next

    ' Ignore the programmatic updates done by the gate and filter services.
    If G_GateBusy Or G_BoardFilterBusy Then Exit Sub
    
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

        ' Gate the activity filter (A2) before redrawing.
        If nRow = 1 Then
            oVisor = ThisComponent.Sheets.getByName("view")
            sNewActivity = Trim(oVisor.getCellByPosition(0, 1).String)
            If Not ValidateActivityAccess(sNewActivity) Then
                ' A2 was reverted; redraw the previous activity.
                binFrozenFocus = True
                Call CustomDataSort()
                binFrozenFocus = False
                Exit Sub
            End If
            ' Refresh the board state lists for the new activity.
            Call RefreshBoardStateFilters()
        End If
       
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
        ' Determine the activity context so per-activity states can be filtered.
        Dim sActCtx As String : sActCtx = ""
        If IdActive > 0 Then
            sActCtx = GetTaskActivityById(oData, IdActive)
        ElseIf oModel.hasByName("CmbActivity") Then
            sActCtx = Trim(oDialogue.getControl("CmbActivity").Text)
        End If
        ' We pass the model and the view separately to the external macro
        Call ServPopulateCmbState(oModel.getByName("CmbState"), oDialogue.getControl("CmbState"), oData, IdActive, sActCtx)
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
        Dim sTypeItem As String : sTypeItem = G_ManagedItemType

        ' Populate the selector list. For states, the scope filter comes first.
        If sTypeItem = "State" Then
            Call PopulateScopeFilter(oDialogue)
            Call ServPopulateCmbItemFiltered(oDialogue, oData, "General")
        Else
            Call ServPopulateCmbItem(oModel.getByName("CmbItem"), oDialogue.getControl("CmbItem"), oData, sTypeItem)
        End If

        ' Initialize the management controls (no item selected yet).
        G_UpdatingPanel = True
        G_SelectedItem = ""
        G_SelectedPos = 0
        G_SelectedRole = ""
        G_SelectedScope = ""
        G_SelectedActState = ""
        If oModel.hasByName("TxtNewName") Then
            oDialogue.getControl("TxtNewName").Text = ""
        End If
        If oModel.hasByName("CmbPos") Then
            oDialogue.getControl("CmbPos").setText("")
            oDialogue.getControl("CmbPos").Model.Enabled = False
        End If
        If oModel.hasByName("CmbRole") Then
            oDialogue.getControl("CmbRole").setVisible(sTypeItem = "State")
            oDialogue.getControl("CmbRole").setText("")
            oDialogue.getControl("CmbRole").Model.Enabled = False
        End If
        If oModel.hasByName("CmbScope") Then
            oDialogue.getControl("CmbScope").setVisible(sTypeItem = "State")
        End If
        If oModel.hasByName("LblScope") Then
            oDialogue.getControl("LblScope").setVisible(sTypeItem = "State")
        End If
        If oModel.hasByName("LblRole") Then
            oDialogue.getControl("LblRole").setVisible(sTypeItem = "State")
        End If
        If oModel.hasByName("CmbActState") Then
            oDialogue.getControl("CmbActState").setVisible(sTypeItem = "Activity")
            oDialogue.getControl("CmbActState").setText("")
            oDialogue.getControl("CmbActState").Model.Enabled = False
        End If
        If oModel.hasByName("LblActState") Then
            oDialogue.getControl("LblActState").setVisible(sTypeItem = "Activity")
        End If
        G_UpdatingPanel = False
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
Sub ServPopulateCmbState(oModelCombo As Object, oControlView As Object, oData As Object, IdActive As Long, sActivity As String)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long
    Dim sValue As String, sStateSearch As String
    Dim sScope As String, sRole As String
    Dim vlist() As String, nCounter As Long : nCounter = 0
    Dim idxState As Long : idxState = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Load data into the array
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vlist(TotalRows)
    
    Dim sTrashState As String
    sTrashState = GetTrashStateName()
    For i = 1 To TotalRows - 1
        sValue = Trim(oData.getCellByPosition(9, i).String)
        sRole = LCase(Trim(oData.getCellByPosition(10, i).String))
        sScope = LCase(Trim(oData.getCellByPosition(11, i).String))
        If sValue <> "" And LCase(sValue) <> LCase(sTrashState) And sRole <> "trash" Then
            If sScope = "" Or sScope = LCase(Trim(sActivity)) Then
                vlist(nCounter) = sValue
                nCounter = nCounter + 1
            End If
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

REM Repopulates the task state list when the activity changes (scope filtering).
Sub CmbActivityChanged(oEvent As Object)
    Dim oDialogue As Object, oData As Object
    Dim sActivity As String

    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    sActivity = Trim(oDialogue.getControl("CmbActivity").Text)

    Call ServPopulateCmbState(oDialogue.getControl("CmbState").Model, oDialogue.getControl("CmbState"), oData, 0, sActivity)
    oDialogue.getControl("CmbState").setText("")
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
    Dim sTrashState As String
    sTrashState = GetTrashStateName()
    
    For i = 1 To TotalRows - 1
        sActRow = Trim(oData.getCellByPosition(0, i).String)
        sEstRow = Trim(oData.getCellByPosition(2, i).String)
        
        ' GOLDEN RULE: Activity and State must match, and must NOT be the trash state or empty
        If sActRow = sActivityTask And sEstRow = sStateTask And LCase(sEstRow) <> LCase(sTrashState) And sActRow <> "" Then
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
    sTitle = WatermarkValue(oControlView)
    
    ' Write the value in Column B (index 1) of the corresponding row
    oData.getCellByPosition(1, nRowDestiny).String = sTitle
End Sub

REM =========================================================================
REM EXTERNAL SERVICE 4: SAVE DESCRIPTION TO THE DATABASE
REM =========================================================================
Sub ServSaveTxtDescription(oControlView As Object, oData As Object, nRowDestiny As Long)
    Dim sDescription As String
    
    ' Extract the text from the description box (supports multi-line formatting)
    sDescription = WatermarkValue(oControlView)
    
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
        Dim sTrashState As String
        sTrashState = GetTrashStateName()
        If sTrashState = "" Then
            MsgBox "No trash state is defined. Mark a state with the 'trash' role first.", 48, "Trash State Missing"
            Exit Sub
        End If

        ' Block screen update
        oDoc.addActionLock()
        
        ' Change the State to the configured trash state in Column C (index 2)
        oData.getCellByPosition(2, DataRowIdx).String = sTrashState
        
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
        
        ' Clear the task cells instead of removing the row, so the catalog
        ' cells stored in columns H/J on the same row are preserved.
        Dim nClearCol As Long
        For nClearCol = 0 To 5
            oData.getCellByPosition(nClearCol, DataRowIdx).String = ""
        Next nClearCol
        
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
    Call ApplyWatermarks(oDialogue)
    
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
    Call ApplyWatermarks(oDialogue)
    
    ' 3. Smart pre-selection based on the Viewer's active filters
    sActivityFilter = Trim(oVisor.getCellByPosition(0, 1).String) ' Cell A2
    sStateFilter = Trim(oVisor.getCellByPosition(0, 2).String)    ' Cell A3
    
    If sActivityFilter <> "" Then
        oDialogue.getControl("CmbActivity").setText(sActivityFilter)
        ' Ensure the state list matches the activity even if the event did not fire.
        Call ServPopulateCmbState(oDialogue.getControl("CmbState").Model, oDialogue.getControl("CmbState"), oData, 0, sActivityFilter)
    End If
    
    If sStateFilter <> "" Then
        oDialogue.getControl("CmbState").setText(sStateFilter)
    Else
        oDialogue.getControl("CmbState").setText("In progress")
    End If
    
    ' 4. Execution of the dialogue
    If oDialogue.execute() = 1 Then
        ' Validation of mandatory fields
        If WatermarkValue(oDialogue.getControl("TxtTitle")) = "" Then
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
    Dim sDisplay As String

    ' Reset the management panel state.
    G_SelectedItem = ""
    G_SelectedPos = 0
    G_SelectedRole = ""
    G_UpdatingPanel = False
    
    ' 1. Instantiate the unified dialog "DialogPanelItem"
    DialogLibraries.LoadLibrary("Standard")
    oDialogue = CreateUnoDialog(DialogLibraries.Standard.DialogPanelItem)
    
    ' Configure the title dynamically based on the global context
    If G_ManagedItemType = "Activity" Then
        sDisplay = "Activities"
    Else
        sDisplay = "States"
    End If
    oDialogue.Title = "Manage " & sDisplay
    
    ' 2. Call the Reading Hub (ID = 0) to populate the "CmbItem" ComboBox.
    Call InitializeDialogServices(oDialogue, 0)
    Call ApplyWatermarks(oDialogue)
    
    ' Initially disable the Accept button until the user types or selects something.
    oDialogue.getControl("BtnAccept").Model.Enabled = False
    oDialogue.getControl("BtnAccept").Model.Label = "Accept"
    oDialogue.getControl("BtnAccept").Model.BackgroundColor = RGB(240, 240, 240)
    oDialogue.getControl("BtnAccept").Model.TextColor = RGB(0, 0, 0)
    
    ' 3. Expand on-screen panel
    oDialogue.execute()
    oDialogue.dispose()
End Sub

' --------------------------------------------------------------------------

REM =========================================================================
REM CATALOG WRITING HUB: ADD, DELETE OR MODIFY BASED ON CONTEXT
REM =========================================================================
Sub SaveCatalogFromDialog(oEvent As Object)
    Dim oDialogue As Object, oCmb As Object
    Dim sItem As String, sNewName As String, sPos As String, sRoleLabel As String
    Dim sRole As String, sTrash As String, sAddScope As String
    Dim sAddRole As String, sAddActState As String
    Dim sActStateLabel As String, sActState As String
    Dim vList() As String
    Dim bExist As Boolean, bName As Boolean, bPos As Boolean, bRole As Boolean
    Dim bActState As Boolean
    Dim nPos As Long, nCount As Long, nAddPos As Long
    Dim answer As Integer

    oDialogue = oEvent.Source.Context
    oCmb = oDialogue.getControl("CmbItem")
    sItem = Trim(oCmb.Text)

    If sItem = "" Then Exit Sub

    vList() = oCmb.Model.StringItemList
    bExist = ElementExistInList(sItem, vList())

    If Not bExist Then
        ' OPERATION: ADD NEW ELEMENT (position, role and activity state apply).
        sAddScope = ""
        sAddRole = ""
        sAddActState = ""
        nAddPos = 0
        If G_ManagedItemType = "State" Then
            sAddScope = GetSelectedScopeValue(oDialogue)
            sAddRole = RoleLabelToValue(Trim(oDialogue.getControl("CmbRole").Text))
            sPos = Trim(oDialogue.getControl("CmbPos").Text)
            If sPos <> "" And UCase(sPos) <> "F" And IsNumeric(sPos) Then nAddPos = CLng(sPos)
            If LCase(sAddRole) = "trash" Then
                sTrash = GetTrashStateName()
                If sTrash <> "" Then
                    MsgBox "A trash state already exists: """ & sTrash & """. Set it back to normal first.", 48, "State Role"
                    Exit Sub
                End If
            End If
        Else
            sPos = Trim(oDialogue.getControl("CmbPos").Text)
            If sPos <> "" And UCase(sPos) <> "F" And IsNumeric(sPos) Then nAddPos = CLng(sPos)
            sAddActState = ActivityStateLabelToValue(Trim(oDialogue.getControl("CmbActState").Text))
        End If
        answer = MsgBox("Do you want to add the new " & G_ManagedItemType & ": """ & sItem & """?", 4 + 32, "Confirm Addition")
        If answer = 6 Then
            oDialogue.endExecute()
            Call AddCatalogItem(sItem, G_ManagedItemType, sAddScope, nAddPos, sAddRole, sAddActState)
        End If
        Exit Sub
    End If

    ' Existing item: determine the pending changes.
    nCount = CountCatalog(G_ManagedItemType)
    sNewName = WatermarkValue(oDialogue.getControl("TxtNewName"))
    bName = (sNewName <> "" And LCase(sNewName) <> LCase(sItem))

    sPos = Trim(oDialogue.getControl("CmbPos").Text)
    nPos = ResolvePosition(sPos, nCount)
    bPos = (nPos <> G_SelectedPos)

    bRole = False
    sRole = ""
    bActState = False
    sActState = ""
    If G_ManagedItemType = "State" Then
        sRoleLabel = Trim(oDialogue.getControl("CmbRole").Text)
        sRole = RoleLabelToValue(sRoleLabel)
        bRole = (sRole <> G_SelectedRole)
    Else
        sActStateLabel = Trim(oDialogue.getControl("CmbActState").Text)
        sActState = ActivityStateLabelToValue(sActStateLabel)
        bActState = (LCase(sActState) <> LCase(G_SelectedActState))
    End If

    If Not bName And Not bPos And Not bRole And Not bActState Then
        ' OPERATION: CASCADE ELIMINATION
        oDialogue.endExecute()
        Call RunCascadeDelete(sItem, G_ManagedItemType)
        Exit Sub
    End If

    ' OPERATION: MODIFY (rename / position / role / activity state)
    If bName Then
        answer = MsgBox("Rename " & G_ManagedItemType & ": """ & sItem & """ -> """ & sNewName & """?", 4 + 32, "Confirm Rename")
        If answer <> 6 Then Exit Sub
    End If

    If bRole And LCase(sRole) = "trash" Then
        sTrash = GetTrashStateName()
        If sTrash <> "" And LCase(sTrash) <> LCase(sItem) And LCase(sTrash) <> LCase(sNewName) Then
            MsgBox "A trash state already exists: """ & sTrash & """. Set it back to normal first.", 48, "State Role"
            Exit Sub
        End If
    End If

    oDialogue.endExecute()
    Call ModifyCatalogItem(sItem, sNewName, bName, bPos, nPos, bRole, sRole, False, "", bActState, sActState, G_ManagedItemType)
End Sub

REM =========================================================================
REM ACTIVITY-STATE CATALOG PANEL (COLUMNS M..P)
REM =========================================================================

REM Opens the activity-state catalog editor.
Sub ManageActivityStates()
    Dim oDialogue As Object
    Dim oData As Object

    DialogLibraries.LoadLibrary("Standard")
    oDialogue = CreateUnoDialog(DialogLibraries.Standard.DialogPanelActState)
    oData = ThisComponent.Sheets.getByName("data")

    G_UpdatingPanel = True
    Call ServPopulateActStatesCombo(oDialogue, oData)
    Dim vTypes(3) As String
    vTypes(0) = "none"
    vTypes(1) = "manual"
    vTypes(2) = "date-from"
    vTypes(3) = "date-until"
    oDialogue.getControl("CmbType").Model.StringItemList = vTypes()
    oDialogue.getControl("TxtNewName").Text = ""
    oDialogue.getControl("CmbType").setText("")
    oDialogue.getControl("TxtParam").Text = ""
    oDialogue.getControl("TxtMessage").Text = ""
    oDialogue.getControl("BtnAccept").Model.Enabled = False
    oDialogue.getControl("BtnAccept").Model.Label = "Accept"
    oDialogue.getControl("BtnAccept").Model.BackgroundColor = RGB(240, 240, 240)
    oDialogue.getControl("BtnAccept").Model.TextColor = RGB(0, 0, 0)
    G_UpdatingPanel = False
    Call ApplyWatermarks(oDialogue)

    oDialogue.Title = "Manage Activity States"
    oDialogue.execute()
    oDialogue.dispose()
End Sub

REM Fills CmbItem with the activity-state names (column M).
Sub ServPopulateActStatesCombo(oDialogue As Object, oData As Object)
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long, n As Long
    Dim vList() As String

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    ReDim vList(TotalRows)

    n = 0
    For i = 1 To TotalRows - 1
        If Trim(oData.getCellByPosition(12, i).String) <> "" Then
            vList(n) = Trim(oData.getCellByPosition(12, i).String)
            n = n + 1
        End If
    Next i
    If n > 0 Then
        ReDim Preserve vList(n - 1)
        oDialogue.getControl("CmbItem").Model.StringItemList = vList()
    Else
        Dim vEmpty(0) As String
        oDialogue.getControl("CmbItem").Model.StringItemList = vEmpty()
    End If
End Sub

REM Returns the activity-state catalog row for a name (-1 if missing).
Function FindActStateRow(oData As Object, sName As String) As Long
    Dim oCursor As Object
    Dim TotalRows As Long, i As Long

    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    FindActStateRow = -1
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(12, i).String)) = LCase(Trim(sName)) Then
            FindActStateRow = i
            Exit Function
        End If
    Next i
End Function

REM True when an activity currently uses the given activity-state.
Function ActivityStateIsUsed(sState As String) As Boolean
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1

    ActivityStateIsUsed = False
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(8, i).String)) = LCase(Trim(sState)) Then
            ActivityStateIsUsed = True
            Exit Function
        End If
    Next i
End Function

REM Handles selection changes inside the activity-state editor.
Sub CmbActStateItemChanged(oEvent As Object)
    Dim oDialogue As Object, oData As Object
    Dim sText As String
    Dim vList() As String
    Dim bExist As Boolean
    Dim nRow As Long

    If G_UpdatingPanel Then Exit Sub

    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    sText = Trim(oEvent.Source.Text)

    G_UpdatingPanel = True
    oDialogue.getControl("TxtNewName").Text = ""
    G_UpdatingPanel = False

    vList() = oDialogue.getControl("CmbItem").Model.StringItemList
    bExist = ElementExistInList(sText, vList())

    nRow = -1
    If bExist Then nRow = FindActStateRow(oData, sText)

    G_UpdatingPanel = True
    If nRow > 0 Then
        oDialogue.getControl("CmbType").setText(Trim(oData.getCellByPosition(13, nRow).String))
        oDialogue.getControl("TxtParam").Text = Trim(oData.getCellByPosition(14, nRow).String)
        oDialogue.getControl("TxtMessage").Text = Trim(oData.getCellByPosition(15, nRow).String)
    Else
        oDialogue.getControl("CmbType").setText("")
        oDialogue.getControl("TxtParam").Text = ""
        oDialogue.getControl("TxtMessage").Text = ""
    End If
    G_UpdatingPanel = False
    Call ApplyWatermarks(oDialogue)

    Call UpdateActStatePanelButton(oEvent)
End Sub

REM Recomputes the action button of the activity-state editor.
Sub UpdateActStatePanelButton(oEvent As Object)
    Dim oDialogue As Object, oBtn As Object, oData As Object
    Dim sText As String, sNewName As String, sType As String, sParam As String, sMessage As String
    Dim vList() As String
    Dim bExist As Boolean, bChanged As Boolean
    Dim nRow As Long

    If G_UpdatingPanel Then Exit Sub

    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    oBtn = oDialogue.getControl("BtnAccept").Model
    sText = Trim(oDialogue.getControl("CmbItem").Text)

    If sText = "" Then
        oBtn.Enabled = False
        oBtn.Label = "Accept"
        oBtn.BackgroundColor = RGB(240, 240, 240)
        oBtn.TextColor = RGB(0, 0, 0)
        Exit Sub
    End If

    vList() = oDialogue.getControl("CmbItem").Model.StringItemList
    bExist = ElementExistInList(sText, vList())

    If Not bExist Then
        oBtn.Enabled = True
        oBtn.Label = "[+] Add"
        oBtn.BackgroundColor = RGB(40, 167, 69)
        oBtn.TextColor = RGB(255, 255, 255)
        Exit Sub
    End If

    bChanged = False
    Dim bName As Boolean, bType As Boolean, bParam As Boolean, bMessage As Boolean
    Dim sIcons As String
    bName = False
    bType = False
    bParam = False
    bMessage = False
    sNewName = WatermarkValue(oDialogue.getControl("TxtNewName"))
    If sNewName <> "" And LCase(sNewName) <> LCase(sText) Then
        bName = True
        bChanged = True
    End If
    nRow = FindActStateRow(oData, sText)
    If nRow > 0 Then
        sType = Trim(oDialogue.getControl("CmbType").Text)
        sParam = WatermarkValue(oDialogue.getControl("TxtParam"))
        sMessage = WatermarkValue(oDialogue.getControl("TxtMessage"))
        If LCase(sType) <> LCase(Trim(oData.getCellByPosition(13, nRow).String)) Then
            bType = True
            bChanged = True
        End If
        If sParam <> Trim(oData.getCellByPosition(14, nRow).String) Then
            bParam = True
            bChanged = True
        End If
        If sMessage <> Trim(oData.getCellByPosition(15, nRow).String) Then
            bMessage = True
            bChanged = True
        End If
    End If

    If bChanged Then
        sIcons = ""
        If bName Then sIcons = sIcons & "[✎]"
        If bType Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[◈]"
        End If
        If bParam Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[⏱]"
        End If
        If bMessage Then
            If sIcons <> "" Then sIcons = sIcons & "&"
            sIcons = sIcons & "[💬]"
        End If
        oBtn.Enabled = True
        oBtn.Label = "Modify" & sIcons
        oBtn.BackgroundColor = RGB(0, 123, 255)
        oBtn.TextColor = RGB(255, 255, 255)
    Else
        oBtn.Enabled = True
        oBtn.Label = "⮿ Delete"
        oBtn.BackgroundColor = RGB(220, 53, 69)
        oBtn.TextColor = RGB(255, 255, 255)
    End If
End Sub

REM Saves (add / modify / delete) an activity-state from the editor.
Sub SaveActStateFromDialog(oEvent As Object)
    Dim oDialogue As Object, oData As Object
    Dim sText As String, sNewName As String, sType As String, sParam As String, sMessage As String
    Dim vList() As String
    Dim bExist As Boolean, bChanged As Boolean
    Dim nRow As Long
    Dim answer As Integer

    oDialogue = oEvent.Source.Context
    oData = ThisComponent.Sheets.getByName("data")
    sText = Trim(oDialogue.getControl("CmbItem").Text)
    If sText = "" Then Exit Sub

    vList() = oDialogue.getControl("CmbItem").Model.StringItemList
    bExist = ElementExistInList(sText, vList())

    sNewName = WatermarkValue(oDialogue.getControl("TxtNewName"))
    sType = LCase(Trim(oDialogue.getControl("CmbType").Text))
    sParam = WatermarkValue(oDialogue.getControl("TxtParam"))
    sMessage = WatermarkValue(oDialogue.getControl("TxtMessage"))
    If sNewName = "" Then sNewName = sText

    If Not bExist Then
        answer = MsgBox("Add the activity state """ & sNewName & """?", 4 + 32, "Confirm Addition")
        If answer = 6 Then
            oDialogue.endExecute()
            Call AddActStateItem(sNewName, sType, sParam, sMessage)
        End If
        Exit Sub
    End If

    ' "active" is a system state: it can be edited but not renamed.
    If LCase(sText) = "active" Then sNewName = sText

    bChanged = False
    If LCase(sNewName) <> LCase(sText) Then bChanged = True
    nRow = FindActStateRow(oData, sText)
    If nRow > 0 Then
        If sType <> LCase(Trim(oData.getCellByPosition(13, nRow).String)) Then bChanged = True
        If sParam <> Trim(oData.getCellByPosition(14, nRow).String) Then bChanged = True
        If sMessage <> Trim(oData.getCellByPosition(15, nRow).String) Then bChanged = True
    End If

    If bChanged Then
        answer = MsgBox("Update the activity state """ & sText & """?", 4 + 32, "Confirm Update")
        If answer = 6 Then
            oDialogue.endExecute()
            Call ModifyActStateItem(sText, sNewName, sType, sParam, sMessage)
        End If
    Else
        If LCase(sText) = "active" Then
            MsgBox "The ""active"" state is a system state and cannot be deleted.", 48, "Protected Activity State"
            Exit Sub
        End If
        answer = MsgBox("Delete the activity state """ & sText & """?", 4 + 16 + 256, "Confirm Deletion")
        If answer = 6 Then
            oDialogue.endExecute()
            Call DeleteActStateItem(sText)
        End If
    End If
End Sub

REM Adds a new activity-state entry to columns M..P.
Sub AddActStateItem(sName As String, sType As String, sParam As String, sMessage As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, nFree As Long

    oData = ThisComponent.Sheets.getByName("data")
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    If TotalRows < 2 Then TotalRows = 2

    nFree = -1
    For i = 1 To TotalRows
        If Trim(oData.getCellByPosition(12, i).String) = "" Then
            nFree = i
            Exit For
        End If
    Next i
    If nFree = -1 Then nFree = TotalRows

    oData.getCellByPosition(12, nFree).String = sName
    oData.getCellByPosition(13, nFree).String = sType
    oData.getCellByPosition(14, nFree).String = sParam
    oData.getCellByPosition(15, nFree).String = sMessage

    MsgBox "Activity state successfully added.", 64, "Success"
End Sub

REM Renames/updates an activity-state entry and keeps activity references in sync.
Sub ModifyActStateItem(sOld As String, sNew As String, sType As String, sParam As String, sMessage As String)
    Dim oData As Object, oCursor As Object
    Dim TotalRows As Long, i As Long, nRow As Long

    oData = ThisComponent.Sheets.getByName("data")
    nRow = FindActStateRow(oData, sOld)
    If nRow = -1 Then Exit Sub

    oData.getCellByPosition(12, nRow).String = sNew
    oData.getCellByPosition(13, nRow).String = sType
    oData.getCellByPosition(14, nRow).String = sParam
    oData.getCellByPosition(15, nRow).String = sMessage

    ' Keep the activity references (column I) in sync.
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(8, i).String)) = LCase(Trim(sOld)) Then
            oData.getCellByPosition(8, i).String = sNew
        End If
    Next i

    MsgBox "Activity state successfully updated.", 64, "Success"
End Sub

REM Deletes an activity-state entry from columns M..P.
Sub DeleteActStateItem(sName As String)
    Dim oData As Object, oCursor As Object
    Dim nRow As Long, i As Long, TotalRows As Long

    If LCase(Trim(sName)) = "active" Then Exit Sub

    oData = ThisComponent.Sheets.getByName("data")
    nRow = FindActStateRow(oData, sName)
    If nRow = -1 Then Exit Sub

    ' Reset any activity that currently uses this state back to active.
    oCursor = oData.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    TotalRows = oCursor.RangeAddress.EndRow + 1
    For i = 1 To TotalRows - 1
        If LCase(Trim(oData.getCellByPosition(8, i).String)) = LCase(Trim(sName)) Then
            oData.getCellByPosition(8, i).String = ""
        End If
    Next i

    oData.getCellByPosition(12, nRow).String = ""
    oData.getCellByPosition(13, nRow).String = ""
    oData.getCellByPosition(14, nRow).String = ""
    oData.getCellByPosition(15, nRow).String = ""

    MsgBox "Activity state successfully deleted.", 64, "Success"
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
