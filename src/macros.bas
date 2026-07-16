Global G_TipoItemGestionado As String
Global G_TextoAtajo As String
Global G_UsadoAtajo As Boolean
Global binFocoCongelado As Boolean
Global oSeleccionListener As Object
Global oHojaEscuchada As Object

Sub AlternarInterfaz
    Dim oDoc As Object
    Dim oLayout As Object
    
    oDoc = ThisComponent
    oLayout = oDoc.CurrentController.Frame.LayoutManager
    
    ' Si la interfaz está visible, la oculta. Si está oculta, la muestra.
    If oLayout.isVisible() Then
        oLayout.setVisible(False)
    Else
        oLayout.setVisible(True)
    End If
End Sub


REM =========================================================================
REM FASE 2 - MOTOR DE RENDERIZADO DATA-DRIVEN (REEMPLAZO DE MATRICIALES)
REM =========================================================================
Sub DibujarTablero()
    Dim oDoc As Object, oVisor As Object, oDatos As Object
    Dim sActividad As String, sEstadoA As String, sEstadoB As String
    Dim totalFilasDatos As Long, i As Long
    Dim filaActualA As Long, filaActualB As Long
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("visor")
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' 1. Desactivar actualización de pantalla para máxima velocidad
    oDoc.addActionLock()
    
    ' 2. Limpiar rangos de visualización antiguos y rangos de IDs ocultos (Filas 4 a 300 -> índices 3 a 299)
    ' Limpiamos columnas A, B, C, D (Tareas, Subtareas, Descripciones) y E, F, G (IDs ocultos)
    For i = 3 To 299
        oVisor.getCellByPosition(0, i).String = "" ' Col A
        oVisor.getCellByPosition(1, i).String = "" ' Col B
        oVisor.getCellByPosition(2, i).String = "" ' Col C
        oVisor.getCellByPosition(3, i).String = "" ' Col D
        oVisor.getCellByPosition(4, i).Value = 0   ' Col E (ID Oculto Tareas Col A)
        oVisor.getCellByPosition(5, i).Value = 0   ' Col F (ID Oculto Tareas Col B)
    Next i
    
    ' 3. Leer Filtros del Visor
    sActividad = Trim(oVisor.getCellByPosition(0, 1).String) ' A2 (Celda combinada A2:D2)
    sEstadoA = Trim(oVisor.getCellByPosition(0, 2).String)   ' A3 (Cabecera Estado 1)
    sEstadoB = Trim(oVisor.getCellByPosition(1, 2).String)   ' B3 (Cabecera Estado 2)
    
    If sActividad = "" Then
        oDoc.removeActionLock()
        Exit Sub
    End If
    
    ' 4. Determinar el tamaño de la Base de Datos
    Dim oCursor As Object
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilasDatos = oCursor.RangeAddress.EndRow + 1
    
    ' Inicializar punteros de escritura para el Visor (Empezamos en fila 4 -> índice 3)
    filaActualA = 3
    filaActualB = 3
    
    ' 5. Recorrer la Base de Datos e Inyectar en las Columnas correspondientes
    For i = 1 To totalFilasDatos - 1
        Dim sRegActividad As String, sRegTarea As String, sRegEstado As String
        Dim idTarea As Long
        
        sRegActividad = Trim(oDatos.getCellByPosition(0, i).String) ' Col A
        sRegTarea = Trim(oDatos.getCellByPosition(1, i).String)     ' Col B
        sRegEstado = Trim(oDatos.getCellByPosition(2, i).String)    ' Col C
        idTarea = CLng(oDatos.getCellByPosition(5, i).Value)        ' Col F (ID Único Inmutable)
        
        ' Si coincide con la actividad bajo consulta
        If sRegActividad = sActividad Then
            
            ' ¿Pertenece al primer Estado (Columna A del Visor)?
            If sRegEstado = sEstadoA And filaActualA <= 299 Then
                oVisor.getCellByPosition(0, filaActualA).String = sRegTarea
                oVisor.getCellByPosition(4, filaActualA).Value = idTarea ' Guarda ID en Columna E
                filaActualA = filaActualA + 1
            End If
                
            ' ¿Pertenece al segundo Estado (Columna B del Visor)?
            If sRegEstado = sEstadoB And filaActualB <= 299 Then
                oVisor.getCellByPosition(1, filaActualB).String = sRegTarea
                oVisor.getCellByPosition(5, filaActualB).Value = idTarea ' Guarda ID en Columna F
                filaActualB = filaActualB + 1
            End If
            
        End If
    Next i
    
    ' 6. Liberar bloqueo de pantalla y refrescar
    oDoc.removeActionLock()
    oDoc.calculateAll()
    
' ---- REPARACIÓN EN FASE 2 PRO: SINCRONIZAR DESCRIPCIÓN AL REDIBUJAR ----
    Dim idFocoActual As Long
    Dim oSelActual As Object
    
    oSelActual = oDoc.CurrentSelection
    If oSelActual.supportsService("com.sun.star.sheet.SheetCell") Then
        ' REUTILIZACIÓN ESTRELLA N° 6: Leer el ID de la posición física actual
        idFocoActual = ObtenerIdDesdePosicion(oVisor, oSelActual.CellAddress.Row, oSelActual.CellAddress.Column)
        
        ' REUTILIZACIÓN ESTRELLA N° 4: Inyectar la descripción legítima de ese ID
        EjecutarInyeccionDescripcion(oVisor, idFocoActual)
        
        ' Actualizar también la celda de depuración E1
        oVisor.getCellByPosition(4, 0).Value = idFocoActual
    End If
    ' -----------------------------------------------------------------------
    
    oDoc.calculateAll()
End Sub

REM --- MOTOR DE ORDENAMIENTO DE DATOS OPTIMIZADO ---
REM =========================================================================
REM PUNTO 5 AUDITORÍA: EL GUARDIÁN SUPREMO - AUDITA, CURA Y ORDENA (F2)
REM =========================================================================
Sub OrdenarDatosPersonalizado()
    Dim oDoc As Object, oDatos As Object, oCursor As Object
    Dim totalFilasDatos As Long, i As Long, j As Long
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' =========================================================================
    ' PASO 1: BLINDAJE Y CURACIÓN DE LA BASE DE DATOS
    ' Antes de mover una sola fila, reparamos cualquier omisión de IDs u Órdenes
    ' =========================================================================
    AsignarIdsFaltantesBD()
    AsignarOrdenesFaltantesBD()
    
    ' 2. Encontrar el límite real de filas en la hoja datos post-curación
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilasDatos = oCursor.RangeAddress.EndRow + 1
    
    ' Si sólo están las cabeceras o está vacío, salimos de forma segura
    If totalFilasDatos <= 2 Then Exit Sub
    
    ' 3. Cargar listas de referencia de prioridad desde las columnas H y J
    Dim listaActividades() As Variant, listaEstados() As Variant
    listaActividades = oDatos.getCellRangeByName("H2:H100").getDataArray()
    listaEstados = oDatos.getCellRangeByName("J2:J20").getDataArray()
    
    ' 4. Extraer toda la matriz de datos actual (Columnas A a F, desde fila 2)
    ' Expandido estrictamente hasta la columna F (índice 5) para arrastrar los IDs inmutables
    Dim oRangoCompleto As Object
    Dim matrizDatos() As Variant
    oRangoCompleto = oDatos.getCellRangeByPosition(0, 1, 5, totalFilasDatos - 1)
    matrizDatos = oRangoCompleto.getDataArray()
    
    Dim nPacientes As Long
    nPacientes = UBound(matrizDatos)
    
    Dim sActividad As String, sEstado As String
    Dim temporal As Variant
    Dim pAct1 As Long, pAct2 As Long
    Dim pEst1 As Long, pEst2 As Long
    Dim ord1 As Double, ord2 As Double
    
    ' 5. Algoritmo de ordenamiento por peso de aparición en memoria
    For i = 0 To nPacientes - 1
        For j = i + 1 To nPacientes
            ' --- TAREA 1 ---
            sActividad = matrizDatos(i)(0)
            sEstado = matrizDatos(i)(2)
            ord1 = CDbl(matrizDatos(i)(4))
            
            pAct1 = ObtenerIndiceLista(sActividad, listaActividades)
            pEst1 = ObtenerIndiceLista(sEstado, listaEstados)
            
            ' --- TAREA 2 ---
            sActividad = matrizDatos(j)(0)
            sEstado = matrizDatos(j)(2)
            ord2 = CDbl(matrizDatos(j)(4))
            
            pAct2 = ObtenerIndiceLista(sActividad, listaActividades)
            pEst2 = ObtenerIndiceLista(sEstado, listaEstados)
            
            ' Criterio de ordenación de tu filosofía de diseño:
            ' 1° Jerarquía de Actividad (Columna H)
            ' 2° Jerarquía de Estado (Columna J)
            ' 3° Posición interna / Interpolación (Columna E)
            If (pAct1 > pAct2) Or _
               (pAct1 = pAct2 And pEst1 > pEst2) Or _
               (pAct1 = pAct2 And pEst1 = pEst2 And ord1 > ord2) Then
                ' Switch completo de filas en memoria
                temporal = matrizDatos(i)
                matrizDatos(i) = matrizDatos(j)
                matrizDatos(j) = temporal
            End If
        Next j
    Next i
    
    ' =========================================================================
    ' UBICACIÓN CORREGIDA: ÍNDICE SECUENCIAL GLOBAL CON FILTRO ANTIFANTASMAS (F2 PRO)
    ' Normaliza la base de datos entera saltándose las filas residuales vacías
    ' =========================================================================
    For i = 0 To nPacientes
        ' Verificamos si la fila actual de la matriz realmente contiene una tarea válida
        If Trim(matrizDatos(i)(0)) <> "" And Trim(matrizDatos(i)(1)) <> "" Then
            ' Si es una tarea real, recibe su índice secuencial global corrido
            matrizDatos(i)(4) = CDbl(i + 1)
        Else
            ' SI LA FILA ESTÁ VACÍA (FANTASMA): Limpiamos cualquier residuo numérico en memoria
            matrizDatos(i)(4) = "" ' Borra el Orden (Col E)
            matrizDatos(i)(5) = "" ' Borra el ID (Col F)
        End If
    Next i
    ' =========================================================================
    
    ' 6. Volcar la matriz perfectamente ordenada e íntegra de vuelta a las celdas
    oRangoCompleto.setDataArray(matrizDatos)
    
    ' 7. Sincronización automática: Redibujar el lienzo Kanban del Visor
    DibujarTablero()
End Sub

REM =========================================================================
REM FUNCIÓN AUXILIAR: BUSCAR EL PESO/ÍNDICE DENTRO DE LAS LISTAS DE REFERENCIA
REM =========================================================================
Function ObtenerIndiceLista(sValor As String, ByRef lista() As Variant) As Long
    Dim i As Long
    sValor = Trim(LCase(sValor))
    
    For i = 0 To UBound(lista)
        If Trim(lista(i)(0)) <> "" Then
            If Trim(LCase(lista(i)(0))) = sValor Then
                ObtenerIndiceLista = i ' Devuelve la posición física (0, 1, 2...) como su nivel de prioridad
                Exit Function
            End If
        End If
    Next i
    
    ' Si por alguna razón la actividad/estado no está en la lista H o J, se le asigna la prioridad más baja
    ObtenerIndiceLista = 999 
End Function



REM =========================================================================
REM MOTOR UNIFICADO DE BORRADO EN CASCADA
REM =========================================================================
Sub EjecutarBorradoEnCascada(sItem As String, sTipoItem As String)
    Dim oDoc As Object, oDatos As Object, oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim contadorTareas As Long: contadorTareas = 0
    Dim respuesta As Integer
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' 1. Validación de seguridad para Estados Base
    If sTipoItem = "Estado" Then
        If sItem = "En curso" Or sItem = "Pendiente" Then
            respuesta = MsgBox("Aviso: Estás intentando eliminar un estado base del sistema ('" & sItem & "')." & Chr(13) & _
                               "Esto podría alterar la visualización por defecto de las columnas." & Chr(13) & _
                               "¿Deseas continuar de todos modos?", 4 + 32, "Validación de Estado Base")
            If respuesta <> 6 Then Exit Sub
        End If
    End If
    
    ' 2. Encontrar límites de la BD
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    ' 3. Contar tareas afectadas
    For i = 1 To totalFilas - 1
        Dim sValorComparar As String
        If sTipoItem = "Actividad" Then
            sValorComparar = oDatos.getCellByPosition(0, i).String ' Columna A
        Else
            sValorComparar = oDatos.getCellByPosition(2, i).String ' Columna C
        End If
        
        If sValorComparar = sItem Then
            contadorTareas = contadorTareas + 1
        End If
    Next i
    
    ' 4. Confirmación de seguridad crítica e irreversible
    respuesta = MsgBox("⚠️ ADVERTENCIA DE ELIMINACIÓN CRÍTICA ⚠️" & Chr(13) & Chr(13) & _
                       "Vas a eliminar " & sTipoItem & ": """ & sItem & """" & Chr(13) & _
                       "Se destruirán de forma PERMANENTE [" & contadorTareas & "] tareas asociadas." & Chr(13) & Chr(13) & _
                       "¿Confirmas la operación?", 4 + 16 + 256, "Confirmar Cascada de Datos")
                       
    If respuesta = 6 Then ' SÍ
        oDoc.addActionLock()
        
        ' A. Borrar tareas de abajo hacia arriba para mantener estables los índices de fila
        For i = totalFilas - 1 To 1 Step -1
            Dim sValorComparar2 As String
            If sTipoItem = "Actividad" Then
                sValorComparar2 = oDatos.getCellByPosition(0, i).String
            Else
                sValorComparar2 = oDatos.getCellByPosition(2, i).String
            End If
            
            If LCase(Trim(sValorComparar2)) = LCase(Trim(sItem)) Then
                oDatos.Rows.removeByIndex(i, 1)
            End If
        Next i
        
        ' B. Borrar el elemento del catálogo de configuración
        If sTipoItem = "Actividad" Then
            Dim mCatAct() As Variant
            mCatAct = oDatos.getCellRangeByName("H2:H100").getDataArray()
            For i = 0 To UBound(mCatAct)
                If LCase(Trim(mCatAct(i)(0))) = LCase(Trim(sItem)) Then
                    oDatos.getCellByPosition(7, i + 1).String = "" ' Columna H (índice 7)
                    Exit For
                End If
            Next i
        Else
            Dim mCatEst() As Variant
            mCatEst = oDatos.getCellRangeByName("J2:J20").getDataArray()
            For i = 0 To UBound(mCatEst)
                If LCase(Trim(mCatEst(i)(0))) = LCase(Trim(sItem)) Then
                    oDatos.getCellByPosition(9, i + 1).String = "" ' Columna J (índice 9)
                    Exit For
                End If
            Next i
        End If
        
        ' C. Reordenar y redibujar el tablero Kanban
        OrdenarDatosPersonalizado()
        
        oDoc.removeActionLock()
        oDoc.calculateAll()
        
        MsgBox sTipoItem & " y sus tareas han sido purgados de la base de datos.", 64, "Purga Completada"
    End If
End Sub






' --- 1. ACTIVAR EL ESCUCHADOR (Se debe ejecutar al abrir el archivo) ---
Sub IniciarListenerSeleccion()
    Dim oDoc As Object
    oDoc = ThisComponent
    oHojaEscuchada = oDoc.Sheets.getByName("visor")
    
    ' Crear el objeto escuchador de eventos de selección
    oSeleccionListener = CreateUnoListener("VisorSeleccion_", "com.sun.star.view.XSelectionChangeListener")
    oDoc.CurrentController.addSelectionChangeListener(oSeleccionListener)
End Sub

Sub QuitarListenerSeleccion()
    On Error Resume Next
    ThisComponent.CurrentController.removeSelectionChangeListener(oSeleccionListener)
End Sub


REM =========================================================================
REM COMPONENTE AISLADO: EXTRACTOR PURO DE ID DESCE COLUMNAS OCULTAS (E=4, F=5)
REM =========================================================================
Function ObtenerIdDesdePosicion(oVisor As Object, nFila As Long, nColumna As Long) As Long
    Dim idEncontrado As Long
    idEncontrado = 0
    
    ' El rango válido del Kanban es Filas de la 4 a la 300 (índices 3 a 299)
    If nFila >= 3 And nFila <= 299 Then
        If nColumna = 0 Then
            ' Columna A (0) -> Leer ID de la Columna E (4)
            idEncontrado = CLng(oVisor.getCellByPosition(4, nFila).Value)
        ElseIf nColumna = 1 Then
            ' Columna B (1) -> Leer ID de la Columna F (5)
            idEncontrado = CLng(oVisor.getCellByPosition(5, nFila).Value)
        End If
    End If
    
    ObtenerIdDesdePosicion = idEncontrado
End Function

REM =========================================================================
REM LISTENER ULTRA-SIMPLIFICADO PARA PRUEBAS DE ENLACE DE ID
REM =========================================================================
Sub VisorSeleccion_selectionChanged(oEvent)
    Dim oDoc As Object, oVisor As Object, oSel As Object
    Dim nFila As Long, nColumna As Long
    Dim sHojaActiva As String
    Dim idDetectado As Long
    
    On Error Resume Next
    
    
    
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    
    ' 1. ESCAPE: Validar que sea una sola celda
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    
    ' 2. ESCAPE: Operar estrictamente en la hoja "visor"
    If oSel.Spreadsheet.Name <> "visor" Then Exit Sub
    
    nFila = oSel.CellAddress.Row
    nColumna = oSel.CellAddress.Column
    oVisor = oDoc.Sheets.getByName("visor")
    
    ' 3. OBTENER EL ID USANDO NUESTRO COMPONENTE SEPARADO
    idDetectado = ObtenerIdDesdePosicion(oVisor, nFila, nColumna)
    
    ' 4. DEPOSITAR EL ID EN E1 PARA COMPROBACIÓN VISUAL
    ' Celda E1 en base cero es Columna 4, Fila 0
    oVisor.getCellByPosition(4, 0).Value = idDetectado
    
    ' Lanzar Proceso 1: Gestión de Descripciones
    EjecutarInyeccionDescripcion(oVisor, idDetectado)
    
    ' Refrescar pantalla de forma sutil
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM EVENTO DE HOJA: CONTENIDO MODIFICADO EN LA HOJA VISOR (REFRESCO KANBAN)
REM =========================================================================
Sub Visor_onContentChanged(oEvent As Object)
    Dim nRow As Long, nCol As Long
    
    On Error Resume Next
    
    ' 1. ESCAPE: Validar que el cambio provenga de una sola celda física
    If Not oEvent.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    
    nRow = oEvent.CellAddress.Row
    nCol = oEvent.CellAddress.Column
    
    ' 2. EVALUAR COORDENADAS CRÍTICAS DE FILTRADO:
    ' - A2 (Fila 1, Col 0) o C2 (Fila 1, Col 2) -> Filtro de Actividad activa
    ' - A3 (Fila 2, Col 0) -> Filtro de Estado Izquierdo (Columna A)
    ' - B3 (Fila 2, Col 1) -> Filtro de Estado Derecho (Columna B)
    If (nRow = 1 And (nCol = 0 Or nCol = 2)) Or _
       (nRow = 2 And nCol = 0) Or _
       (nRow = 2 And nCol = 1) Then
       
        ' Congelar temporalmente el Listener de selección para evitar bucles visuales 
        ' mientras DibujarTablero reescribe las celdas del visor
        binFocoCongelado = True
        
        ' Redibujar síncronamente el tablero para reflejar el nuevo filtro
        'Call DibujarTablero()
        Call OrdenarDatosPersonalizado()
                
        ' Reactivar el Listener de selección una vez completado el redibujado
        binFocoCongelado = False
    End If
End Sub

REM =========================================================================
REM FASE 2 - COMPONENTE 1: BUSCADOR DE HUECOS Y GENERADOR DE ID DISPONIBLE
REM =========================================================================
Function ObtenerSiguienteIdDisponible() As Long
    Dim oDoc As Object, oDatos As Object
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim idCandidato As Long
    Dim idExiste As Boolean
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' Encontrar el límite real de la hoja datos
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    ' Si no hay filas de datos (solo cabecera), el primer ID es 1
    If totalFilas <= 2 Then
        ObtenerSiguienteIdDisponible = 1
        Exit Function
    End If
    
    ' LÓGICA DE REUTILIZACIÓN DE HUECOS:
    ' Empezamos probando desde el ID = 1 hacia arriba.
    ' El primer número que NO encontremos en la columna F, ese será nuestro ID elegido.
    idCandidato = 1
    Do
        idExiste = False
        ' Recorrer la columna F (índice 5) buscando si el idCandidato ya está ocupado
        For i = 1 To totalFilas - 1
            If CLng(oDatos.getCellByPosition(5, i).Value) = idCandidato Then
                idExiste = True
                Exit For ' Ya está ocupado, no hace falta seguir buscando en esta vuelta
            End If
        Next i
        
        ' Si no existía, ¡hemos encontrado un hueco o el final de la secuencia!
        If Not idExiste Then
            Exit Do
        End If
        
        ' Si existía, probamos con el siguiente número
        idCandidato = idCandidato + 1
    Loop
    
    ObtenerSiguienteIdDisponible = idCandidato
End Function


REM =========================================================================
REM FASE 2 - COMPONENTE 2: BARREDORA COMPLETA DE REGISTROS SIN ID
REM =========================================================================
Sub AsignarIdsFaltantesBD()
    Dim oDoc As Object, oDatos As Object
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim sActividad As String, sTarea As String, sEstado As String
    Dim idActual As Long, nuevoId As Long
    Dim conteoAsignados As Long
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    conteoAsignados = 0
    
    ' Desactivar actualización de pantalla para procesar en milisegundos
    oDoc.addActionLock()
    
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    ' Recorrer desde la fila 2 (índice 1) en adelante
    For i = 1 To totalFilas - 1
        ' Leer campos obligatorios de validación
        sActividad = Trim(oDatos.getCellByPosition(0, i).String) ' Col A
        sTarea = Trim(oDatos.getCellByPosition(1, i).String)     ' Col B
        sEstado = Trim(oDatos.getCellByPosition(2, i).String)    ' Col C
        idActual = oDatos.getCellByPosition(5, i).Value          ' Col F
        
        ' REGLA DE VALIDACIÓN:
        ' Si tiene Actividad, Tarea y Estado (Campos llenos) Y el ID es 0 o está vacío
        If sActividad <> "" And sTarea <> "" And sEstado <> "" And idActual = 0 Then
            
            ' Consultar el siguiente ID disponible (reutilizando huecos si existen)
            nuevoId = ObtenerSiguienteIdDisponible()
            
            ' Inyectar el ID en la Columna F
            oDatos.getCellByPosition(5, i).Value = nuevoId
            conteoAsignados = conteoAsignados + 1
        End If
    Next i
    
    oDoc.removeActionLock()
    
    ' Mensaje sutil de confirmación si se ejecutó manualmente y asignó algo
    If conteoAsignados > 0 Then
        MsgBox "Se han auditado los registros. IDs asignados con éxito: " & conteoAsignados, 64, "Mantenimiento de IDs"
    End If
End Sub

REM =========================================================================
REM PROCESO AISLADO: INYECTAR TÍTULO (D3) Y DESCRIPCIÓN (D4) EN EL VISOR
REM =========================================================================
Sub EjecutarInyeccionDescripcion(oVisor As Object, idBusqueda As Long)
    Dim oDoc As Object, oDatos As Object, oCursor As Object
    Dim totalFilasDatos As Long, i As Long
    Dim sDescripcion As String
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' 1. Limpiar la columna D por completo antes de pintar (Solo la celda D4, al ser un rango combinado)
    ' Esto garantiza que no queden "fantasmas" visuales de celdas anteriores
    oVisor.getCellByPosition(3, 3).String = ""
    oVisor.getCellByPosition(3, 2).String = "" ' Celda D3 (índice Columna 3, Fila 2)
    
    ' 2. Si el ID es 0 (celda vacía), dejamos limpio y salimos
    If idBusqueda = 0 Then Exit Sub
    
    ' 3. Buscar la descripción en la base de datos por ID
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilasDatos = oCursor.RangeAddress.EndRow + 1
    
    sTitulo = ""
    sDescripcion = ""
    For i = 1 To totalFilasDatos - 1
        If CLng(oDatos.getCellByPosition(5, i).Value) = idBusqueda Then
            sDescripcion = oDatos.getCellByPosition(3, i).String ' Columna D de Datos
            sTitulo = oDatos.getCellByPosition(1, i).String      ' Columna B (índice 1 - Título/Tarea)
            Exit For
        End If
    Next i
    
    ' 4. Escribir el resultado en el Visor
    oVisor.getCellByPosition(3, 2).String = sTitulo       ' Escribe en D3 (Título)
    oVisor.getCellByPosition(3, 3).String = sDescripcion
End Sub

REM =========================================================================
REM MOTOR ÚNICO DE PRIORIDADES: VERSIÓN ÍNDICE SECUENCIAL GLOBAL (F2 PRO)
REM =========================================================================
Function ObtenerSiguienteOrden(sActividad As String, sEstado As String, Optional nPosDeseada As Long) As Double
    Dim oDoc As Object, oDatos As Object, oCursor As Object
    Dim totalFilasActuales As Long, i As Long
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilasActuales = oCursor.RangeAddress.EndRow + 1
    
    If IsMissing(nPosDeseada) Then nPosDeseada = 0
    
    ' =========================================================================
    ' ESCENARIO A: INTERPOLACIÓN GLOBAL DENTRO DEL KANBAN
    ' =========================================================================
    If nPosDeseada > 0 Then
        Dim mHistorico() As Variant, oRangoTareas As Object
        oRangoTareas = oDatos.getCellRangeByPosition(0, 1, 4, totalFilasActuales - 1)
        mHistorico = oRangoTareas.getDataArray()
        
        ' Filtramos para ubicar los índices globales de las tareas de este grupo específico
        Dim mIndicesGlobales() As Long, cFiltrados As Long : cFiltrados = 0
        For i = 0 To UBound(mHistorico)
            If Trim(mHistorico(i)(0)) = sActividad And Trim(mHistorico(i)(2)) = sEstado Then
                ReDim Preserve mIndicesGlobales(cFiltrados)
                mIndicesGlobales(cFiltrados) = i ' Guardamos la posición física de la fila
                cFiltrados = cFiltrados + 1
            End If
        Next i
        
        If cFiltrados > 0 Then
            ' Determinar los pesos globales en la columna E basados en la posición deseada
            If nPosDeseada <= 1 Then
                ' Al principio del grupo: Peso de la primera tarea del grupo menos 0.1
                ObtenerSiguienteOrden = CDbl(mHistorico(mIndicesGlobales(0))(4)) - 0.1
            ElseIf nPosDeseada > cFiltrados Then
                ' Al final del grupo: Peso de la última tarea del grupo más 0.1
                ObtenerSiguienteOrden = CDbl(mHistorico(mIndicesGlobales(cFiltrados - 1))(4)) + 0.1
            Else
                ' En medio: Punto medio decimal entre la tarea elegida y la anterior
                Dim ordenAnterior As Double, ordenSiguiente As Double
                ordenAnterior = CDbl(mHistorico(mIndicesGlobales(nPosDeseada - 2))(4))
                ordenSiguiente = CDbl(mHistorico(mIndicesGlobales(nPosDeseada - 1))(4))
                ObtenerSiguienteOrden = ordenAnterior + ((ordenSiguiente - ordenAnterior) / 2)
            End If
            Exit Function
        End If
    End If
    
    ' =========================================================================
    ' ESCENARIO B: CONSECUTIVO AL FINAL ABSOLUTO DE LA TABLA
    ' =========================================================================
    ' Si la tabla está vacía (solo cabecera), inicia en 1, sino va al final de la lista corrida
    If totalFilasActuales <= 2 Then
        ObtenerSiguienteOrden = 1.0
    Else
        ObtenerSiguienteOrden = CDbl(totalFilasActuales)
    End If
End Function

REM =========================================================================
REM COMPONENTE DEFENSIVO: REPARA Y ASIGNA NÚMEROS DE ORDEN FALTANTES (F2)
REM =========================================================================
Sub AsignarOrdenesFaltantesBD()
    Dim oDoc As Object, oDatos As Object, oCursor As Object
    Dim totalFilasDatos As Long, i As Long
    Dim sActividad As String, sEstado As String
    Dim oCeldaOrden As Object
    Dim nOrdenAsignado As Double
    
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' 1. Encontrar el límite real de filas en la hoja datos
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilasDatos = oCursor.RangeAddress.EndRow + 1
    
    ' Si sólo están las cabeceras, no hay nada que auditar
    If totalFilasDatos <= 2 Then Exit Sub
    
    ' 2. Barrido fila por fila analizando la columna E (índice 4)
    For i = 1 To totalFilasDatos - 1
        sActividad = Trim(oDatos.getCellByPosition(0, i).String) ' Col A
        sEstado = Trim(oDatos.getCellByPosition(2, i).String)    ' Col C
        oCeldaOrden = oDatos.getCellByPosition(4, i)             ' Col E
        
        ' Una fila es válida si al menos tiene Actividad y Nombre de Tarea
        If sActividad <> "" And Trim(oDatos.getCellByPosition(1, i).String) <> "" Then
            
            ' Detectar si el orden es nulo, vacío o cero
            If oCeldaOrden.Type = com.sun.star.table.CellContentType.EMPTY Or oCeldaOrden.Value = 0 Then
                
                ' Solicitamos al componente experto el siguiente número entero para este grupo
                ' Nota: No pasamos el tercer parámetro porque queremos que vaya al final de su lista
                nOrdenAsignado = ObtenerSiguienteOrden(sActividad, sEstado)
                
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
    Dim oDoc As Object, oVisor As Object, oDatos As Object
    Dim oModel As Object, oSel As Object
    
    ' 1. CARGA DE ENTORNO GLOBAL (Siempre se usará)
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("visor")
    oDatos = oDoc.Sheets.getByName("datos")
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
        Call ServPoblarCmbActividad(oModel.getByName("CmbActividad"), oDialogo.getControl("CmbActividad"), oDatos, idActivo)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 2: POBLAR COMBO DE ESTADOS
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbEstado") Then
        ' Pasamos el modelo y la vista por separado a la macro externa
        Call ServPoblarCmbEstado(oModel.getByName("CmbEstado"), oDialogo.getControl("CmbEstado"), oDatos, idActivo)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 3: INYECTAR TÍTULO (Contextual al ID Activo o Selección)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitulo") Then
        Dim oCtrlTitulo As Object
        oCtrlTitulo = oModel.getByName("TxtTitulo")
        ' Limpio y atómico: Solo se encarga del texto del título
        Call ServGestionarTxtTitulo(oCtrlTitulo, idActivo, oDatos)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 4: INYECTAR DESCRIPCIÓN
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescripcion") Then
        Dim oCtrlDesc As Object
        oCtrlDesc = oModel.getByName("TxtDescripcion")
        Call ServGestionarTxtDescripcion(oCtrlDesc, idActivo, oDatos)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 5: CALCULAR Y ASIGNAR POSICIÓN RELATIVA
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosicion") Then
        ' Enviamos modelo y vista por separado
        Call ServCalcularNumPosicion(oModel.getByName("NumPosicion"), oDialogo.getControl("NumPosicion"), idActivo, oDatos, oModel)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 6 (UNIFICADO): POBLAR COMBO DE BORRADO DE ITEMS (CmbItem)
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbItem") Then
        Dim sTipoItem As String : sTipoItem = ""
        ' Detectar dinámicamente el contexto por el título de la ventana
        If InStr(LCase(oDialogo.Title), "actividad") > 0 Then
            sTipoItem = "Actividad"
        ElseIf InStr(LCase(oDialogo.Title), "estado") > 0 Then
            sTipoItem = "Estado"
        End If
        
        Call ServPoblarCmbItem(oModel.getByName("CmbItem"), oDialogo.getControl("CmbItem"), oDatos, sTipoItem)
    End If
    
    ' Reactivar actualización de pantalla
    oDoc.removeActionLock()
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 1: POBLAR Y SELECCIONAR ACTIVIDAD (VISTA)
REM =========================================================================
Sub ServPoblarCmbActividad(oModelCombo As Object, oControlView As Object, oDatos As Object, idActivo As Long)
    Dim  oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim sValor As String, sActividadBuscar As String
    Dim vLista() As String, nContador As Long : nContador = 0
    Dim idxActividad As Long : idxActividad = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Cargar datos en el Array
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    ReDim vLista(totalFilas)
    
    For i = 1 To totalFilas - 1
        sValor = Trim(oDatos.getCellByPosition(7, i).String)
        If sValor <> "" Then
            vLista(nContador) = sValor
            nContador = nContador + 1
        End If
    Next i
    
    ' 2. Inyectar la lista al modelo
    If nContador > 0 Then
        ReDim Preserve vLista(nContador - 1)
        oModelCombo.StringItemList = vLista()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
        Exit Sub
    End If
    
    ' 3. PRE-SELECCIÓN MEDIANTE TU PROPUESTA DE VISTA (.selectItemPos)
    If idActivo > 0 Then
        For i = 1 To totalFilas - 1
            If oDatos.getCellByPosition(5, i).Value = idActivo Then
                sActividadBuscar = oDatos.getCellByPosition(0, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sActividadBuscar Then
                idxActividad = i
                Exit For
            End If
        Next i
        
        If idxActividad <> -1 Then
            ' Método nativo e infalible para forzar el texto visual en la Vista de un ComboBox
            oControlView.setText(sActividadBuscar)
        End If
    End If
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 2: POBLAR Y SELECCIONAR ESTADO (VISTA)
REM =========================================================================
Sub ServPoblarCmbEstado(oModelCombo As Object, oControlView As Object, oDatos As Object, idActivo As Long)
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim sValor As String, sEstadoBuscar As String
    Dim vLista() As String, nContador As Long : nContador = 0
    Dim idxEstado As Long : idxEstado = -1
    
    oModelCombo = oControlView.Model
    
    ' 1. Cargar datos en el Array
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    ReDim vLista(totalFilas)
    
    For i = 1 To totalFilas - 1
        sValor = Trim(oDatos.getCellByPosition(9, i).String)
        If sValor <> "" And LCase(sValor) <> "Papelera" Then
            vLista(nContador) = sValor
            nContador = nContador + 1
        End If
    Next i
    
    ' 2. Inyectar la lista al modelo
    If nContador > 0 Then
        ReDim Preserve vLista(nContador - 1)
        oModelCombo.StringItemList = vLista()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
        Exit Sub
    End If
    
    ' 3. PRE-SELECCIÓN MEDIANTE TU PROPUESTA DE VISTA (.selectItemPos)
    If idActivo > 0 Then
        For i = 1 To totalFilas - 1
            If oDatos.getCellByPosition(5, i).Value = idActivo Then
                sEstadoBuscar = oDatos.getCellByPosition(2, i).String
                Exit For
            End If
        Next i
        
        For i = 0 To UBound(oModelCombo.StringItemList)
            If oModelCombo.StringItemList(i) = sEstadoBuscar Then
                idxEstado = i
                Exit For
            End If
        Next i
        
        If idxEstado <> -1 Then
            ' Método nativo e infalible para forzar el texto visual en la Vista de un ComboBox
            oControlView.setText(sEstadoBuscar)
        End If
    End If
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 3: GESTIONAR TÍTULO DE LA TAREA (ATÓMICO)
REM =========================================================================
Sub ServGestionarTxtTitulo(oModelTitulo As Object, idActivo As Long, oDatos As Object)
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    
    ' Si el ID es 0, es un alta nueva; limpiamos el título y salimos
    If idActivo = 0 Then
        oModelTitulo.Text = ""
        Exit Sub
    End If
    
    ' Buscamos la fila de la tarea en la hoja datos
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To totalFilas - 1
        ' Columna F (Índice 5) contiene el ID de la tarea
        If oDatos.getCellByPosition(5, i).Value = idActivo Then
            ' Inyectar únicamente el título actual (Columna B, Índice 1)
            oModelTitulo.Text = oDatos.getCellByPosition(1, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 4: GESTIONAR DESCRIPCIÓN DE LA TAREA
REM =========================================================================
Sub ServGestionarTxtDescripcion(oModelDesc As Object, idActivo As Long, oDatos As Object)
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    
    ' Si el ID es 0, es una tarea nueva; dejamos la descripción limpia
    If idActivo = 0 Then
        oModelDesc.Text = ""
        Exit Sub
    End If
    
    ' Buscamos la fila de la tarea para extraer la descripción
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To totalFilas - 1
        ' Columna F (Índice 5) contiene el ID de la tarea
        If oDatos.getCellByPosition(5, i).Value = idActivo Then
            ' Columna D (Índice 3) contiene la Descripción
            oModelDesc.Text = oDatos.getCellByPosition(3, i).String
            Exit Sub
        End If
    Next i
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 5: GENERAR MAPA DE POSICIONES VISUALES (CORREGIDO)
REM =========================================================================
Sub ServCalcularNumPosicion(oModelCombo As Object, oControlView As Object, idActivo As Long, oDatos As Object, oModelDialogo As Object)
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long
    Dim sEstadoTarea As String, sActividadTarea As String
    Dim nTotalEnEstado As Long : nTotalEnEstado = 0
    Dim nPosicionActual As Long : nPosicionActual = 0
    Dim vLista() As String
    
    ' 1. Determinar el contexto de ACTIVIDAD y ESTADO en el que opera la tarea
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    If idActivo = 0 Then
        ' Si es nueva tarea, leemos los combos del diálogo
        If oModelDialogo.hasByName("CmbActividad") Then sActividadTarea = Trim(oModelDialogo.getByName("CmbActividad").Text)
        If oModelDialogo.hasByName("CmbEstado") Then sEstadoTarea = Trim(oModelDialogo.getByName("CmbEstado").Text)
        
        ' Valores por defecto si están vacíos
        If sActividadTarea = "" Then sActividadTarea = "General"
        If sEstadoTarea = "" Then sEstadoTarea = "Pendiente"
    Else
        ' Si la tarea existe, extraemos su Actividad (Col A) y Estado (Col C) reales de la fila
        For i = 1 To totalFilas - 1
            If oDatos.getCellByPosition(5, i).Value = idActivo Then
                sActividadTarea = oDatos.getCellByPosition(0, i).String ' Col A
                sEstadoTarea    = oDatos.getCellByPosition(2, i).String ' Col C
                Exit For
            End If
        Next i
    End If
    
    ' 2. Contar de forma precisa (Filtrando por Actividad + Estado y omitiendo filas basura)
    Dim sEstFila As String, sActFila As String
    
    For i = 1 To totalFilas - 1
        sActFila = Trim(oDatos.getCellByPosition(0, i).String)
        sEstFila = Trim(oDatos.getCellByPosition(2, i).String)
        
        ' REGLA DE ORO: Deben coincidir Actividad Y Estado, y NO ser papelera ni estar vacíos
        If sActFila = sActividadTarea And sEstFila = sEstadoTarea And LCase(sEstFila) <> "Papelera" And sActFila <> "" Then
            nTotalEnEstado = nTotalEnEstado + 1
            
            ' Si encontramos el ID que estamos buscando, guardamos su posición exacta en este subconjunto
            If idActivo > 0 Then
                If oDatos.getCellByPosition(5, i).Value = idActivo Then
                    nPosicionActual = nTotalEnEstado
                End If
            End If
        End If
    Next i
    
    ' 3. Construir la lista de opciones del Combo basadas en el conteo real
    Dim nLimiteOpciones As Long
    nLimiteOpciones = nTotalEnEstado
    
    ' Creamos el array para los enteros (1, 2, 3...) más la opción "F"
    ReDim vLista(nLimiteOpciones)
    
    For i = 1 To nLimiteOpciones
        vLista(i - 1) = CStr(i)
    Next i
    vLista(nLimiteOpciones) = "F"
    
    ' Inyectamos la lista limpia de posiciones
    oModelCombo.StringItemList = vLista()
    
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
Sub ServPoblarCmbItem(oModelCombo As Object, oControlView As Object, oDatos As Object, sTipoItem As String)
    Dim oCursor As Object
    Dim totalFilas As Long, i As Long, nContador As Long
    Dim sValor As String
    Dim vLista() As String
    
    oModelCombo = oControlView.Model
    
    ' Obtener límites en la hoja datos
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    ReDim vLista(totalFilas)
    nContador = 0
    
    ' Extraer textos del catálogo correspondiente
    For i = 1 To totalFilas - 1
        If sTipoItem = "Actividad" Then
            sValor = Trim(oDatos.getCellByPosition(7, i).String) ' Columna H (índice 7)
        Else
            sValor = Trim(oDatos.getCellByPosition(9, i).String) ' Columna J (índice 9)
        End If
        
        If sValor <> "" Then
            vLista(nContador) = sValor
            nContador = nContador + 1
        End If
    Next i
    
    ' Inyectar el catálogo en el ComboBox
    If nContador > 0 Then
        ReDim Preserve vLista(nContador - 1)
        oModelCombo.StringItemList = vLista()
    Else
        Dim vVacio(0) As String
        oModelCombo.StringItemList = vVacio()
    End If
End Sub





'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM HUB CENTRAL DE ESCRITURA AUTÓNOMO (ESTRUCTURA ATÓMICA)
REM =========================================================================
Sub GuardarDatosDesdeDialogo(oDialogo As Object, idTarea As Long)
    Dim oDoc As Object, oDatos As Object
    Dim oModel As Object, oCursor As Object
    Dim nFilaDestino As Long
    Dim totalFilas As Long, i As Long
    Dim nuevoId As Long
    Dim binEsNuevo As Boolean
    
    ' 1. Carga de entorno
    oDoc = ThisComponent
    oDatos = oDoc.Sheets.getByName("datos")
    oModel = oDialogo.Model
    
    ' Desactivar actualización de pantalla para máxima velocidad y evitar parpadeos
    oDoc.addActionLock()
    
    ' =========================================================================
    ' BLOQUE 1: DETERMINACIÓN DE PARÁMETROS EN MEMORIA (SIN MODIFICAR LA BD)
    ' =========================================================================
    If idTarea = 0 Then
        ' --- ESCENARIO A: REGISTRO DE NUEVA TAREA ---
        binEsNuevo = True
        
        oCursor = oDatos.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        nFilaDestino = oCursor.RangeAddress.EndRow + 1
        
        ' Calcular el nuevo ID disponible únicamente en memoria
        nuevoId = ObtenerSiguienteIdDisponible()
        idTarea = nuevoId
    Else
        ' --- ESCENARIO B: EDICIÓN DE TAREA EXISTENTE ---
        binEsNuevo = False
        
        oCursor = oDatos.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        totalFilas = oCursor.RangeAddress.EndRow + 1
        
        nFilaDestino = -1
        ' Escaneo de base de datos para localizar la fila por ID (Columna F / índice 5)
        For i = 1 To totalFilas - 1
            If CLng(oDatos.getCellByPosition(5, i).Value) = idTarea Then
                nFilaDestino = i
                Exit For
            End If
        Next i
        
        ' Control de seguridad si el ID no se localiza en la BD
        If nFilaDestino = -1 Then
            oDoc.removeActionLock()
            MsgBox "Error: No se pudo localizar la tarea con ID " & idTarea & " en la base de datos.", 48, "Error de Sincronización"
            Exit Sub
        End If
    End If
    
    ' =========================================================================
    ' BLOQUE 2: PROCESAMIENTO Y ESCRITURA DE DATOS EN LA BD
    ' =========================================================================
    
    ' Inyección del ID Único (Columna F / índice 5)
    ' Se centraliza en este bloque junto con el resto de escrituras físicas
    oDatos.getCellByPosition(5, nFilaDestino).Value = idTarea
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 1: GUARDAR ACTIVIDAD
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbActividad") Then
        ' Extrae el texto del combo e inyecta en Col A (0)
        Call ServGuardarCmbActividad(oDialogo.getControl("CmbActividad"), oDatos, nFilaDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 2: GUARDAR ESTADO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("CmbEstado") Then
        ' [Pendiente de implementar]: Extrae el texto del combo e inyecta en Col C (2)
        Call ServGuardarCmbEstado(oDialogo.getControl("CmbEstado"), oDatos, nFilaDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 3: GUARDAR TÍTULO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtTitulo") Then
        ' [Pendiente de implementar]: Extrae el texto del título e inyecta en Col B (1)
        Call ServGuardarTxtTitulo(oDialogo.getControl("TxtTitulo"), oDatos, nFilaDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 4: GUARDAR DESCRIPCIÓN
    REM ---------------------------------------------------------------------
    If oModel.hasByName("TxtDescripcion") Then
        ' [Pendiente de implementar]: Extrae la descripción e inyecta en Col D (3)
        Call ServGuardarTxtDescripcion(oDialogo.getControl("TxtDescripcion"), oDatos, nFilaDestino)
    End If
    
    REM ---------------------------------------------------------------------
    REM SERVICIO 5: GUARDAR POSICIÓN / ORDEN RELATIVO
    REM ---------------------------------------------------------------------
    If oModel.hasByName("NumPosicion") Then
        ' [Pendiente de implementar]: Procesa la posición ("F" o valor numérico) y calcula el peso en Col E (4)
        Call ServGuardarNumPosicion(oDialogo.getControl("NumPosicion"), oDatos, nFilaDestino, oModel, binEsNuevo)
    ElseIf binEsNuevo Then
        ' Si es una nueva tarea y el diálogo no expone control de posición, se asigna al final por defecto
        Call ServGuardarNumPosicionPorDefecto(oDatos, nFilaDestino, oModel)
    End If
    
    ' 3. Liberar bloqueo de pantalla
    oDoc.removeActionLock()
    
    ' 4. Mantenimiento y actualización automática
    ' Ordena la base de datos físicamente y redibuja la pantalla
    OrdenarDatosPersonalizado()
    oDoc.calculateAll()
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 1: GUARDAR ACTIVIDAD EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarCmbActividad(oControlView As Object, oDatos As Object, nFilaDestino As Long)
    Dim sActividad As String
    
    ' Extraer el texto limpio directamente de la vista del ComboBox
    sActividad = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna A (índice 0) de la fila correspondiente
    oDatos.getCellByPosition(0, nFilaDestino).String = sActividad
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 2: GUARDAR ESTADO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarCmbEstado(oControlView As Object, oDatos As Object, nFilaDestino As Long)
    Dim sEstado As String
    
    ' Extraer el texto limpio directamente de la vista del ComboBox de estados
    sEstado = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna C (índice 2) de la fila correspondiente
    oDatos.getCellByPosition(2, nFilaDestino).String = sEstado
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 3: GUARDAR TÍTULO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarTxtTitulo(oControlView As Object, oDatos As Object, nFilaDestino As Long)
    Dim sTitulo As String
    
    ' Extraer el texto limpio del cuadro de texto (TextField)
    sTitulo = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna B (índice 1) de la fila correspondiente
    oDatos.getCellByPosition(1, nFilaDestino).String = sTitulo
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 4: GUARDAR DESCRIPCIÓN EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarTxtDescripcion(oControlView As Object, oDatos As Object, nFilaDestino As Long)
    Dim sDescripcion As String
    
    ' Extraer el texto del cuadro de descripción (admite formato multilínea)
    sDescripcion = Trim(oControlView.Text)
    
    ' Escribir el valor en la Columna D (índice 3) de la fila correspondiente
    oDatos.getCellByPosition(3, nFilaDestino).String = sDescripcion
End Sub

REM =========================================================================
REM SERVICIO EXTERNO 5: GUARDAR POSICIÓN / ORDEN RELATIVO EN LA BASE DE DATOS
REM =========================================================================
Sub ServGuardarNumPosicion(oControlView As Object, oDatos As Object, nFilaDestino As Long, oModelDialogo As Object, binEsNuevo As Boolean)
    Dim sActividad As String, sEstado As String
    Dim sPosText As String
    Dim nOrdenCalculado As Double
    
    ' 1. Leer la actividad y el estado reales ya escritos físicamente en la fila destino
    sActividad = oDatos.getCellByPosition(0, nFilaDestino).String
    sEstado = oDatos.getCellByPosition(2, nFilaDestino).String
    
    ' 2. Extraer el texto de la vista del ComboBox NumPosicion
    sPosText = Trim(oControlView.Text)
    
    ' 3. Evaluar la entrada y solicitar el peso al motor de prioridades
    If LCase(sPosText) = "f" Or sPosText = "" Or Not IsNumeric(sPosText) Then
        ' Caso "F", vacío o entrada inválida: Colocar al final absoluto del grupo
        nOrdenCalculado = ObtenerSiguienteOrden(sActividad, sEstado, 0)
    Else
        ' Caso numérico: Interpolar la posición en medio de la lista de tareas
        nOrdenCalculado = ObtenerSiguienteOrden(sActividad, sEstado, CLng(sPosText))
    End If
    
    ' 4. Guardar el peso decimal en la Columna E (índice 4)
    oDatos.getCellByPosition(4, nFilaDestino).Value = nOrdenCalculado
End Sub

REM =========================================================================
REM SERVICIO AUXILIAR 5: ASIGNAR ORDEN RELATIVO POR DEFECTO (AL FINAL)
REM =========================================================================
Sub ServGuardarNumPosicionPorDefecto(oDatos As Object, nFilaDestino As Long, oModelDialogo As Object)
    Dim sActividad As String, sEstado As String
    Dim nOrdenCalculado As Double
    
    ' 1. Leer el contexto de grupo actual de la fila
    sActividad = oDatos.getCellByPosition(0, nFilaDestino).String
    sEstado = oDatos.getCellByPosition(2, nFilaDestino).String
    
    ' 2. Solicitar un orden al final de su respectiva lista
    nOrdenCalculado = ObtenerSiguienteOrden(sActividad, sEstado, 0)
    
    ' 3. Escribir en la Columna E (índice 4)
    oDatos.getCellByPosition(4, nFilaDestino).Value = nOrdenCalculado
End Sub




'------------------------------------------------------------------------------------------------------

REM =========================================================================
REM DIALOGOS Y BOTONES
REM =========================================================================
REM =========================================================================
REM MACRO UNIFICADA: ENVIAR A PAPELERA (SOPORTA DIÁLOGO Y COMANDO)
REM =========================================================================
Sub MoverTareaAPapelera(Optional oEvent As Object)
    Dim oDoc As Object, oDatos As Object, oVisor As Object, oSel As Object
    Dim oDialogo As Object
    Dim idActivo As Long, totalFilas As Long, i As Long
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
    oVisor = oDoc.Sheets.getByName("visor")
    oDatos = oDoc.Sheets.getByName("datos")
    oSel = oDoc.CurrentSelection
    
    ' Salida de seguridad si no hay una sola celda seleccionada
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    idActivo = ObtenerIdDesdePosicion(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If idActivo <= 0 Then
        MsgBox "No se detectó ninguna tarea activa para mover a la papelera.", 48, "Atención"
        Exit Sub
    End If
    
    ' 3. Localizar la fila física en la base de datos (Columna F / índice 5)
    Dim oCursor As Object
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To totalFilas - 1
        If CLng(oDatos.getCellByPosition(5, i).Value) = idActivo Then
            filaDatosIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Ejecución del cambio de estado lógico
    If filaDatosIdx <> -1 Then
        ' Bloquear actualización de pantalla
        oDoc.addActionLock()
        
        ' Cambiar el Estado a "Papelera" en la Columna C (índice 2)
        oDatos.getCellByPosition(2, filaDatosIdx).String = "Papelera"
        
        ' Asignar un orden muy alto para sacarla de las prioridades activas
        oDatos.getCellByPosition(4, filaDatosIdx).Value = 99999
        
        ' Limpiar descripciones en el Visor
        EjecutarInyeccionDescripcion(oVisor, 0)
        
        ' Reordenar la base de datos y redibujar el tablero síncronamente
        OrdenarDatosPersonalizado()
        
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
    Dim oDoc As Object, oDatos As Object, oVisor As Object, oSel As Object
    Dim oDialogo As Object
    Dim idActivo As Long, totalFilas As Long, i As Long
    Dim filaDatosIdx As Long : filaDatosIdx = -1
    Dim respuesta As Integer
    Dim sNombreTareaBorrar As String, sActividadBorrar As String
    
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
    oVisor = oDoc.Sheets.getByName("visor")
    oDatos = oDoc.Sheets.getByName("datos")
    oSel = oDoc.CurrentSelection
    
    ' Salida de seguridad
    If Not oSel.supportsService("com.sun.star.sheet.SheetCell") Then Exit Sub
    idActivo = ObtenerIdDesdePosicion(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
    
    If idActivo <= 0 Then
        MsgBox "No se detectó ninguna tarea activa para eliminar.", 48, "Atención"
        Exit Sub
    End If
    
    ' 3. Localizar la fila física en la base de datos (Columna F / índice 5)
    Dim oCursor As Object
    oCursor = oDatos.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    totalFilas = oCursor.RangeAddress.EndRow + 1
    
    For i = 1 To totalFilas - 1
        If CLng(oDatos.getCellByPosition(5, i).Value) = idActivo Then
            filaDatosIdx = i
            Exit For
        End If
    Next i
    
    ' 4. Confirmación crítica y eliminación física
    If filaDatosIdx <> -1 Then
        sActividadBorrar = oDatos.getCellByPosition(0, filaDatosIdx).String
        sNombreTareaBorrar = oDatos.getCellByPosition(1, filaDatosIdx).String
        
        ' Cuadro confirmatorio crítico
        respuesta = MsgBox("⚠️ ¿Estás seguro de eliminar permanentemente esta tarea? ⚠️" & Chr(13) & Chr(13) & _
                           "📌 Actividad: " & sActividadBorrar & Chr(13) & _
                           "📝 Tarea: """ & sNombreTareaBorrar & """" & Chr(13) & _
                           "🆔 ID Único: " & idActivo & Chr(13) & Chr(13) & _
                           "Esta acción eliminará la fila por completo de la base de datos. Es irreversible.", 4 + 16 + 256, "Confirmar Eliminación Permanente")
                           
        If respuesta <> 6 Then Exit Sub ' Si cancela, salimos
        
        ' Bloquear actualización de pantalla
        oDoc.addActionLock()
        
        ' Eliminar fila de la base de datos
        oDatos.Rows.removeByIndex(filaDatosIdx, 1)
        
        ' Limpiar descripciones residuales en el visor (Celda D4)
        EjecutarInyeccionDescripcion(oVisor, 0)
        
        ' Reordenar secuencialmente y redibujar el lienzo Kanban
        OrdenarDatosPersonalizado()
        
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
    oVisor = oDoc.Sheets.getByName("visor")
    oSel = oDoc.CurrentSelection
    
    ' 1. Detectar el ID de la tarea seleccionada en el visor
    idActivo = 0
    If oSel.supportsService("com.sun.star.sheet.SheetCell") Then
        idActivo = ObtenerIdDesdePosicion(oVisor, oSel.CellAddress.Row, oSel.CellAddress.Column)
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
    binFocoCongelado = True
    
    ' 3. Invocar al Hub de Lectura con el ID seleccionado (Esto mostrará BtnEliminar y BtnPapelera)
    Call InicializarServiciosDialogo(oDialogo, idActivo)
    
    ' 4. Ejecución del diálogo
    If oDialogo.execute() = 1 Then
        ' Guardar las modificaciones sobre el registro existente
        Call GuardarDatosDesdeDialogo(oDialogo, idActivo)
    End If
    
    ' Liberar recursos de forma segura
    binFocoCongelado = False
    oDialogo.dispose()
End Sub

REM =========================================================================
REM MACRO DE CREACIÓN: REGISTRAR NUEVA TAREA (USANDO DIÁLOGO UNIFICADO)
REM =========================================================================
Sub RegistrarNuevaTarea()
    Dim oDoc As Object, oVisor As Object, oDatos As Object
    Dim oDialogo As Object
    Dim sActividadFiltro As String, sEstadoFiltro As String
    
    oDoc = ThisComponent
    oVisor = oDoc.Sheets.getByName("visor")
    oDatos = oDoc.Sheets.getByName("datos")
    
    ' 1. Cargar e instanciar el diálogo unificado "DialogoPanelTarea"
    DialogLibraries.LoadLibrary("Standard")
    oDialogo = CreateUnoDialog(DialogLibraries.Standard.DialogoPanelTarea)
    
    ' Congelar el listener temporalmente para evitar parpadeos visuales
    binFocoCongelado = True
    
    ' 2. Invocar al Hub de Lectura con ID = 0 (Esto ocultará automáticamente BtnEliminar y BtnPapelera)
    Call InicializarServiciosDialogo(oDialogo, 0)
    
    ' 3. Pre-selección inteligente basada en los filtros activos del Visor
    sActividadFiltro = Trim(oVisor.getCellByPosition(0, 1).String) ' Celda A2
    sEstadoFiltro = Trim(oVisor.getCellByPosition(0, 2).String)    ' Celda A3
    
    If sActividadFiltro <> "" Then
        oDialogo.getControl("CmbActividad").setText(sActividadFiltro)
    End If
    
    If sEstadoFiltro <> "" Then
        oDialogo.getControl("CmbEstado").setText(sEstadoFiltro)
    Else
        oDialogo.getControl("CmbEstado").setText("En curso")
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
    binFocoCongelado = False
    oDialogo.dispose()
End Sub

REM =========================================================================
REM PUNTO DE ENTRADA: GESTIONAR PROYECTOS / ACTIVIDADES
REM =========================================================================
Sub GestionarActividades()
    G_TipoItemGestionado = "Actividad"
    Call AbrirPanelItem()
End Sub

REM =========================================================================
REM PUNTO DE ENTRADA: GESTIONAR COLUMNAS / ESTADOS
REM =========================================================================
Sub GestionarEstados()
    G_TipoItemGestionado = "Estado"
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
    oDialogo.Title = "Gestión de " & G_TipoItemGestionado & "es"
    
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
    Dim vLista() As String
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
    
    ' 2. Verificar si el texto ya existe en la lista del combo
    vLista() = oCmbItem.Model.StringItemList
    binExiste = ElementoExisteEnLista(sTexto, vLista())
    
    ' 3. Aplicar mutación visual dinámica al botón
    If binExiste Then
        ' --- MODO ELIMINACIÓN (Existe en la lista) ---
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
    Dim oDialogo As Object, oDatos As Object, oCmbItem As Object
    Dim sTexto As String
    Dim vLista() As String
    Dim binExiste As Boolean
    Dim respuesta As Integer
    
    oDialogo = oEvent.Source.Context
    oDatos = ThisComponent.Sheets.getByName("datos")
    oCmbItem = oDialogo.getControl("CmbItem")
    sTexto = Trim(oCmbItem.Text)
    
    If sTexto = "" Then Exit Sub
    
    vLista() = oCmbItem.Model.StringItemList
    binExiste = ElementoExisteEnLista(sTexto, vLista())
    
    If binExiste Then
        ' OPERACIÓN: ELIMINACIÓN EN CASCADA
        oDialogo.endExecute()
        Call EjecutarBorradoEnCascada(sTexto, G_TipoItemGestionado)
    Else
        ' OPERACIÓN: AGREGAR NUEVO ELEMENTO
        respuesta = MsgBox("¿Deseas agregar la nueva " & G_TipoItemGestionado & ": """ & sTexto & """?", 4 + 32, "Confirmar Adición")
        
        If respuesta = 6 Then ' SÍ
            Dim mDatos() As Variant, i As Long, filaVacia As Long
            Dim nColIdx As Long
            
            filaVacia = -1
            If G_TipoItemGestionado = "Actividad" Then
                nColIdx = 7
                mDatos = oDatos.getCellRangeByName("H2:H100").getDataArray()
            Else
                nColIdx = 9
                mDatos = oDatos.getCellRangeByName("J2:J20").getDataArray()
            End If
            
            For i = 0 To UBound(mDatos)
                If Trim(mDatos(i)(0)) = "" Then
                    filaVacia = i + 2
                    Exit For
                End If
            Next i
            
            If filaVacia <> -1 Then
                oDatos.getCellByPosition(nColIdx, filaVacia - 1).String = sTexto
                oDialogo.endExecute()
                
                OrdenarDatosPersonalizado()
                MsgBox G_TipoItemGestionado & " agregada correctamente al catálogo.", 64, "Éxito"
            Else
                MsgBox "Error: El catálogo de " & G_TipoItemGestionado & "es está lleno.", 48, "Límite Excedido"
            End If
        End If
    End If
End Sub

REM =========================================================================
REM FUNCIÓN AUXILIAR: COMPARACIÓN INSENSIBLE A MAYÚSCULAS EN ARREGLOS
REM =========================================================================
Function ElementoExisteEnLista(sBuscar As String, ByRef vLista() As String) As Boolean
    Dim i As Long
    sBuscar = LCase(Trim(sBuscar))
    For i = 0 To UBound(vLista)
        If LCase(Trim(vLista(i))) = sBuscar Then
            ElementoExisteEnLista = True
            Exit Function
        End If
    Next i
    ElementoExisteEnLista = False
End Function
