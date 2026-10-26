Attribute VB_Name = "modControlInventarios"
Option Explicit

' Control de Inventarios - Excel VBA
' Proyecto educativo para pequeñas y medianas operaciones.
' Hojas utilizadas: Productos, Movimientos, Dashboard

Private Const SH_PRODUCTOS As String = "Productos"
Private Const SH_MOVIMIENTOS As String = "Movimientos"
Private Const SH_DASHBOARD As String = "Dashboard"

Public Sub ConfigurarSistema()
    Application.ScreenUpdating = False
    CrearHojaProductos
    CrearHojaMovimientos
    CrearHojaDashboard
    ActualizarDashboard
    Application.ScreenUpdating = True
    MsgBox "Sistema configurado. Registra productos y luego ingresa movimientos.", vbInformation, "Control de Inventarios"
End Sub

Private Sub CrearHojaProductos()
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHoja(SH_PRODUCTOS)
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:G1").Value = Array("SKU", "Producto", "Categoría", "Unidad", "Stock mínimo", "Costo unitario", "Activo")
        ws.Range("A2:G2").Value = Array("SKU-001", "Producto de ejemplo", "General", "UND", 10, 5, "Sí")
    End If
    FormatearEncabezado ws, "A1:G1"
    ws.Columns("A:G").AutoFit
    ws.Columns("F").NumberFormat = """S/ ""#,##0.00"
    ws.Columns("E").NumberFormat = "0"
    ws.Range("A1:G1").AutoFilter
End Sub

Private Sub CrearHojaMovimientos()
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHoja(SH_MOVIMIENTOS)
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:H1").Value = Array("Fecha", "Tipo", "SKU", "Producto", "Cantidad", "Costo unitario", "Documento", "Responsable")
    End If
    FormatearEncabezado ws, "A1:H1"
    ws.Columns("A:H").AutoFit
    ws.Columns("A").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("F").NumberFormat = """S/ ""#,##0.00"
    ws.Range("A1:H1").AutoFilter
End Sub

Private Sub CrearHojaDashboard()
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHoja(SH_DASHBOARD)
    ws.Cells.Clear
    ws.Range("A1:B1").Merge
    ws.Range("A1").Value = "DASHBOARD | CONTROL DE INVENTARIOS"
    ws.Range("A3").Value = "Indicador"
    ws.Range("B3").Value = "Valor"
    ws.Range("A4").Value = "Productos activos"
    ws.Range("A5").Value = "Unidades en stock"
    ws.Range("A6").Value = "Valor estimado del inventario"
    ws.Range("A7").Value = "Productos bajo stock mínimo"
    FormatearEncabezado ws, "A3:B3"
    With ws.Range("A1:B1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns("A:B").AutoFit
End Sub

Public Sub RegistrarProducto()
    Dim ws As Worksheet, sku As String, nombre As String, categoria As String
    Dim unidad As String, minimo As Double, costo As Double, fila As Long
    Set ws = ObtenerOCrearHoja(SH_PRODUCTOS)
    If ws.Range("A1").Value = "" Then CrearHojaProductos

    sku = Trim(InputBox("Ingrese el SKU o código único del producto:", "Nuevo producto"))
    If sku = "" Then Exit Sub
    If ExisteSKU(sku) Then
        MsgBox "Ese SKU ya existe. Use un código único.", vbExclamation
        Exit Sub
    End If
    nombre = Trim(InputBox("Nombre del producto:", "Nuevo producto"))
    If nombre = "" Then Exit Sub
    categoria = Trim(InputBox("Categoría:", "Nuevo producto", "General"))
    unidad = Trim(InputBox("Unidad de medida (UND, KG, LT, etc.):", "Nuevo producto", "UND"))
    minimo = Val(InputBox("Stock mínimo:", "Nuevo producto", "0"))
    costo = Val(InputBox("Costo unitario (S/):", "Nuevo producto", "0"))

    fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(fila, 1).Value = sku
    ws.Cells(fila, 2).Value = nombre
    ws.Cells(fila, 3).Value = categoria
    ws.Cells(fila, 4).Value = unidad
    ws.Cells(fila, 5).Value = minimo
    ws.Cells(fila, 6).Value = costo
    ws.Cells(fila, 7).Value = "Sí"
    ws.Columns("A:G").AutoFit
    ActualizarDashboard
    MsgBox "Producto registrado: " & nombre, vbInformation
End Sub

Public Sub RegistrarEntrada()
    RegistrarMovimiento "Entrada"
End Sub

Public Sub RegistrarSalida()
    RegistrarMovimiento "Salida"
End Sub

Private Sub RegistrarMovimiento(ByVal tipoMovimiento As String)
    Dim sku As String, documento As String, responsable As String
    Dim cantidad As Double, costo As Double, nombre As String
    Dim wsM As Worksheet, fila As Long, stockActual As Double
    sku = Trim(InputBox("Ingrese el SKU del producto:", tipoMovimiento & " de inventario"))
    If sku = "" Then Exit Sub
    If Not BuscarProducto(sku, nombre, costo) Then
        MsgBox "No se encontró el SKU. Registre el producto primero.", vbExclamation
        Exit Sub
    End If

    cantidad = Val(InputBox("Cantidad para " & LCase(tipoMovimiento) & ":", tipoMovimiento, "1"))
    If cantidad <= 0 Then
        MsgBox "La cantidad debe ser mayor que cero.", vbExclamation
        Exit Sub
    End If
    If tipoMovimiento = "Salida" Then
        stockActual = ObtenerStock(sku)
        If cantidad > stockActual Then
            MsgBox "Stock insuficiente. Disponible: " & stockActual, vbCritical
            Exit Sub
        End If
    End If
    documento = Trim(InputBox("Documento de referencia (opcional):", tipoMovimiento, ""))
    responsable = Trim(InputBox("Responsable:", tipoMovimiento, Application.UserName))

    Set wsM = ObtenerOCrearHoja(SH_MOVIMIENTOS)
    If wsM.Range("A1").Value = "" Then CrearHojaMovimientos
    fila = wsM.Cells(wsM.Rows.Count, "A").End(xlUp).Row + 1
    wsM.Cells(fila, 1).Value = Now
    wsM.Cells(fila, 2).Value = tipoMovimiento
    wsM.Cells(fila, 3).Value = sku
    wsM.Cells(fila, 4).Value = nombre
    wsM.Cells(fila, 5).Value = cantidad
    wsM.Cells(fila, 6).Value = costo
    wsM.Cells(fila, 7).Value = documento
    wsM.Cells(fila, 8).Value = responsable
    wsM.Cells(fila, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    wsM.Cells(fila, 6).NumberFormat = """S/ ""#,##0.00"
    wsM.Columns("A:H").AutoFit
    ActualizarDashboard
    MsgBox tipoMovimiento & " registrada correctamente.", vbInformation
End Sub

Public Sub ActualizarDashboard()
    Dim wsD As Worksheet, wsP As Worksheet, ultima As Long, i As Long
    Dim sku As String, activo As String, stock As Double, minimo As Double
    Dim activos As Long, unidades As Double, valor As Double, bajoMinimo As Long
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(SH_PRODUCTOS)
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    On Error GoTo 0
    If wsP Is Nothing Or wsD Is Nothing Then Exit Sub

    ultima = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultima
        sku = Trim(CStr(wsP.Cells(i, 1).Value))
        activo = LCase(Trim(CStr(wsP.Cells(i, 7).Value)))
        If sku <> "" And activo <> "no" Then
            activos = activos + 1
            stock = ObtenerStock(sku)
            minimo = Val(wsP.Cells(i, 5).Value)
            unidades = unidades + stock
            valor = valor + stock * Val(wsP.Cells(i, 6).Value)
            If stock < minimo Then bajoMinimo = bajoMinimo + 1
        End If
    Next i
    wsD.Range("B4").Value = activos
    wsD.Range("B5").Value = unidades
    wsD.Range("B6").Value = valor
    wsD.Range("B7").Value = bajoMinimo
    wsD.Range("B6").NumberFormat = """S/ ""#,##0.00"
    wsD.Range("A3:B7").Borders.LineStyle = xlContinuous
    wsD.Columns("A:B").AutoFit
End Sub

Public Sub GenerarReporteStock()
    Dim wsP As Worksheet, wsR As Worksheet, ultima As Long, i As Long, fila As Long
    Dim sku As String, stock As Double
    Set wsP = ObtenerOCrearHoja(SH_PRODUCTOS)
    Set wsR = ObtenerOCrearHoja("ReporteStock")
    wsR.Cells.Clear
    wsR.Range("A1:F1").Value = Array("SKU", "Producto", "Unidad", "Stock actual", "Stock mínimo", "Estado")
    FormatearEncabezado wsR, "A1:F1"
    ultima = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    fila = 2
    For i = 2 To ultima
        sku = Trim(CStr(wsP.Cells(i, 1).Value))
        If sku <> "" Then
            stock = ObtenerStock(sku)
            wsR.Cells(fila, 1).Value = sku
            wsR.Cells(fila, 2).Value = wsP.Cells(i, 2).Value
            wsR.Cells(fila, 3).Value = wsP.Cells(i, 4).Value
            wsR.Cells(fila, 4).Value = stock
            wsR.Cells(fila, 5).Value = Val(wsP.Cells(i, 5).Value)
            wsR.Cells(fila, 6).Value = IIf(stock < Val(wsP.Cells(i, 5).Value), "REPOSICIÓN", "OK")
            If stock < Val(wsP.Cells(i, 5).Value) Then
                wsR.Range(wsR.Cells(fila, 1), wsR.Cells(fila, 6)).Interior.Color = RGB(255, 220, 220)
            End If
            fila = fila + 1
        End If
    Next i
    wsR.Columns("A:F").AutoFit
    wsR.Range("A1:F1").AutoFilter
    wsR.Activate
    MsgBox "Reporte de stock generado en la hoja ReporteStock.", vbInformation
End Sub

Private Function ObtenerStock(ByVal sku As String) As Double
    Dim ws As Worksheet, ultima As Long, i As Long, total As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_MOVIMIENTOS)
    On Error GoTo 0
    If ws Is Nothing Then
        ObtenerStock = 0
        Exit Function
    End If
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultima
        If StrComp(Trim(CStr(ws.Cells(i, 3).Value)), sku, vbTextCompare) = 0 Then
            If LCase(Trim(CStr(ws.Cells(i, 2).Value))) = "entrada" Then
                total = total + Val(ws.Cells(i, 5).Value)
            ElseIf LCase(Trim(CStr(ws.Cells(i, 2).Value))) = "salida" Then
                total = total - Val(ws.Cells(i, 5).Value)
            End If
        End If
    Next i
    ObtenerStock = total
End Function

Private Function BuscarProducto(ByVal sku As String, ByRef nombre As String, ByRef costo As Double) As Boolean
    Dim ws As Worksheet, ultima As Long, i As Long
    Set ws = ObtenerOCrearHoja(SH_PRODUCTOS)
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultima
        If StrComp(Trim(CStr(ws.Cells(i, 1).Value)), sku, vbTextCompare) = 0 Then
            nombre = CStr(ws.Cells(i, 2).Value)
            costo = Val(ws.Cells(i, 6).Value)
            BuscarProducto = True
            Exit Function
        End If
    Next i
    BuscarProducto = False
End Function

Private Function ExisteSKU(ByVal sku As String) As Boolean
    Dim nombre As String, costo As Double
    ExisteSKU = BuscarProducto(sku, nombre, costo)
End Function

Private Function ObtenerOCrearHoja(ByVal nombreHoja As String) As Worksheet
    On Error Resume Next
    Set ObtenerOCrearHoja = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0
    If ObtenerOCrearHoja Is Nothing Then
        Set ObtenerOCrearHoja = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ObtenerOCrearHoja.Name = nombreHoja
    End If
End Function

Private Sub FormatearEncabezado(ByVal ws As Worksheet, ByVal rango As String)
    With ws.Range(rango)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
