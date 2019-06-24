Option Explicit
Function comision_para_la_empresa(fecha As Date, paquete As String) As Double

    'Declaración de variables
    Dim Precios As Worksheet, i As Integer
    
    Set Precios = Sheets("Precios")
    
    'Recibe la fecha y el paquete a contratar. Devuelve la comision en soles para la empresa Manos por esta venta.
    If Weekday(fecha) = vbSunday Then
        For i = 1 To 6
            If Precios.Cells(10, 8 + i) = paquete Then
                comision_para_la_empresa = Precios.Cells(12, 8 + i)
                Exit For
            End If
        Next i
    Else
        For i = 1 To 6
            If Precios.Cells(10, 2 + i) = paquete Then
                comision_para_la_empresa = Precios.Cells(12, 2 + i)
                Exit For
            End If
        Next i
    End If
        
End Function

Function pago_al_trabajador(fecha As Date, paquete As String) As Double

    'Declaración de variables
    Dim Precios As Worksheet, i As Integer
    
    Set Precios = Sheets("Precios")
    
    'Recibe la fecha y el paquete a contratar. Devuelve la comision en soles para la empresa Manos por esta venta.
    If Weekday(fecha) = vbSunday Then
        For i = 1 To 6
            If Precios.Cells(10, 8 + i) = paquete Then
                pago_al_trabajador = Precios.Cells(13, 8 + i)
            End If
        Next i
    Else
        For i = 1 To 6
            If Precios.Cells(10, 2 + i) = paquete Then
                pago_al_trabajador = Precios.Cells(13, 2 + i)
            End If
        Next i
    End If
    
End Function

Function numero_de_mes(mes As String) As Integer
    'Recibe un cadena de texto de un mes. Devuelve que número de mes es.
    Select Case mes
        Case "Enero"
            numero_de_mes = 1
        Case "Febrero"
            numero_de_mes = 2
        Case "Marzo"
            numero_de_mes = 3
        Case "Abril"
            numero_de_mes = 4
        Case "Mayo"
            numero_de_mes = 5
        Case "Junio"
            numero_de_mes = 6
        Case "Julio"
            numero_de_mes = 7
        Case "Agosto"
            numero_de_mes = 8
        Case "Setiembre"
            numero_de_mes = 9
        Case "Octubre"
            numero_de_mes = 10
        Case "Noviembre"
            numero_de_mes = 11
        Case "Diciembre"
            numero_de_mes = 12
    End Select
            
End Function

Function fecha_de_servicio_es_valida(dia As String, mes As String, anho As String) As Boolean

    'Recibe un dia, un mes y un año. Devuelve True si esta fecha es válida.

    Dim fecha_de_servicio, fecha_actual As Date
    
    fecha_de_servicio = DateValue(dia & "/" & mes & "/" & anho)
    fecha_actual = Date
    
    If fecha_de_servicio > fecha_actual Then
        'Una fecha de servicio es valida si es mayor al dia de hoy
        fecha_de_servicio_es_valida = True
    Else
        fecha_de_servicio_es_valida = False
    End If
    
End Function

Sub resetear_mensajes_de_error(formulario As Worksheet)
    'Borra los mensajes de error en el formulario
    formulario.Cells(11, 2) = ""
    formulario.Cells(11, 4) = ""
    formulario.Cells(11, 6) = ""
    formulario.Cells(14, 2) = ""
    formulario.Cells(14, 5) = ""
    formulario.Cells(17, 2) = ""
    formulario.Cells(17, 5) = ""
End Sub

Sub resetear_campos(formulario As Worksheet)
    'Borra todos los datos de todos los campos en el formulario
    formulario.Cells(10, 2) = ""
    formulario.Cells(10, 4) = ""
    formulario.Cells(10, 6) = "2019"
    formulario.Cells(13, 2) = ""
    formulario.Cells(13, 5) = ""
    formulario.Cells(16, 2) = ""
    formulario.Cells(16, 5) = ""
End Sub

Sub anadir_nuevo_boton()
    Dim formulario As Worksheet
    Set formulario = Sheets("Servicio_nuevo")
    'Borra todos los datos de todos los campos en el formulario
    Call resetear_campos(formulario)
    'Borra los mensajes de error en el formulario
    Call resetear_mensajes_de_error(formulario)
    'DespuŽs de borrar todo, redirecciona al formulario
    formulario.Visible = True
    formulario.Activate
    
End Sub

Sub cancelar_boton()
    'Redirecciona al usuario a la pantalla de inicio
    Sheets("Servicios").Activate
    'Oculta el formulario
    Sheets("Servicio_nuevo").Visible = False
    Sheets("Servicio_editar").Visible = False
End Sub

Sub consultar()
    Sheets("Servicios_datos").Visible = True
    Sheets("Servicios_datos").Activate
End Sub

Sub editar_boton()
    'En caso de error, se activa un ErrorHandler
    On Error GoTo ErrorHandler
    Dim Servicio As Worksheet, ID_servicio As Integer
    Set Servicio = Sheets("Servicio_editar")
    ID_servicio = InputBox("Introduzca el ID del servicio que desea editar")
ErrorHandler:
    Select Case Err.Number
        'El error 13 se produce cuando el dato ingresado no es un integer
        Case 13
            'Se le indica al usuario que el dato es erróneo y que intente nuevamente
            MsgBox ("No ingresaste un número. Intenta nuevamente.")
            Exit Sub
    End Select
    'Pide al usuario el ID del Servicio que desea editar
    If Validadores.ID_servicio_es_valido(ID_servicio) Then
        'Se borran los datos
        Call resetear_mensajes_de_error(Servicio)
        Call resetear_campos(Servicio)
        'Se llena el formulaio de edición con los datos actuales
        Call completar_datos(ID_servicio)
        'Se establece el formulario de edición como Visible
        Servicio.Visible = True
        'Se redirecciona al usuario al formulario de edición
        Servicio.Activate
    Else
        'Si no se encuentra un ID que coincida, se muestra el error y el ID ingresado
        MsgBox ("No se encontró ningun servicio que coincida con el ID = " & CStr(ID_servicio))
    End If
    
End Sub

Function datos_servicio_son_validos(dia As String, mes As String, anho As String, hora_de_inicio As String, paquete_a_contratar As String, ejecutivo_de_venta As String, ID_cliente As Integer) As Boolean
    'Evalua si los datos del trabajador son válidos
    'Si no son válidos, devuelve falso; caso contrario, devuelve Verdadero
    'Por cada dato mal escrito, se muestra un mensaje de error
    
    Set formulario = Sheets("Servicio_nuevo")
    datos_servicio_son_validos = True
    If Not (Validadores.texto_es_valido(mes)) Then
        formulario.Cells(11, 2) = "Debe indicar el dia"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(mes)) Then
        formulario.Cells(11, 4) = "Debe indicar el mes."
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(anho)) Then
        formulario.Cells(11, 6) = "Debe indicar el año"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    If datos_servicio_son_validos Then
        If Not (fecha_de_servicio_es_valida(dia, mes, anho)) Then
            formulario.Cells(11, 2) = "La fecha es inv‡lida"
            datos_servicio_son_validos = datos_servicio_son_validos * False
        End If
    End If
    If Not (Validadores.texto_es_valido(hora_de_inicio)) Then
        formulario.Cells(14, 2) = "Debe especificar la hora de inicio"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(paquete_a_contratar)) Then
        formulario.Cells(14, 5) = "Debe especificar el paquete"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(ejecutivo_de_venta)) Then
        formulario.Cells(17, 2) = "Debe especificar el ejecutivo de venta"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    
    If Not (Validadores.ID_cliente_es_valido(CInt(ID_cliente))) Then
        formulario.Cells(17, 5) = "Ningún cliente coincide con este ID"
        datos_servicio_son_validos = datos_servicio_son_validos * False
    End If
    
End Function

Function datos_servicio_edicion_son_validos(dia As String, mes As String, anho As String, hora_de_inicio As String, paquete_a_contratar As String, estado As String, ID_trabajador As String) As Boolean

    Set formulario = Sheets("Servicio_editar")
    datos_servicio_edicion_son_validos = True
    If Not (Validadores.texto_es_valido(mes)) Then
        formulario.Cells(11, 2) = "Debe indicar el dia"
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(mes)) Then
        formulario.Cells(11, 4) = "Debe indicar el mes."
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(anho)) Then
        formulario.Cells(11, 6) = "Debe indicar el año"
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    If datos_servicio_edicion_son_validos Then
        If Not (fecha_de_servicio_es_valida(dia, mes, anho)) Then
            formulario.Cells(11, 2) = "La fecha es inv‡lida"
            datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
        End If
    End If
    If Not (Validadores.texto_es_valido(hora_de_inicio)) Then
        formulario.Cells(14, 2) = "Debe especificar la hora de inicio"
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(paquete_a_contratar)) Then
        formulario.Cells(14, 5) = "Debe especificar el paquete"
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    If Not (Validadores.texto_es_valido(estado)) Then
        formulario.Cells(17, 2) = "Debe especificar el ejecutivo de venta"
        datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
    End If
    
    If ID_trabajador <> "" Then
        If Not (Validadores.ID_trabajador_es_valido(CInt(ID_trabajador))) Then
            formulario.Cells(17, 5) = "Ningún trabajador coincide con este ID"
            datos_servicio_edicion_son_validos = datos_servicio_edicion_son_validos * False
        End If
    End If

End Function


Function el_servicio_es_en_la_manana(hora_de_inicio As Variant, numero_de_horas As Integer) As Boolean
    If Hour(hora_de_inicio) < 12 Then
        el_servicio_es_en_la_manana = True
    Else
        el_servicio_es_en_la_manana = False
    End If
    
End Function

Function el_servicio_es_en_la_tarde(hora_de_inicio As Variant, numero_de_horas As Integer) As Boolean
    If Hour(hora_de_inicio) + numero_de_horas > 12 And Hour(hora_de_inicio) < 18 Then
        el_servicio_es_en_la_tarde = True
    Else
        el_servicio_es_en_la_tarde = False
    End If
    
End Function

Function el_servicio_es_en_la_noche(hora_de_inicio As Variant, numero_de_horas As Integer) As Boolean
    If Hour(hora_de_inicio) + numero_de_horas > 18 Then
        el_servicio_es_en_la_noche = True
    Else
        el_servicio_es_en_la_noche = False
    End If
End Function

Sub mostrar_trabajadores_potenciales(hora_de_inicio As Variant, numero_de_horas As Integer)

    'Declaración de variables
    Dim Trabajadores_todos As Worksheet, Trabajadores_potenciales As Worksheet
    Dim servicio_en_la_manana As Boolean, servicio_en_la_tarde As Boolean, servicio_en_la_noche As Boolean
    Dim trabajador_disponible_en_la_manana As Boolean, trabajador_disponible_en_la_tarde As Boolean
    Dim trabajador_disponible_en_la_noche As Boolean, manana As Boolean
    Dim tarde As Boolean, noche As Boolean
    Dim ID_trabajador As Integer, numero_de_trabajadores_potenciales As Integer, i As Integer
    
    Set Trabajadores_todos = Sheets("Trabajadores_datos")
    Set Trabajadores_potenciales = Sheets("Trabajadores_potenciales")
    
    'Se borran los trabajadores potenciales escritos anteriormente
    Trabajadores_potenciales.Range("B9:F28").Clear
    
    servicio_en_la_manana = el_servicio_es_en_la_manana(hora_de_inicio, numero_de_horas)
    servicio_en_la_tarde = el_servicio_es_en_la_tarde(hora_de_inicio, numero_de_horas)
    servicio_en_la_noche = el_servicio_es_en_la_noche(hora_de_inicio, numero_de_horas)
    i = 2
    numero_de_trabajadores_potenciales = 0
    Do While Trabajadores_todos.Cells(i, 1) <> ""
        trabajador_disponible_en_la_manana = Trabajadores_todos.Cells(i, 7)
        trabajador_disponible_en_la_tarde = Trabajadores_todos.Cells(i, 8)
        trabajador_disponible_en_la_noche = Trabajadores_todos.Cells(i, 9)
        manana = Not (servicio_en_la_manana) Or trabajador_disponible_en_la_manana
        tarde = Not (servicio_en_la_tarde) Or trabajador_disponible_en_la_tarde
        noche = Not (servicio_en_la_noche) Or trabajador_disponible_en_la_noche
        If manana And tarde And noche Then
            ID_trabajador = Trabajadores_todos.Cells(i, 1)
            Trabajadores_potenciales.Cells(9 + numero_de_trabajadores_potenciales, 2) = Trabajadores.nombre_completo(ID_trabajador)
            Trabajadores_potenciales.Cells(9 + numero_de_trabajadores_potenciales, 5) = Trabajadores.celular(ID_trabajador)
            numero_de_trabajadores_potenciales = numero_de_trabajadores_potenciales + 1
        End If
            
        i = i + 1
    Loop
    
    Trabajadores_potenciales.Activate
    MsgBox (CStr(numero_de_trabajadores_potenciales) & " trabajadores potenciales encontrados.")
    
End Sub

Sub buscar_trabajadores_potenciales_boton()
    Dim hora_de_inicio As Variant, numero_de_horas As Integer, formulario As Worksheet
    Set formulario = Sheets("Servicio_editar")
    
    'La hora de inicio esta indicada como un texto y hay que pasarlo a TimeValue
    hora_de_inicio = TimeValue(formulario.Cells(13, 2))
    'El numero de horas es el primer caracter de la cadena Paquete
    numero_de_horas = CInt(Left(formulario.Cells(13, 5), 1))
    
    'Según la hora de inicio y número de horas, se mostrará al usuario los trabajadores potenciales
    Call mostrar_trabajadores_potenciales(hora_de_inicio, numero_de_horas)
    
End Sub

Sub completar_datos(ID_servicio As Integer)

    'Recibe un ID de servicio. Asume que este dato es correcto. Escribe los datos actuales del servicio en el formulario de edición
    
    'Declaración de variables
    Dim formulario As Worksheet, datos As Worksheet
    Dim fila As Integer
    
    Set formulario = Sheets("Servicio_editar")
    Set datos = Sheets("Servicios_datos")
    fila = ID_servicio + 1
    
    'Escritura de datos
    formulario.Cells(7, 2) = ID_servicio
    formulario.Cells(10, 2) = Day(datos.Cells(fila, 4))
    formulario.Cells(10, 4) = Format(datos.Cells(fila, 4), "mmmm")
    formulario.Cells(10, 6) = Year(datos.Cells(fila, 4))
    formulario.Cells(13, 2) = datos.Cells(fila, 5)
    formulario.Cells(13, 5) = datos.Cells(fila, 6)
    formulario.Cells(16, 2) = datos.Cells(fila, 2)
    formulario.Cells(16, 5) = datos.Cells(fila, 12)
    
    
End Sub

Sub Agregar_nuevo_servicio(dia As String, mes As String, anho As String, hora_de_inicio As String, paquete_a_contratar As String, ejecutivo_de_ventas As String, ID_cliente As Integer):
    
    'Declaración de variables
    Dim fecha_de_servicio As Date, x As Integer, O As Range, ID_servicio As Integer, precio_total As Double
        
    x = 0
    Set O = Sheets("Servicios_datos").Range("A1")
    Do
        x = x + 1
    Loop Until O.Offset(x, 0) = ""
    fecha_de_servicio = DateValue(CStr(numero_de_mes(mes)) & "/" & dia & "/" & anho)
    ID_servicio = x
    O.Offset(x, 0) = ID_servicio
    O.Offset(x, 1) = "Falta asignar trabajador"
    O.Offset(x, 2) = Date
    O.Offset(x, 3) = fecha_de_servicio
    O.Offset(x, 4) = hora_de_inicio
    O.Offset(x, 5) = paquete_a_contratar
    O.Offset(x, 6) = ejecutivo_de_ventas
    O.Offset(x, 7) = ID_cliente
    O.Offset(x, 8) = Cliente.Direccion(ID_cliente)
    O.Offset(x, 9) = Cliente.Referencias(ID_cliente)
    O.Offset(x, 10) = Cliente.Distrito(ID_cliente)
    O.Offset(x, 12) = comision_para_la_empresa(fecha_de_servicio, paquete_a_contratar)
    O.Offset(x, 13) = pago_al_trabajador(fecha_de_servicio, paquete_a_contratar)
    precio_total = comision_para_la_empresa(fecha_de_servicio, paquete_a_contratar) + pago_al_trabajador(fecha_de_servicio, paquete_a_contratar)
    O.Offset(x, 14) = precio_total
    
End Sub

Sub Actualizar_servicio(ID_servicio As Integer, dia As String, mes As String, anho As String, hora_de_inicio As String, paquete_a_contratar As String, estado As String, ID_trabajador As String)

    'Recibe los datos del servicio y los actualiza en la base de datos. Asume que estos datos son correctos
    
    'Declaracion de variables
    Dim Servicios As Worksheet, fecha_de_servicio As Date, precio_total As Double, fila As Integer

    Set Servicios = Sheets("Servicios_datos")
    fila = ID_servicio + 1
    fecha_de_servicio = DateValue(CStr(numero_de_mes(mes)) & "/" & dia & "/" & anho)
    'Escritura de datos
    If ID_trabajador <> "" And estado <> "Cancelado" Then
        estado = "Trabajador asignado"
    ElseIf ID_trabajador = "" And estado <> "Cancelado" Then
        estado = "Falta asignar trabajador"
    End If
    
    Servicios.Cells(fila, 2) = estado
    Servicios.Cells(fila, 4) = fecha_de_servicio
    Servicios.Cells(fila, 5) = hora_de_inicio
    Servicios.Cells(fila, 6) = paquete_a_contratar
    Servicios.Cells(fila, 12) = ID_trabajador
    Servicios.Cells(fila, 13) = comision_para_la_empresa(fecha_de_servicio, paquete_a_contratar)
    Servicios.Cells(fila, 14) = pago_al_trabajador(fecha_de_servicio, paquete_a_contratar)
    precio_total = comision_para_la_empresa(fecha_de_servicio, paquete_a_contratar) + pago_al_trabajador(fecha_de_servicio, paquete_a_contratar)
    Servicios.Cells(fila, 15) = precio_total
    
End Sub


Sub Formulario_edicion_servicio():

    Dim formulario As Worksheet
    Set formulario = Sheets("Servicio_editar")
    'El dia y el año son String porque cabe la posibilidad que el usuario  no lo escriba y lo deje vacío
    Dim dia As String, mes As String, anho As String
    Dim hora_de_inicio As String, paquete_a_contratar As String, estado As String
    Dim horas As Integer, hora_inicio As Variant, ID_trabajador As String, ID_servicio As Integer
    
    'Se deben borrar los mensajes de error
    Call Servicios.resetear_mensajes_de_error(formulario)
    
    'Lectura de datos
    ID_servicio = formulario.Cells(7, 2)
    dia = formulario.Cells(10, 2)
    mes = formulario.Cells(10, 4)
    anho = formulario.Cells(10, 6)
    hora_de_inicio = formulario.Cells(13, 2)
    paquete_a_contratar = formulario.Cells(13, 5)
    estado = formulario.Cells(16, 2)
    ID_trabajador = formulario.Cells(16, 5)
    
    If datos_servicio_edicion_son_validos(dia, mes, anho, hora_de_inicio, paquete_a_contratar, estado, ID_trabajador) Then
        'Si se han ingresado bien los datos, se pasa a añadir el cliente a la tabla de clientes
        Call Actualizar_servicio(ID_servicio, dia, mes, anho, hora_de_inicio, paquete_a_contratar, estado, ID_trabajador)
        'Se muestra un mensaje al usuario que se agregó exitosamente
        MsgBox ("Servicio actualizado exitosamente")
        'Se borran los datos que puso el cliente
        Call resetear_campos(formulario)
        'Se oculta el formulario
        formulario.Visible = False
    Else
        MsgBox ("Datos incorrectos. Revise las observaciones.")
    End If
    
End Sub

Sub Formulario_nuevo_servicio():
    Dim formulario As Worksheet
    Set formulario = Sheets("Servicio_nuevo")
    'El dia y el a–o son String porque cabe la posibilidad que el usuario  no lo escriba y lo deje vac’o
    Dim dia As String, mes As String, anho As String
    Dim hora_de_inicio As String, paquete_a_contratar As String, ejecutivo_de_ventas As String
    Dim horas As Integer, hora_inicio As Variant, ID_cliente As Integer
    
    'Se deben borrar los mensajes de error
    Call Servicios.resetear_mensajes_de_error(formulario)
    dia = formulario.Cells(10, 2)
    mes = formulario.Cells(10, 4)
    anho = formulario.Cells(10, 6)
    hora_de_inicio = formulario.Cells(13, 2)
    paquete_a_contratar = formulario.Cells(13, 5)
    ejecutivo_de_ventas = formulario.Cells(16, 2)
    ID_cliente = formulario.Cells(16, 5)
    If datos_servicio_son_validos(dia, mes, anho, hora_de_inicio, paquete_a_contratar, ejecutivo_de_ventas, ID_cliente) Then
        'Si se han ingresado bien los datos, se pasa a a–adir el cliente a la tabla de clientes
        Call Agregar_nuevo_servicio(dia, mes, anho, hora_de_inicio, paquete_a_contratar, ejecutivo_de_ventas, ID_cliente)
        'Se muestra un mensaje al usuario que se a–adi— exitosamente
        MsgBox ("Servicio añadido exitosamente")
        'Se muestran los trabajadores que podr’an estar disponibles para este servicio
        horas = CInt(Left(paquete_a_contratar, 1))
        hora_inicio = TimeValue(hora_de_inicio)
        Call mostrar_trabajadores_potenciales(hora_inicio, horas)
        'Se borran los datos que puso el cliente
        Call resetear_mensajes_de_error(formulario)
        Call resetear_campos(formulario)
        'Se oculta el formulario
        formulario.Visible = False
    Else
        MsgBox ("Datos incorrectos. Revise las observaciones.")
    End If
End Sub
