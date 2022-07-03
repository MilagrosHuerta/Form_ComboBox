Attribute VB_Name = "a_Funcion_ComboBox"
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                   Form ComboBox                      --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para enviar mensajes con otro formato       --- '
' ---    y poder posicionarlo donde quieras                --- '
' ------------------------------------------------------------ '
Option Explicit
Public Nombre_Mensaje_C, Mensaje_Mostrar_C, Dato_Input As String
Public Titulo_Mensaje_C As Integer
Public Boton_A, Boton_C As String
Public continuar_C As Integer
Public numBotones_C As Byte
Public Posicion_Izda_C, Posicion_Top_C  As Integer
Public Valor_Dato As Variant
Public Lista_Combo As String
Public Separador_Dato As String
Function frmComboBox(texoMensaje As Variant, btnTexto As String, _
                   Optional datoEntrada As Variant, Optional separadorDatoEntrada As String, Optional tituloForm As String, _
                   Optional posLeftForm As Integer, Optional posTopForm As Integer, _
                   Optional textRigth As Boolean, Optional readRigth As Boolean)
' ------------------------------------------------------------------------------------------------- '
' --- Boton de Texto solo puede ser "bO" o "bOC", si se pone otro valor, toma "bOC" por defecto --- '
' ------------------------------------------------------------------------------------------------- '
Dim i, j, N As Integer
Dim longMensajeCorto, longMensajeMedio As Integer
    numBotones_C = Len(btnTexto) - 1
    longMensajeCorto = 160
    longMensajeMedio = 360
    If LCase(btnTexto) = "bo" Then btnTexto = "bO"
    If LCase(btnTexto) = "boc" Then btnTexto = "bOC"
    
    If btnTexto <> "bO" And btnTexto <> "bOC" Then btnTexto = "bOC"

    Select Case btnTexto
    Case "bO"
        Boton_A = "Aceptar"
    Case "bOC"
        Boton_A = "Aceptar"
        Boton_C = "Cancelar"
    Case Else
        End
    End Select
    
    Nombre_Mensaje_C = tituloForm
    Lista_Combo = datoEntrada
    Separador_Dato = separadorDatoEntrada
    Mensaje_Mostrar_C = texoMensaje
    Posicion_Izda_C = posLeftForm
    Posicion_Top_C = posTopForm
        
    Form_ComboBox.Show

End Function

