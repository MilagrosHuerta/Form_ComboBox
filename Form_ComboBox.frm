VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ComboBox 
   Caption         =   "Form InputBox"
   ClientHeight    =   6675
   ClientLeft      =   4110
   ClientTop       =   4470
   ClientWidth     =   14055
   OleObjectBlob   =   "Form_ComboBox.frx":0000
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form_ComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                   Form ComboBox                      --- '
' ------------------------------------------------------------ '
Private Sub UserForm_Initialize()
Dim a, b, i As Integer
Dim Posicion As Integer
Dim Cadena As Variant
Dim Primer_Dato  As Variant
'Dim separadorDatoEntrada As String
' ----------------------------------------------------------- '
' --- Tamaño mensaje en funcion de la longitud del texto ---  '
' ----------------------------------------------------------- '
    If Len(Mensaje_Mostrar_C) < 200 Then
        i = 1
    ElseIf Len(Mensaje_Mostrar_C) < 300 Then
        i = 2
    ElseIf Len(Mensaje_Mostrar_C) < 400 Then
        i = 3
    ElseIf Len(Mensaje_Mostrar_C) < 500 Then
        i = 4
    ElseIf Len(Mensaje_Mostrar_C) < 600 Then
        i = 5
    Else
        i = 6
    End If
    a = 30 * (2 + i)
    b = 208 + 2 * 30 * i
    
    With Form_ComboBox
        .Caption = Nombre_Mensaje_C
        If Posicion_Izda_C = 0 And Posicion_Top_C = 0 Then
            .StartUpPosition = 2     ' Centrar en pantalla
        Else
            .Left = Posicion_Izda_C
            .Top = Posicion_Top_C
        End If
        .Height = b + 26
    End With
    
    frmMensaje_Mostrar.Height = a - 10
    frmCombo_Box.Top = a
    frmBoton_Aceptar.Top = a + 22
    frmBoton_Cancelar.Top = a + 22
    TextoBoton_Aceptar.Top = a + 23.5
    TextoBoton_Cancelar.Top = a + 23.5
    frmMensaje_Mostrar = Mensaje_Mostrar_C
' ----------------------------------------------------------------------- '
' --- Nombre de Botones y visibles en funcion de la variable btnTexto --- '
' ----------------------------------------------------------------------- '
    frmBoton_Aceptar.Visible = True
    TextoBoton_Aceptar.Text = Boton_A

    If numBotones_C = 1 Then
        frmBoton_Cancelar.Visible = False
        TextoBoton_Cancelar.Visible = False
    ElseIf numBotones_C = 2 Then
'  TextoBoton_Cancelar.Text = Boton_C ' Si quires cambiar el texto del botón
        frmBoton_Cancelar.Visible = True
        TextoBoton_Cancelar.Visible = True
        TextoBoton_Cancelar.Text = Boton_C
    End If
    Cadena = Lista_Combo
    If InStr(Cadena, Separador_Dato) = 0 Then
        Primer_Dato = Cadena
    Else
        Primer_Dato = Left(Cadena, InStr(Cadena, Separador_Dato) - 1)
    End If
    With Me.frmCombo_Box
        .Clear
        .DropButtonStyle = fmDropButtonStyleArrow
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        .Style = fmStyleDropDownList
        For i = 1 To Len(Cadena)
            Posicion = InStr(Cadena, Separador_Dato)
            If Posicion = 0 Then Posicion = Len(Cadena) + 1
            
            .AddItem Left(Cadena, Posicion - 1)
            Cadena = Mid(Cadena, Posicion + 1, Len(Cadena))
            If Cadena = "" Then Exit For
        Next i
       .Value = Primer_Dato
    End With
End Sub
Private Sub frmCombo_Box_Change()
    Valor_Dato = Me.frmCombo_Box
End Sub
Private Sub frmBoton_Aceptar_Click()
    Valor_Dato = Me.frmCombo_Box
    continuar_C = vbOK
    Unload Me
End Sub
Private Sub frmBoton_Cancelar_Click()
    continuar_C = vbCancel
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Macro para evitar que se pueda cerrar el formulario en la X de arriba a la derecha
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
