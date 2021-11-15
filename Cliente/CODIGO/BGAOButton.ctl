VERSION 5.00
Begin VB.UserControl BGAOButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   Begin VB.Label CMDText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BGAOButton"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "BGAOButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BGAOButtonStyle
    s_Small = 0
    s_Normal = 1
    s_Large = 2
    b_Normal = 3
    b_Large = 4
    sb_Normal = 5
End Enum
'
Private BGTipo As BGAOButtonStyle

'Default Property Values:
Const m_def_ToolTipText = ""
'Property Variables:
Dim m_ToolTipText As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse y después lo vuelve a presionar y liberar sobre un objeto."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."

Private ThisForm As Form

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Devuelve o establece el texto mostrado en la barra de título de un objeto o bajo el icono de un objeto."
    Caption = CMDText.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    CMDText.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = CMDText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set CMDText.Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Devuelve o establece el estilo negrita de una fuente."
    FontBold = CMDText.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    CMDText.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Devuelve o establece el estilo cursiva de una fuente."
    FontItalic = CMDText.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    CMDText.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Especifica el nombre de la fuente que aparece en cada fila del nivel especificado."
    FontName = CMDText.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    CMDText.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CMDText,CMDText,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Especifica el tamaño (en puntos) de la fuente que aparece en cada fila del nivel especificado."
    FontSize = CMDText.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    CMDText.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property


Private Sub CMDText_Click()
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un objeto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Cambia el ancho y el alto de un control de usuario."
    UserControl.Size Width, Height
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Devuelve o establece el texto mostrado cuando el mouse se sitúa sobre un control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Function Init(ByVal TIPO As BGAOButtonStyle, Optional ByVal Text As String = vbNullString) As Variant
    BGTipo = TIPO
    If Text <> vbNullString Then Caption = Text
    
    Select Case BGTipo
        Case BGAOButtonStyle.s_Small
            UserControl.Width = 20 * 15
            UserControl.Height = 20 * 15
        Case BGAOButtonStyle.s_Normal
            UserControl.Width = 100 * 15
            UserControl.Height = 19 * 15
        Case BGAOButtonStyle.s_Large
            UserControl.Width = 200 * 15
            UserControl.Height = 19 * 15
        Case BGAOButtonStyle.b_Normal
            UserControl.Width = 100 * 15
            UserControl.Height = 31 * 15
        Case BGAOButtonStyle.b_Large
            UserControl.Width = 200 * 15
            UserControl.Height = 31 * 15
        Case BGAOButtonStyle.sb_Normal
            UserControl.Width = 70 * 15
            UserControl.Height = 31 * 15
    End Select

    UserControl.Picture = General_Set_GUI("b" & BGTipo + 1)
    CMDText.Width = (UserControl.Width / 15)
    CMDText.Top = (UserControl.Height / 15 - CMDText.Height) / 2
End Function

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    CMDText.Caption = PropBag.ReadProperty("Caption", "BGAOButton")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set CMDText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    CMDText.FontBold = PropBag.ReadProperty("FontBold", 0)
    CMDText.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    CMDText.FontName = PropBag.ReadProperty("FontName", "Verdana")
    CMDText.FontSize = PropBag.ReadProperty("FontSize", 8)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", CMDText.Caption, "BGAOButton")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", CMDText.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", CMDText.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", CMDText.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", CMDText.FontName, "")
    Call PropBag.WriteProperty("FontSize", CMDText.FontSize, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
End Sub

