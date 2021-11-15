VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private InventoryTickCount As Long


Private TempItem As Byte

Private ShowValue As Boolean
Private ImgContainer As Boolean

Private Inventory() As Inventory    'User's inventory
Private SlotsValidos As Byte


Private WithEvents InventoryWindow As PictureBox
Attribute InventoryWindow.VB_VarHelpID = -1
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
        InventoryTickCount = GetTickCount
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawInventory
    End If
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventory
End Sub

Public Sub ChangeSlotItemAmount(ByVal Slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    If Slot <= 0 Then Exit Sub
    Inventory(Slot).Amount = Amount
    
    'Update Amount
    Call DrawInventory
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    MaxHit = Inventory(Slot).MaxHit
End Property

Public Property Get MinHit(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    MinHit = Inventory(Slot).MinHit
End Property

Public Property Get MaxDef(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    MaxDef = Inventory(Slot).MaxDef
End Property

Public Property Get MinDef(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    MinDef = Inventory(Slot).MinDef
End Property

Public Property Get GrhIndex(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    GrhIndex = Inventory(Slot).GrhIndex
End Property

Public Property Get Valor(ByVal Slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    If Slot <= 0 Then Exit Property
    Valor = Inventory(Slot).Valor
End Property

Public Property Get Amount(ByVal Slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If Slot <= 0 Then Exit Property
    If Slot = FLAGORO Then
        Amount = UserGLD
    ElseIf Slot >= LBound(Inventory) And Slot <= UBound(Inventory) Then
        Amount = Inventory(Slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    If Slot <= 0 Then Exit Property
    OBJIndex = Inventory(Slot).OBJIndex
End Property

Public Property Get OBJType(ByVal Slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    If Slot <= 0 Then Exit Property
    OBJType = Inventory(Slot).OBJType
End Property

Public Property Get ItemName(ByVal Slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
    If Slot <= 0 Then Exit Property
'***************************************************
    ItemName = Inventory(Slot).name
End Property

Public Property Get Equipped(ByVal Slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    If Slot <= 0 Then Exit Property
    Equipped = Inventory(Slot).Equipped
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************

    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Property Let SlotValidos(ByVal Cant As Byte)
    SlotsValidos = Cant
End Property

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim Slot As Integer
    '   Update Inventory
    For Slot = 1 To Me.MaxObjs
        DrawInventory
    Next Slot
End Sub

Public Sub SetItem(ByVal Slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String, Optional ByVal PuedeUsar As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(Slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = PuedeUsar
    End With
    
    Call DrawInventory
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************

    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    'Make sure it's within limits
    If TempItem <= UBound(Inventory) Then
        'Make sure slot isn't empty
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
End Function

Private Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
    Dim LoopC As Long

    Dim DestRect As RECT

    With DestRect
        .bottom = InventoryWindow.Height
        .Right = InventoryWindow.Width
    End With

    
    'Clear the inventory window
    Engine_BeginScene
    
    Call TileEngine_Render_GrhIndex(24646, 0, 0, 0, ColorData.Blanco, 0, False)
    
        For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
            
            startdX = ((LoopC - 1) Mod (frmMain.PicInv.Width / 37)) * 37
            startdY = ((LoopC - 1) \ (frmMain.PicInv.Width / 37)) * 37
            
            
            If (LoopC <= SlotsValidos And LoopC <> SelectedItem) Then  ' SLOT NO SELECCIONADO
                
                If Inventory(LoopC).PuedeUsar = True Then
                    If Inventory(LoopC).GrhIndex Then Call TileEngine_Render_GrhIndex(Inventory(LoopC).GrhIndex, startdX + 2, startdY + 2, 0, ColorData.BlancoAB)
                    If Inventory(LoopC).Equipped Then Fonts_Render_String "E", startdX + 32, startdY + 24, ColorData.BlancoAB(1), 1, 2
                Else
                    If Inventory(LoopC).GrhIndex Then
                        Call TileEngine_Render_GrhIndex(Inventory(LoopC).GrhIndex, startdX + 2, startdY + 2, 0, ColorData.RojoAB)
                        Fonts_Render_String "X", startdX + 32, startdY + 24, ColorData.RojoAB(1), 1, 2
                    End If
                End If
                
            ElseIf LoopC = SelectedItem Then ' SLOT SELECCIONADO
            
                If Inventory(LoopC).PuedeUsar = True Then
                    If Inventory(LoopC).GrhIndex Then Call TileEngine_Render_GrhIndex(Inventory(LoopC).GrhIndex, startdX + 2, startdY + 2, 0, ColorData.Blanco)
                    If Inventory(LoopC).Equipped Then Fonts_Render_String "E", startdX + 32, startdY + 24, ColorData.Blanco(1), 1, 2
                    Fonts_Render_String Inventory(LoopC).Amount, startdX + 1, startdY, ColorData.Blanco(1), False, 2
                Else
                    If Inventory(LoopC).GrhIndex Then
                        Call TileEngine_Render_GrhIndex(Inventory(LoopC).GrhIndex, startdX + 2, startdY + 2, 0, ColorData.Rojo)
                        Fonts_Render_String "X", startdX + 32, startdY + 24, ColorData.Rojo(1), 1, 2
                    Fonts_Render_String Inventory(LoopC).Amount, startdX + 1, startdY, ColorData.Rojo(1), False, 2
                    End If
                End If



            ElseIf LoopC > SlotsValidos Then   'NO VALIDO
                If Inventory(LoopC).GrhIndex Then
                    Call TileEngine_Render_GrhIndex(Inventory(LoopC).GrhIndex, startdX + 2, startdY + 2, 0, ColorData.Rojo)
                End If
            End If
             

        Next LoopC
        
    Engine_EndScene DestRect, InventoryWindow.hWnd
End Sub

Public Sub DrawInv(Optional ByVal r As Boolean = False)
    If r = False Then SlotsValidos = 30

    If Settings.DinamicInventory Then
        If GetTickCount - InventoryTickCount > 1000 Then
            Call DrawInventory
            InventoryTickCount = GetTickCount
        End If
    Else
        DrawInventory
    End If
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 36, _
                            Optional ByVal TileHeight As Integer = 36, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True)

    Set InventoryWindow = InvPic

    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then
        If frmMain.UsandoDrag Then
            frmMain.UsabaDrag = True
            frmMain.UsandoDrag = False
        End If
       Exit Sub
    End If
    
    
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1

    'Standelf
    If TempItem <= UBound(Inventory) Then
        If GetKeyState(vbKeyTab) < 0 And Inventory(TempItem).name <> "" Then
            frmMain.invInfo.Visible = True
            
            frmMain.Label2(0).Caption = Inventory(TempItem).name
            frmMain.Label2(1).Caption = "Hit Min/Max: " & Inventory(TempItem).MinHit & "/" & Inventory(TempItem).MaxHit & vbNewLine & _
                                                    "Def Min/Max: " & Inventory(TempItem).MinDef & "/" & Inventory(TempItem).MaxDef & vbNewLine & _
                                                        "Cantidad: " & Inventory(TempItem).Amount
                                                        
            frmMain.invInfo.Left = (frmMain.PicInv.Left + (frmMain.PicInv.Width / 2)) - (frmMain.invInfo.Width / 2)
            frmMain.invInfo.Top = frmMain.PicInv.Top + Y + 20
        ElseIf frmMain.invInfo.Visible = True Then
            frmMain.invInfo.Visible = False
        End If
        
    End If
    
    If Not frmMain.UsabaDrag And Not frmMain.UsandoDrag Then
        frmMain.PicInv.MousePointer = vbArrow
    End If
    
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then
        If frmMain.UsandoDrag Then
            frmMain.UsabaDrag = True
            frmMain.UsandoDrag = False
        End If
        Exit Sub
    End If
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Or InvSelectedItem Then Call DrawInventory
    End If
    
    If frmMain.UsandoDrag And TempItem <= MaxObjs And TempItem >= 0 And Not frmMain.UsabaDrag Then
        WriteMoverInventario prevSelItem, TempItem
        frmMain.UsandoDrag = False
        frmMain.PicInv.MousePointer = vbArrow
    Else
        frmMain.UsandoDrag = False
        frmMain.PicInv.MousePointer = vbArrow
    End If
End Sub
