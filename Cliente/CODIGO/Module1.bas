Attribute VB_Name = "Module1"
Option Explicit

Public Function testersss(ByVal Xi As Integer, Yi As Integer)
Dim xz As Long

If (Xi Or Yi) = 0 Then Exit Function

    ' The center
        For xz = 0 To 3
            MapData(Xi, Yi).Vertex_Offset(xz) = MapData(Xi, Yi).Vertex_Offset(xz) + 50
        Next xz
        
        
    'MiddleTop
        MapData(Xi, Yi - 1).Vertex_Offset(2) = MapData(Xi, Yi).Vertex_Offset(0)
        MapData(Xi, Yi - 1).Vertex_Offset(3) = MapData(Xi, Yi).Vertex_Offset(1)
        
        MapData(Xi, Yi - 1).Vertex_Offset(0) = MapData(Xi, Yi - 2).Vertex_Offset(2)
        MapData(Xi, Yi - 1).Vertex_Offset(1) = MapData(Xi, Yi - 2).Vertex_Offset(3)
        
    'MiddleBottom
        MapData(Xi, Yi + 1).Vertex_Offset(0) = MapData(Xi, Yi - 1).Vertex_Offset(2)
        MapData(Xi, Yi + 1).Vertex_Offset(1) = MapData(Xi, Yi - 1).Vertex_Offset(3)
        
        MapData(Xi, Yi + 1).Vertex_Offset(2) = MapData(Xi, Yi + 2).Vertex_Offset(0)
        MapData(Xi, Yi + 1).Vertex_Offset(3) = MapData(Xi, Yi + 2).Vertex_Offset(1)
        
    'Middle Left
        MapData(Xi - 1, Yi).Vertex_Offset(1) = MapData(Xi, Yi).Vertex_Offset(0)
        MapData(Xi - 1, Yi).Vertex_Offset(3) = MapData(Xi, Yi).Vertex_Offset(2)
        
        MapData(Xi - 1, Yi).Vertex_Offset(0) = MapData(Xi - 2, Yi).Vertex_Offset(1)
        MapData(Xi - 1, Yi).Vertex_Offset(2) = MapData(Xi - 2, Yi).Vertex_Offset(3)
    
    'Middle Right
        MapData(Xi + 1, Yi).Vertex_Offset(0) = MapData(Xi, Yi).Vertex_Offset(1)
        MapData(Xi + 1, Yi).Vertex_Offset(2) = MapData(Xi, Yi).Vertex_Offset(3)
        
        MapData(Xi + 1, Yi).Vertex_Offset(1) = MapData(Xi + 2, Yi).Vertex_Offset(0)
        MapData(Xi + 1, Yi).Vertex_Offset(3) = MapData(Xi + 2, Yi).Vertex_Offset(2)
    
    
    
    ' TOP LEFT CORNER
        MapData(Xi - 1, Yi - 1).Vertex_Offset(0) = MapData(Xi - 1, Yi - 2).Vertex_Offset(2)
        MapData(Xi - 1, Yi - 1).Vertex_Offset(1) = MapData(Xi - 1, Yi - 2).Vertex_Offset(3)
        
        MapData(Xi - 1, Yi - 1).Vertex_Offset(2) = MapData(Xi - 1, Yi + 1).Vertex_Offset(0)
        MapData(Xi - 1, Yi - 1).Vertex_Offset(3) = MapData(Xi - 1, Yi + 1).Vertex_Offset(1)
                
    ' BOTTOM LEFT CORNER
        MapData(Xi - 1, Yi + 1).Vertex_Offset(0) = MapData(Xi - 1, Yi).Vertex_Offset(2)
        MapData(Xi - 1, Yi + 1).Vertex_Offset(1) = MapData(Xi - 1, Yi).Vertex_Offset(3)
        
        MapData(Xi - 1, Yi + 1).Vertex_Offset(2) = MapData(Xi - 1, Yi + 2).Vertex_Offset(0)
        MapData(Xi - 1, Yi + 1).Vertex_Offset(3) = MapData(Xi - 1, Yi + 2).Vertex_Offset(1)
            
            
            
            
    ' TOP RIGHT CORNER
        MapData(Xi + 1, Yi - 1).Vertex_Offset(0) = MapData(Xi, Yi - 1).Vertex_Offset(1)
        MapData(Xi + 1, Yi - 1).Vertex_Offset(2) = MapData(Xi, Yi - 1).Vertex_Offset(3)
        
        MapData(Xi + 1, Yi - 1).Vertex_Offset(1) = MapData(Xi + 2, Yi + 1).Vertex_Offset(0)
        MapData(Xi + 1, Yi - 1).Vertex_Offset(3) = MapData(Xi + 2, Yi + 1).Vertex_Offset(2)
                
    ' BOTTOM RIGHT CORNER
        MapData(Xi + 1, Yi + 1).Vertex_Offset(0) = MapData(Xi, Yi + 1).Vertex_Offset(1)
        MapData(Xi + 1, Yi + 1).Vertex_Offset(2) = MapData(Xi, Yi + 1).Vertex_Offset(3)
        
        MapData(Xi + 1, Yi + 1).Vertex_Offset(1) = MapData(Xi + 2, Yi + 1).Vertex_Offset(0)
        MapData(Xi + 1, Yi + 1).Vertex_Offset(3) = MapData(Xi + 2, Yi + 1).Vertex_Offset(1)
                 
            
End Function

