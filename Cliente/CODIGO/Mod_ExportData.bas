Attribute VB_Name = "Mod_ExportData"

Public Function CheckFile(ByVal FileT As String) As Boolean

    If General_File_Exist(FileT, vbNormal) Then
        CheckFile = True
    Else
        Dim FileC
        FileC = FreeFile
            Open FileT For Binary Access Write As FileC
            Close #File
    End If
    
End Function



Public Function Exportar_GRH()
    Dim i As Long, z As Long, tmpToSave As String
    CheckFile App.Path & "\Graficos.ini"

    General_Write_Var App.Path & "\Graficos.ini", "Init", "NumGRH", UBound(GrhData)
    For i = 1 To UBound(GrhData)
        
        'Animación
        If GrhData(i).NumFrames > 1 Then
        
            tmpToSave = GrhData(i).NumFrames
            For z = 1 To GrhData(i).NumFrames

                tmpToSave = tmpToSave & "-" & GrhData(i).Frames(z)
            
            Next z
            
            'Velocidad
            'tmpToSave = tmpToSave & ((GrhData(i).NumFrames * 1000) / 18)
        'GRH Simple
        Else
            'El GRH está vacio
            If GrhData(i).FileNum <= 0 Then
                tmpToSave = 0
            'Grh con contenido
            Else
                tmpToSave = "1-Others\" & GrhData(i).FileNum & "-" & GrhData(i).sX & "-" & GrhData(i).sY & "-" & GrhData(i).pixelWidth & "-" & GrhData(i).pixelHeight
            End If
        End If
    
        'Manager.ChangeValue "[GRH]", "GRH" & i, tmpToSave
        General_Write_Var App.Path & "\Graficos.ini", "GRH", "GRH" & i, tmpToSave
        DoEvents
    Next i
    
    MsgBox "Se exportó el Graficos.ini con éxito"
End Function
