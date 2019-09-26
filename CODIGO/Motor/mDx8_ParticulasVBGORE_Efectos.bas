Attribute VB_Name = "mDx8_ParticulasVBGORE_Efectos"
Option Explicit

'Constants With The Order Number For Each Effect
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect

'esto lo declaro aca por que falta parece solo usarse en este modulo
Private WeatherEffectIndex As Integer

Public Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long
    
    With Effect(EffectIndex)
        
        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function

        'Return the index of the used slot
        Effect_Snow_Begin = EffectIndex

        'Set The Effect's Variables
        .EffectNum = EffectNum_Snow      'Set the effect number
        .ParticleCount = Particles       'Set the number of particles
        .Used = True                     'Enabled the effect
        .Gfx = Gfx                       'Set the graphic

        'Set the number of particles left to the total avaliable
        .ParticlesLeft = .ParticleCount

        'Set the float variables
        .FloatSize = Effect_FToDW(10)    'Size of the particles

        'Redim the number of particles
        ReDim .Particles(0 To .ParticleCount)
        ReDim .PartVertex(0 To .ParticleCount)

        'Create the particles
        For LoopC = 0 To .ParticleCount
            Set .Particles(LoopC) = New Particle
            .Particles(LoopC).Used = True
            .PartVertex(LoopC).Rhw = 1
            Call Effect_Snow_Reset(EffectIndex, LoopC, 1)
        Next LoopC

        'Set the initial time
        .PreviousFrame = timeGetTime
    
    End With

End Function

Public Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, _
                             ByVal Index As Long, _
                             Optional ByVal FirstReset As Byte = 0)
    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
    '*****************************************************************
    
    With Effect(EffectIndex)
    
        If FirstReset = 1 Then

            'The very first reset
            Call .Particles(Index).ResetIt(-200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0)

        Else

            'Any reset after first
            Call .Particles(Index).ResetIt(-200 + (Rnd * (frmMain.ScaleWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0)

            If .Particles(Index).sngX < -20 Then .Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
            If .Particles(Index).sngX > frmMain.ScaleWidth Then .Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
            If .Particles(Index).sngY > frmMain.ScaleHeight Then .Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

        End If

        'Set the color
        Call .Particles(Index).ResetColor(1, 1, 1, 0.8, 0)

    End With
    
End Sub

Public Sub Effect_Snow_Update(ByVal EffectIndex As Integer)

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long
    
    With Effect(EffectIndex)
    
        'Calculate the time difference
        ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
        .PreviousFrame = timeGetTime

        'Go through the particle loop
        For LoopC = 0 To .ParticleCount

            'Check if particle is in use
            If .Particles(LoopC).Used Then

                'Update The Particle
                Call .Particles(LoopC).UpdateParticle(ElapsedTime)

                'Check if to reset the particle
                If .Particles(LoopC).sngX < -200 Then .Particles(LoopC).sngA = 0
                If .Particles(LoopC).sngX > (frmMain.ScaleWidth + 200) Then .Particles(LoopC).sngA = 0
                If .Particles(LoopC).sngY > (frmMain.ScaleHeight + 200) Then .Particles(LoopC).sngA = 0

                'Time for a reset, baby!
                If .Particles(LoopC).sngA <= 0 Then

                    'Reset the particle
                    Call Effect_Snow_Reset(EffectIndex, LoopC)

                Else

                    'Set the particle information on the particle vertex
                    .PartVertex(LoopC).Color = D3DColorMake(.Particles(LoopC).sngR, .Particles(LoopC).sngG, .Particles(LoopC).sngB, .Particles(LoopC).sngA)
                    .PartVertex(LoopC).x = .Particles(LoopC).sngX
                    .PartVertex(LoopC).y = .Particles(LoopC).sngY

                End If

            End If

        Next LoopC

    End With
    
End Sub

Public Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long
    
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
    
    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex
        
    With Effect(EffectIndex)

        'Set the effect's variables
        .EffectNum = EffectNum_Rain      'Set the effect number
        .ParticleCount = Particles       'Set the number of particles
        .Used = True                     'Enabled the effect
        .Gfx = Gfx                       'Set the graphic

        'Set the number of particles left to the total avaliable
        .ParticlesLeft = .ParticleCount

        'Set the float variables
        .FloatSize = Effect_FToDW(10)    'Size of the particles

        'Redim the number of particles
        ReDim .Particles(0 To .ParticleCount)
        ReDim .PartVertex(0 To .ParticleCount)

        'Create the particles
        For LoopC = 0 To .ParticleCount
            Set .Particles(LoopC) = New Particle
            .Particles(LoopC).Used = True
            .PartVertex(LoopC).Rhw = 1
            Call Effect_Rain_Reset(EffectIndex, LoopC, 1)
        Next LoopC

        'Set The Initial Time
        .PreviousFrame = timeGetTime
    
    End With
    
End Function

Public Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, _
                              ByVal Index As Long, _
                              Optional ByVal FirstReset As Byte = 0)

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
    '*****************************************************************
    With Effect(EffectIndex)

        If FirstReset = 1 Then
        
            'The very first reset
            Call .Particles(Index).ResetIt(-200 + (Rnd * (frmMain.ScaleWidth + 400)), Rnd * (frmMain.ScaleHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0)

        Else

            'Any reset after first
            Call .Particles(Index).ResetIt(-200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0)

            If .Particles(Index).sngX < -20 Then .Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
            If .Particles(Index).sngX > frmMain.ScaleWidth Then .Particles(Index).sngY = Rnd * (frmMain.ScaleHeight + 50)
            If .Particles(Index).sngY > frmMain.ScaleHeight Then .Particles(Index).sngX = Rnd * (frmMain.ScaleWidth + 50)

        End If

        'Set the color
        Call .Particles(Index).ResetColor(1, 1, 1, 0.4, 0)

    End With
    
End Sub

Public Sub Effect_Rain_Update(ByVal EffectIndex As Integer)

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long
    
    With Effect(EffectIndex)
    
        'Calculate the time difference
        ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
        .PreviousFrame = timeGetTime

        'Go through the particle loop
        For LoopC = 0 To .ParticleCount

            'Check if the particle is in use
            If .Particles(LoopC).Used Then

                'Update the particle
                Call .Particles(LoopC).UpdateParticle(ElapsedTime)

                'Check if to reset the particle
                If .Particles(LoopC).sngX < -200 Then .Particles(LoopC).sngA = 0
                If .Particles(LoopC).sngX > (frmMain.ScaleWidth + 200) Then .Particles(LoopC).sngA = 0
                If .Particles(LoopC).sngY > (frmMain.ScaleHeight + 200) Then .Particles(LoopC).sngA = 0

                'Time for a reset, baby!
                If .Particles(LoopC).sngA <= 0 Then

                    'Reset the particle
                    Call Effect_Rain_Reset(EffectIndex, LoopC)

                Else

                    'Set the particle information on the particle vertex
                    .PartVertex(LoopC).Color = D3DColorMake(.Particles(LoopC).sngR, .Particles(LoopC).sngG, .Particles(LoopC).sngB, .Particles(LoopC).sngA)
                    .PartVertex(LoopC).x = .Particles(LoopC).sngX
                    .PartVertex(LoopC).y = .Particles(LoopC).sngY

                End If

            End If

        Next LoopC
    
    End With
    
End Sub

Public Function Effect_Summon_Begin(ByVal x As Single, _
                                    ByVal y As Single, _
                                    ByVal Gfx As Integer, _
                                    ByVal Particles As Integer, _
                                    Optional ByVal Progression As Single = 0) As Integer

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
    '*****************************************************************
    Dim EffectIndex As Integer
    Dim LoopC       As Long

    With Effect(EffectIndex)

        'Get the next open effect slot
        EffectIndex = Effect_NextOpenSlot

        If EffectIndex = -1 Then Exit Function

        'Return the index of the used slot
        Effect_Summon_Begin = EffectIndex

        'Set The Effect's Variables
        .EffectNum = EffectNum_Summon    'Set the effect number
        .ParticleCount = Particles       'Set the number of particles
        .Used = True                     'Enable the effect
        .x = x                           'Set the effect's X coordinate
        .y = y                           'Set the effect's Y coordinate
        .Gfx = Gfx                       'Set the graphic
        .Progression = Progression       'If we loop the effect

        'Set the number of particles left to the total avaliable
        .ParticlesLeft = .ParticleCount

        'Set the float variables
        .FloatSize = Effect_FToDW(8)    'Size of the particles

        'Redim the number of particles
        ReDim .Particles(0 To .ParticleCount)
        ReDim .PartVertex(0 To .ParticleCount)

        'Create the particles
        For LoopC = 0 To .ParticleCount
            Set .Particles(LoopC) = New Particle
            .Particles(LoopC).Used = True
            .PartVertex(LoopC).Rhw = 1
            Call Effect_Summon_Reset(EffectIndex, LoopC)
        Next LoopC

        'Set The Initial Time
        .PreviousFrame = timeGetTime
    
    End With
    
End Function

Public Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
    '*****************************************************************
    Dim x As Single
    Dim y As Single
    Dim r As Single
    
    With Effect(EffectIndex)
    
        If .Progression > 1000 Then
            .Progression = .Progression + 1.4
        Else
            .Progression = .Progression + 0.5
        End If

        r = (Index / 30) * Exp(Index / .Progression)
        x = r * Cos(Index)
        y = r * Sin(Index)
    
        'Reset the particle
        Call .Particles(Index).ResetIt(.x + x, .y + y, 0, 0, 0, 0)
        Call .Particles(Index).ResetColor(0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2))
    
    End With
    
End Sub

Public Sub Effect_Summon_Update(ByVal EffectIndex As Integer)

    '*****************************************************************
    'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
    '*****************************************************************
    Dim ElapsedTime As Single
    Dim LoopC       As Long
    
    With Effect(EffectIndex)
    
        'Calculate The Time Difference
        ElapsedTime = (timeGetTime - .PreviousFrame) * 0.01
        .PreviousFrame = timeGetTime

        'Go Through The Particle Loop
        For LoopC = 0 To .ParticleCount

            'Check If Particle Is In Use
            If .Particles(LoopC).Used Then

                'Update The Particle
                Call .Particles(LoopC).UpdateParticle(ElapsedTime)

                'Check if the particle is ready to die
                If .Particles(LoopC).sngA <= 0 Then

                    'Check if the effect is ending
                    If .Progression < 1800 Then

                        'Reset the particle
                        Call Effect_Summon_Reset(EffectIndex, LoopC)

                    Else

                        'Disable the particle
                        .Particles(LoopC).Used = False

                        'Subtract from the total particle count
                        .ParticlesLeft = .ParticlesLeft - 1

                        'Check if the effect is out of particles
                        If .ParticlesLeft = 0 Then .Used = False

                        'Clear the color (dont leave behind any artifacts)
                        .PartVertex(LoopC).Color = 0

                    End If

                Else
            
                    'Set the particle information on the particle vertex
                    .PartVertex(LoopC).Color = D3DColorMake(.Particles(LoopC).sngR, .Particles(LoopC).sngG, .Particles(LoopC).sngB, .Particles(LoopC).sngA)
                    .PartVertex(LoopC).x = .Particles(LoopC).sngX
                    .PartVertex(LoopC).y = .Particles(LoopC).sngY

                End If

            End If

        Next LoopC
    
    End With
    
End Sub

Public Sub Engine_Weather_Update()

    If bRain And bLluvia(UserMap) = 1 And CurMapAmbient.Rain = True Then
    
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)
            
        ElseIf Effect(WeatherEffectIndex).EffectNum <> eParticulas.Rain Then
            Call Effect_Kill(WeatherEffectIndex)
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)
            
        ElseIf Not Effect(WeatherEffectIndex).Used Then
            WeatherEffectIndex = Effect_Rain_Begin(9, 500)

        End If

    End If

    If CurMapAmbient.Snow = True Then
    
        If WeatherEffectIndex <= 0 Then
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)
            
        ElseIf Effect(WeatherEffectIndex).EffectNum <> eParticulas.Rain Then
            Call Effect_Kill(WeatherEffectIndex)
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)
            
        ElseIf Not Effect(WeatherEffectIndex).Used Then
            WeatherEffectIndex = Effect_Snow_Begin(14, 200)

        End If

    End If
            
    If CurMapAmbient.Fog <> -1 Then
        Call Engine_Weather_UpdateFog
    End If
    
   'esto lo saco hay que ver para que se usaba en vbgore
   ' If OnRampageImgGrh <> 0 Then
   '     Call Draw_GrhIndex(OnRampageImgGrh, 0, 0, 0, Normal_RGBList(), 0, True)
   ' End If
    
End Sub

Public Sub Engine_Weather_UpdateFog()
    '*****************************************************************
    'Update the fog effects
    '*****************************************************************

    Dim i           As Long
    Dim x           As Long
    Dim y           As Long
    Dim CC(3)       As Long
    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop

    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop

    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop

    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop

    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop

    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop

    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop
    
    'Render fog 2
    x = 2
    y = -1
    
    With CurMapAmbient
    
        CC(1) = D3DColorARGB(.Fog, 255, 255, 255)
        CC(2) = D3DColorARGB(.Fog, 255, 255, 255)
        CC(3) = D3DColorARGB(.Fog, 255, 255, 255)
        CC(0) = D3DColorARGB(.Fog, 255, 255, 255)

        For i = 1 To WeatherFogCount
            Call Draw_GrhIndex(27300, (x * 512) + WeatherFogX2, (y * 512) + WeatherFogY2, 0, CC(), 0, False)
            x = x + 1

            If x > (1 + (ScreenWidth \ 512)) Then
                x = 0
                y = y + 1

            End If

        Next i
            
        'Render fog 1
        x = 0
        y = 0
        CC(1) = D3DColorARGB(.Fog / 2, 255, 255, 255)
        CC(2) = D3DColorARGB(.Fog / 2, 255, 255, 255)
        CC(3) = D3DColorARGB(.Fog / 2, 255, 255, 255)
        CC(0) = D3DColorARGB(.Fog / 2, 255, 255, 255)
    
    End With

    For i = 1 To WeatherFogCount
        
        Call Draw_GrhIndex(27301, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 0, CC(), 0, False)
        
        x = x + 1

        If x > (2 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i

End Sub

Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
'*****************************************************************
Dim EffectIndex As Integer

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000

End Function

Private Function Effect_FToDW(F As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData Buf, 0, 4, 1, F
    DirectD3D8.BufferGetData Buf, 0, 4, 1, Effect_FToDW

End Function
