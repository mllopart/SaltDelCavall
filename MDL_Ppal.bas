Attribute VB_Name = "MDL_Ppal"
Option Explicit
'--------------------------CONSTANTS---------------------------------------------

Public Const tamany_x = 10 'Tamany x de la matriu
Public Const tamany_y = 10 'Tamany y de la matriu


'--------------------------VARIABLES---------------------------------------------

Public m(1 To tamany_x, 1 To tamany_y) As Integer 'matriu principal que contindra
                                   'totes les possibilitats
                                   
                                                     
Public inici_x As Integer 'Marca la posició x d'inici
Public inici_y As Integer 'Marca la posició y d'inici

Public mov As Integer 'Indica el número de moviments que fem

Public Type posi
    p_x As Integer
    p_y As Integer
    numero As Integer
    final As Boolean
End Type

Public segon As Boolean
    



'-------------------------FUNCIONS i PROCEDIMENTS-------------------------------------


Public Function Buscar_Millor_Posicio(ByVal x, ByVal y) As posi

    Dim i As Integer
    Dim j As Integer
    Dim millor_moviment As Integer
    Dim posicio_millor As posi
    Dim millor_moviment2 As Integer
    
    millor_moviment = 8
    millor_moviment2 = 8
    posicio_millor.final = False
    
           'Primera possibilitat (dos dreta, un avall)
            If y + 2 <= tamany_y Then
                If x + 1 <= tamany_x Then
                   If m(x + 1, y + 2) <= millor_moviment And m(x + 1, y + 2) <> 0 Then
                            millor_moviment = m(x + 1, y + 2)
                            posicio_millor.p_x = x + 1
                            posicio_millor.p_y = y + 2
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Segona possibilitat (dos esquerra, un avall)
            If y - 2 >= 1 Then
                If x + 1 <= tamany_x Then
                    If m(x + 1, y - 2) <= millor_moviment And m(x + 1, y - 2) <> 0 Then
                            millor_moviment = m(x + 1, y - 2)
                            posicio_millor.p_x = x + 1
                            posicio_millor.p_y = y - 2
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Tercera possibilitat (dos dreta, un amunt)
            If y + 2 <= tamany_y Then
                If x - 1 >= 1 Then
                    If m(x - 1, y + 2) <= millor_moviment And m(x - 1, y + 2) <> 0 Then
                            millor_moviment = m(x - 1, y + 2)
                            posicio_millor.p_x = x - 1
                            posicio_millor.p_y = y + 2
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Cuarta possibilitat (dos esquerra, un amunt)
            If y - 2 >= 1 Then
                If x - 1 >= 1 Then
                    If m(x - 1, y - 2) <= millor_moviment And m(x - 1, y - 2) <> 0 Then
                            millor_moviment = m(x - 1, y - 2)
                            posicio_millor.p_x = x - 1
                            posicio_millor.p_y = y - 2
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Quinta possibilitat (un esquerra, dos amunt)
            If y - 1 >= 1 Then
                If x - 2 >= 1 Then
                    If m(x - 2, y - 1) <= millor_moviment And m(x - 2, y - 1) <> 0 Then
                            millor_moviment = m(x - 2, y - 1)
                            posicio_millor.p_x = x - 2
                            posicio_millor.p_y = y - 1
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Sexta possibilitat (un dreta, dos amunt)
            If y + 1 <= tamany_y Then
                If x - 2 >= 1 Then
                    If m(x - 2, y + 1) <= millor_moviment And m(x - 2, y + 1) <> 0 Then
                            millor_moviment = m(x - 2, y + 1)
                            posicio_millor.p_x = x - 2
                            posicio_millor.p_y = y + 1
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Setena possibilitat (un dreta, dos avall)
            If y + 1 <= tamany_y Then
                If x + 2 <= tamany_x Then
                    If m(x + 2, y + 1) <= millor_moviment And m(x + 2, y + 1) <> 0 Then
                            millor_moviment = m(x + 2, y + 1)
                            posicio_millor.p_x = x + 2
                            posicio_millor.p_y = y + 1
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
            
            'Vuitena possibilitat (un esquerra, dos avall)
            If y - 1 >= 1 Then
                If x + 2 <= tamany_x Then
                    If m(x + 2, y - 1) <= millor_moviment And m(x + 2, y - 1) <> 0 Then
                            millor_moviment = m(x + 2, y - 1)
                            posicio_millor.p_x = x + 2
                            posicio_millor.p_y = y - 1
                            posicio_millor.numero = millor_moviment
                    End If
                End If
            End If
    
    If posicio_millor.p_x = 0 And posicio_millor.p_y = 0 Then
        posicio_millor.final = True
    End If
    
    Buscar_Millor_Posicio = posicio_millor
    
End Function

Public Function Triar_Posicio_Inici() As posi
'Busca la posició inicial en que començara el cavall i la guarda en una variable
    
    Dim x As Integer
    Dim y As Integer
    Dim posici As posi
    
    x = 0
    y = 0
    
    While x = 0
        Randomize
        x = Int((tamany_x * Rnd) + 1)
    Wend
    
    While y = 0
        Randomize
        y = Int((tamany_y * Rnd) + 1)
    Wend
    
    
    posici.p_x = x
    posici.p_y = y
    
    posici.numero = m(x, y)
    
    Triar_Posicio_Inici = posici
    
    
End Function
Public Sub Buscar_Posicions_Possibles()
'Omple la matriu amb les possibilitats que te cada casella de la matriu
'Sempre pensant amb el moviment del cavall

    Dim x As Integer
    Dim y As Integer
    Dim pos As Integer

    
    
    For x = 1 To tamany_x
        For y = 1 To tamany_y
            
            pos = 0
            
            'Primera possibilitat (dos dreta, un avall)
            If y + 2 <= tamany_y Then
                If x + 1 <= tamany_x Then
                    If (m(x + 1, y + 2) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Segona possibilitat (dos esquerra, un avall)
            If y - 2 >= 1 Then
                If x + 1 <= tamany_x Then
                    If (m(x + 1, y - 2) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Tercera possibilitat (dos dreta, un amunt)
            If y + 2 <= tamany_y Then
                If x - 1 >= 1 Then
                    If (m(x - 1, y + 2) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Cuarta possibilitat (dos esquerra, un amunt)
            If y - 2 >= 1 Then
                If x - 1 >= 1 Then
                    If (m(x - 1, y - 2) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Quinta possibilitat (un esquerra, dos amunt)
            If y - 1 >= 1 Then
                If x - 2 >= 1 Then
                    If (m(x - 2, y - 1) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Sexta possibilitat (un dreta, dos amunt)
            If y + 1 <= tamany_y Then
                If x - 2 >= 1 Then
                    If (m(x - 2, y + 1) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Setena possibilitat (un dreta, dos avall)
            If y + 1 <= tamany_y Then
                If x + 2 <= tamany_x Then
                    If (m(x + 2, y + 1) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            'Vuitena possibilitat (un esquerra, dos avall)
            If y - 1 >= 1 Then
                If x + 2 <= tamany_x Then
                    If (m(x + 2, y - 1) <> 0) Or (segon = False) Then
                        pos = pos + 1
                    End If
                End If
            End If
            
            If m(x, y) <> 0 Or segon = False Then
                m(x, y) = pos
            End If
                   
        Next
    Next
    
    

End Sub
Public Sub Omplir_Matriu_Textos()
'Omple totes les caselles de texte amb el seu corresponent a la matriu
'original

   
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    k = 1
    For i = 1 To tamany_x
        For j = 1 To tamany_y
          FRM_Cavall.Text1(k).Text = m(i, j)
          k = k + 1
        Next
    Next
    FRM_Cavall.L_Moviments.Caption = mov
End Sub
Public Sub Omplir_matriu_picture()
Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    k = 0
    For i = 1 To tamany_y
        For j = 1 To tamany_y
            If m(i, j) = 0 Then FRM_Cavall.Picture1(k).BackColor = &HFF&
            k = k + 1
        Next
    Next
End Sub

Public Sub Reiniciar_pictures()
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 99
        FRM_Cavall.Picture1(i).BackColor = &H8000000A
    Next
    

End Sub
Public Sub Posar_Matriu_a_0()
'Posem la matriu principal a 0

    Dim i As Integer
    Dim j As Integer

    For i = 1 To tamany_x
        For j = 1 To tamany_y
            m(i, j) = 0
        Next
    Next
    
End Sub
