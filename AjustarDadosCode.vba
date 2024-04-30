Function regradetres(dout, F1, F2)
    Dim superior, inferior As Double
    superior = dout * F2
    inferior = F1
    regradetres = superior / inferior
End Function
'---------------------------------'
Function TrianGular(F, F2, F3, dout2, dout3)
    Dim y1, x1, y2 As Double
    y1 = F3 - F2
    x1 = dout3 - dout2
    y2 = F - F2
    TrianGular = ((y2 * x1) / y1) + dout2
End Function

'---------------------------------'
Sub Button1_Click()
    Dim Ft(1 To 40000), Fd(1 To 40000), t(1 To 40000), d(1 To 40000) As Double
    Dim Ftpico(1 To 10), Fdpico(1 To 10), teste As Double
    Dim vardo1, vardo2, vardo3, vardo4, vardo5, vardopicoFt, vardopicoFd As Integer
    Dim ladoquevai, linhas, soma1, igual As Integer
    Dim negativosoma1, negativovardo2, negativo, vezesquepassou As Integer
    Dim linhapico As Integer
    linhapico = 5
    vezesquepassou = 0
    vardo1 = 2
    vardo2 = 2
    vardo3 = 4
    vardo5 = 1
    vardopicoFd = 1
    vardopicoFt = 1
    ladoquevai = 0
    linhas = 1
    soma1 = 2
    igual = 1
    Do Until Cells(vardo1, 1) = Empty
        Ft(vardo1) = Cells(vardo1, 1)
        Ft(vardo1) = CDbl(Format(Ft(vardo1), "00.0000000"))
        Fd(vardo1) = Cells(vardo1, 4)
        Fd(vardo1) = CDbl(Format(Fd(vardo1), "00.0000000"))
        t(vardo1) = Cells(vardo1, 2)
        t(vardo1) = CDbl(Format(t(vardo1), "00000.0000000"))
        d(vardo1) = Cells(vardo1, 5)
        d(vardo1) = CDbl(Format(d(vardo1), "000.0000000"))
        vardo1 = vardo1 + 1
        If Cells(vardo1, 1) = Empty Then
            vardo4 = vardo1
            Do Until Cells(vardo4, 4) = Empty
                Fd(vardo4) = Cells(vardo4, 4)
                d(vardo4) = Cells(vardo4, 5)
                vardo4 = vardo4 + 1
            Loop
        End If
    Loop
    'ACHAR PICOS
    Do Until vardo3 = vardo4
        'PICO POSITIVO Ft
        If Ft(vardo3 - 3) = Ft(vardo3 - 1) And Ft(vardo3) > Ft(vardo3 + 1) Then
            If Ft(vardo3) = Ft(vardo3 - 1) Then
            Else
                Ftpico(vardopicoFt) = Ft(vardo3)
                vardopicoFt = vardopicoFt + 1
            End If
        End If
        'PICO NEGATIVO Ft
        If Ft(vardo3) < Ft(vardo3 - 1) And Ft(vardo3) < Ft(vardo3 + 1) Then
            If vardo3 < 4000 Then
                Ftpico(vardopicoFt) = Ft(vardo3)
                vardopicoFt = vardopicoFt + 1
            End If
        End If
        If Ft(vardo3 - 3) = Ft(vardo3 - 1) And Ft(vardo3) < Ft(vardo3 + 1) Then
            If Ft(vardo3) = Ft(vardo3 - 1) Then
            Else
                Ftpico(vardopicoFt) = Ft(vardo3)
                vardopicoFt = vardopicoFt + 1
            End If
        End If
        'PICO POSITIVO Fd
        If Fd(vardo3) > Fd(vardo3 - 1) And Fd(vardo3) = Fd(vardo3 + 1) Then
            Fdpico(vardopicoFd) = Fd(vardo3)
            vardopicoFd = vardopicoFd + 1
        End If
        'PICO NEGATIVO Fd
        If Fd(vardo3) < Fd(vardo3 - 1) And Fd(vardo3) = Fd(vardo3 + 1) Then
            Fdpico(vardopicoFd) = Fd(vardo3)
            vardopicoFd = vardopicoFd + 1
        End If
        vardo3 = vardo3 + 1
    Loop
    'ESCREVER PICOS NA PLANILHA
    Do Until vardo5 = vardopicoFd
        Cells(linhapico, 11) = Fdpico(vardo5)
        Cells(linhapico, 11).Interior.ColorIndex = 3
        Cells(linhapico, 13) = Ftpico(vardo5)
        Cells(linhapico, 13).Interior.ColorIndex = 3
        vardo5 = vardo5 + 1
        linhapico = linhapico + 1
    Loop
    vardopicoFd = 1
    vardopicoFt = 1
    'ORGANIZAR
    Do Until vardo2 = vardo4
        If Ft(vardo2) > Fd(soma1) Then
        'QUANDO ENCONTRAR PICO
                If Fd(soma1) = Fdpico(vardopicoFd) Or Ft(vardo2) = Ftpico(vardopicoFt) Then
                    If Fd(soma1) = Fdpico(vardopicoFd) Then
                        If vezesquepassou = 1 Then
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                'Cells(linhas, 7).Interior.ColorIndex = 3
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Ft(vardo2) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        Else
                            Cells(linhas, 7) = Ft(vardo2)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        End If
                    Else
                        If vezesquepassou = 1 Then
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1) * -1
                                'Cells(linhas, 7).Interior.ColorIndex = 3
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Fd(soma1) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        Else
                            Cells(linhas, 7) = Fd(soma1)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        End If
                    End If
            vardopicoFt = vardopicoFt + 1
            vardopicoFd = vardopicoFd + 1
            'TRANSFORMAR O CONTRÁRIO
                
                negativosoma1 = soma1
                negativovardo2 = vardo2
                negativo = vardopicoFd
                
                Do Until negativosoma1 = vardo4
                    Fd(negativosoma1) = Fd(negativosoma1) * (-1)
                    Do Until Fdpico(negativo) = Empty
                        Fdpico(negativo) = Fdpico(negativo) * (-1)
                        negativo = negativo + 1
                    Loop
                    negativosoma1 = negativosoma1 + 1
                Loop
                negativo = vardopicoFt
                Do Until negativovardo2 = vardo1
                    Ft(negativovardo2) = Ft(negativovardo2) * (-1)
                    Do Until Ftpico(negativo) = Empty
                        Ftpico(negativo) = Ftpico(negativo) * (-1)
                        negativo = negativo + 1
                    Loop
                    negativovardo2 = negativovardo2 + 1
                Loop
                If vezesquepassou = 0 Then
                    vezesquepassou = 1
                Else
                    vezesquepassou = 0
                End If
        End If
        'CASO ACABE UMA DAS COLUNAS
            If Fd(soma1) = Empty Or Ft(vardo2) = Empty Then
                    If Fd(soma1) = Empty Then
                        If vezesquepassou = 1 Then
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        Exit Sub
                    Else
                        If vezesquepassou = 1 Then
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        Exit Sub
                    End If
            End If
            'COLOCA VALOR MENOR PRIMEIRO
            If vezesquepassou = 1 Then
                Cells(linhas, 7) = Fd(soma1) * -1
                Cells(linhas, 8) = d(soma1)
                linhas = linhas + 1
                soma1 = soma1 + 1
            Else
                Cells(linhas, 7) = Fd(soma1)
                Cells(linhas, 8) = d(soma1)
                linhas = linhas + 1
                soma1 = soma1 + 1
            End If
            'SE CONTINUA MENOR COLOCA
            Do Until Ft(vardo2) < Fd(soma1)
                'QUANDO ENCONTRAR PICO
                If Fd(soma1) = Fdpico(vardopicoFd) Or Ft(vardo2) = Ftpico(vardopicoFt) Then
                    If Fd(soma1) = Fdpico(vardopicoFd) Then
                        If vezesquepassou = 1 Then
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Ft(vardo2) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        Else
                            Cells(linhas, 7) = Ft(vardo2)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        End If
                    Else
                        If vezesquepassou = 1 Then
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Fd(soma1) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        Else
                            Cells(linhas, 7) = Fd(soma1)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        End If
                    End If
                    vardopicoFt = vardopicoFt + 1
                    vardopicoFd = vardopicoFd + 1
                    'TRANSFORMAR O CONTRÁRIO
                
                        negativosoma1 = soma1
                        negativovardo2 = vardo2
                        negativo = vardopicoFd
                        
                        Do Until negativosoma1 = vardo4
                            Fd(negativosoma1) = Fd(negativosoma1) * (-1)
                            Do Until Fdpico(negativo) = Empty
                                Fdpico(negativo) = Fdpico(negativo) * (-1)
                                negativo = negativo + 1
                            Loop
                            negativosoma1 = negativosoma1 + 1
                        Loop
                        negativo = vardopicoFt
                        Do Until negativovardo2 >= vardo1
                            Ft(negativovardo2) = Ft(negativovardo2) * (-1)
                            Do Until Ftpico(negativo) = Empty
                                Ftpico(negativo) = Ftpico(negativo) * (-1)
                                negativo = negativo + 1
                            Loop
                            negativovardo2 = negativovardo2 + 1
                        Loop
                        If vezesquepassou = 0 Then
                            vezesquepassou = 1
                        Else
                            vezesquepassou = 0
                        End If
                End If
                If Fd(soma1) = Empty Or Ft(vardo2) = Empty Then
                    If Fd(soma1) = Empty Then
                        If Ft(vardo2) = Empty Then
                            Exit Sub
                        End If
                        If vezesquepassou = 1 Then
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        Exit Sub
                    Else
                        If vezesquepassou = 1 Then
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        Exit Sub
                    End If
                End If
                If Ft(vardo2) = Fd(soma1) Then
                    If vezesquepassou = 1 Then
                        Cells(linhas, 7) = Ft(vardo2) * -1
                        Cells(linhas, 8) = d(soma1)
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                        vardo2 = vardo2 + 1
                        igual = 0
                        Exit Do
                    Else
                        Cells(linhas, 7) = Ft(vardo2)
                        Cells(linhas, 8) = d(soma1)
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                        vardo2 = vardo2 + 1
                        igual = 0
                        Exit Do
                    End If
                End If
                'COLOCA O QUE SOBROU
                If vezesquepassou = 1 Then
                    Cells(linhas, 7) = Fd(soma1) * -1
                    Cells(linhas, 8) = d(soma1)
                    linhas = linhas + 1
                    soma1 = soma1 + 1
                Else
                    Cells(linhas, 7) = Fd(soma1)
                    Cells(linhas, 8) = d(soma1)
                    linhas = linhas + 1
                    soma1 = soma1 + 1
                End If
            Loop
            If igual = 1 Then
                If vezesquepassou = 1 Then
                    Cells(linhas, 7) = Ft(vardo2) * -1
                    Cells(linhas, 9) = t(vardo2)
                    linhas = linhas + 1
                    vardo2 = vardo2 + 1
                Else
                    Cells(linhas, 7) = Ft(vardo2)
                    Cells(linhas, 9) = t(vardo2)
                    linhas = linhas + 1
                    vardo2 = vardo2 + 1
                End If
            End If
            igual = 1
        Else
            'QUANDO ENCONTRAR PICO
            If Fd(soma1) = Fdpico(vardopicoFd) Or Ft(vardo2) = Ftpico(vardopicoFt) Then
                If Fd(soma1) = Fdpico(vardopicoFd) Then
                    If vezesquepassou = 1 Then
                        Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                            If Ft(vardo2) = Empty Then
                                Exit Do
                            End If
                            Cells(linhas, 7) = Ft(vardo2) * -1
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        Loop
                    Else
                        Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                            If Ft(vardo2) = Empty Then
                                Exit Do
                            End If
                            Cells(linhas, 7) = Ft(vardo2)
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        Loop
                    End If
                    If vezesquepassou = 1 Then
                        Cells(linhas, 7) = Ft(vardo2) * -1
                        Cells(linhas, 7).Interior.ColorIndex = 3
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        vardo2 = vardo2 + 1
                    Else
                        Cells(linhas, 7) = Ft(vardo2)
                        Cells(linhas, 7).Interior.ColorIndex = 3
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        vardo2 = vardo2 + 1
                    End If
                Else
                    If vezesquepassou = 1 Then
                        Do Until Fd(soma1) = Fdpico(vardopicoFd)
                            If Ft(vardo2) = Empty Then
                                Exit Do
                            End If
                            Cells(linhas, 7) = Fd(soma1) * -1
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        Loop
                    Else
                        Do Until Fd(soma1) = Fdpico(vardopicoFd)
                            If Ft(vardo2) = Empty Then
                                Exit Do
                            End If
                            Cells(linhas, 7) = Fd(soma1)
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        Loop
                    End If
                    If vezesquepassou = 1 Then
                        Cells(linhas, 7) = Fd(soma1) * -1
                        Cells(linhas, 7).Interior.ColorIndex = 3
                        Cells(linhas, 8) = d(soma1)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                    Else
                        Cells(linhas, 7) = Fd(soma1)
                        Cells(linhas, 7).Interior.ColorIndex = 3
                        Cells(linhas, 8) = d(soma1)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                    End If
                End If
                vardopicoFt = vardopicoFt + 1
                vardopicoFd = vardopicoFd + 1
                
                'TRANSFORMAR O CONTRÁRIO
                
                negativosoma1 = soma1
                negativovardo2 = vardo2
                negativo = vardopicoFd
                
                Do Until negativosoma1 = vardo4
                    Fd(negativosoma1) = Fd(negativosoma1) * (-1)
                    Do Until Fdpico(negativo) = Empty
                        Fdpico(negativo) = Fdpico(negativo) * (-1)
                        negativo = negativo + 1
                    Loop
                    negativosoma1 = negativosoma1 + 1
                Loop
                negativo = vardopicoFt
                Do Until negativovardo2 = vardo1
                    Ft(negativovardo2) = Ft(negativovardo2) * (-1)
                    Do Until Ftpico(negativo) = Empty
                        Ftpico(negativo) = Ftpico(negativo) * (-1)
                        negativo = negativo + 1
                    Loop
                    negativovardo2 = negativovardo2 + 1
                Loop
                If vezesquepassou = 0 Then
                    vezesquepassou = 1
                Else
                    vezesquepassou = 0
                End If
                
            End If
            If Fd(soma1) = Empty Or Ft(vardo2) = Empty Then
                    If Fd(soma1) = Empty Then
                        If vezesquepassou = 1 Then
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        Exit Sub
                    Else
                        If vezesquepassou = 1 Then
                             Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        Exit Sub
                    End If
            End If
            If vezesquepassou = 1 Then
                Cells(linhas, 7) = Ft(vardo2) * -1
                Cells(linhas, 9) = t(vardo2)
                linhas = linhas + 1
                vardo2 = vardo2 + 1
            Else
                Cells(linhas, 7) = Ft(vardo2)
                Cells(linhas, 9) = t(vardo2)
                linhas = linhas + 1
                vardo2 = vardo2 + 1
            End If
            'QUANDO CONTINUA
            Do Until Ft(vardo2) > Fd(soma1)
                'QUANDO ENCONTRAR PICO
                If Fd(soma1) = Fdpico(vardopicoFd) Or Ft(vardo2) = Ftpico(vardopicoFt) Then
                    If Fd(soma1) = Fdpico(vardopicoFd) Then
                        If vezesquepassou = 1 Then
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until Ft(vardo2) = Ftpico(vardopicoFt)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Ft(vardo2) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        Else
                            Cells(linhas, 7) = Ft(vardo2)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 9) = t(vardo2)
                            linhas = linhas + 1
                            vardo2 = vardo2 + 1
                        End If
                    Else
                        If vezesquepassou = 1 Then
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until Fd(soma1) = Fdpico(vardopicoFd)
                                If Ft(vardo2) = Empty Then
                                    Exit Do
                                End If
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        If vezesquepassou = 1 Then
                            Cells(linhas, 7) = Fd(soma1) * -1
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        Else
                            Cells(linhas, 7) = Fd(soma1)
                            Cells(linhas, 7).Interior.ColorIndex = 3
                            Cells(linhas, 8) = d(soma1)
                            linhas = linhas + 1
                            soma1 = soma1 + 1
                        End If
                    End If
                    
                    vardopicoFt = vardopicoFt + 1
                    vardopicoFd = vardopicoFd + 1
                
                        'TRANSFORMAR O CONTRÁRIO
                        
                        negativosoma1 = soma1
                        negativovardo2 = vardo2
                        negativo = vardopicoFd
                        
                        Do Until negativosoma1 >= vardo4
                            Fd(negativosoma1) = Fd(negativosoma1) * (-1)
                            Do Until Fdpico(negativo) = Empty
                                Fdpico(negativo) = Fdpico(negativo) * (-1)
                                negativo = negativo + 1
                            Loop
                            negativosoma1 = negativosoma1 + 1
                        Loop
                        negativo = vardopicoFt
                        Do Until negativovardo2 = vardo1
                            Ft(negativovardo2) = Ft(negativovardo2) * (-1)
                            Do Until Ftpico(negativo) = Empty
                                Ftpico(negativo) = Ftpico(negativo) * (-1)
                                negativo = negativo + 1
                            Loop
                            negativovardo2 = negativovardo2 + 1
                        Loop
                        If vezesquepassou = 0 Then
                            vezesquepassou = 1
                        Else
                            vezesquepassou = 0
                        End If
                End If
                If Fd(soma1) = Empty Or Ft(vardo2) = Empty Then
                    If Fd(soma1) = Empty Then
                        If vezesquepassou = 1 Then
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2) * -1
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        Else
                            Do Until vardo2 = vardo1
                                Cells(linhas, 7) = Ft(vardo2)
                                Cells(linhas, 9) = t(vardo2)
                                linhas = linhas + 1
                                vardo2 = vardo2 + 1
                            Loop
                        End If
                        Exit Sub
                    Else
                        If vezesquepassou = 1 Then
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1) * -1
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        Else
                            Do Until soma1 = vardo4
                                Cells(linhas, 7) = Fd(soma1)
                                Cells(linhas, 8) = d(soma1)
                                linhas = linhas + 1
                                soma1 = soma1 + 1
                            Loop
                        End If
                        Exit Sub
                    End If
                End If
                If Ft(vardo2) = Fd(soma1) Then
                    If vezesquepassou = 1 Then
                        Cells(linhas, 7) = Ft(vardo2) * -1
                        Cells(linhas, 8) = d(soma1)
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                        vardo2 = vardo2 + 1
                        igual = 0
                        Exit Do
                    Else
                        Cells(linhas, 7) = Ft(vardo2)
                        Cells(linhas, 8) = d(soma1)
                        Cells(linhas, 9) = t(vardo2)
                        linhas = linhas + 1
                        soma1 = soma1 + 1
                        vardo2 = vardo2 + 1
                        igual = 0
                        Exit Do
                    End If
                End If
                If vezesquepassou = 1 Then
                    Cells(linhas, 7) = Ft(vardo2) * -1
                    Cells(linhas, 9) = t(vardo2)
                    linhas = linhas + 1
                    vardo2 = vardo2 + 1
                Else
                    Cells(linhas, 7) = Ft(vardo2)
                    Cells(linhas, 9) = t(vardo2)
                    linhas = linhas + 1
                    vardo2 = vardo2 + 1
                End If
            Loop
            If igual = 1 Then
                If vezesquepassou = 1 Then
                    Cells(linhas, 7) = Fd(soma1) * (-1)
                    Cells(linhas, 8) = d(soma1)
                    linhas = linhas + 1
                    soma1 = soma1 + 1
                Else
                    Cells(linhas, 7) = Fd(soma1)
                    Cells(linhas, 8) = d(soma1)
                    linhas = linhas + 1
                    soma1 = soma1 + 1
                End If
            End If
            igual = 1
        End If
  
    Loop
End Sub

'---------------------------------'
Sub prenncher_click()
    Dim vardo1, vardo2, col As Integer
    Dim linhafim, linha, linhasub As Single
    Dim F, F2, F3 As Double
    Dim dout, dout2, dout3 As Double
    col = 8
    linha = 1
    vardo1 = 1
    vardo2 = 1
    linhafim = 70453
    'MUDA DE COLUNA
    Do Until col >= 10
        'FAZ AS LINHAS
        Do Until linha >= linhafim
            'Primeira linha
            If linha = 1 Then
                If col = 9 Then
                Else
                    F = Cells(linha, 7)
                    F2 = Cells(linha + 1, 7)
                    dout2 = Cells(linha + 1, col)
                    dout = regradetres(dout2, F2, F)
                    Cells(linha, col) = dout
                    Cells(linha, col).Interior.ColorIndex = 3
                End If
            End If
            'ULTIMA LINHA
            If linha = 70082 Then
                If col = 8 Then
                    Do Until linha = linhafim
                        F = Cells(linha, 7)
                        F2 = Cells(linha - 1, 7)
                        dout2 = Cells(linha - 1, col)
                        dout = regradetres(dout, F2, F)
                        Cells(linha, col) = dout
                        Cells(linha, col).Interior.ColorIndex = 3
                        linha = linha + 1
                    Loop
                Else
                    Exit Sub
                End If
            End If
            'TODAS AS OUTRAS
            If Cells(linha, col) = Empty Then
                F = Cells(linha, 7)
                F2 = Cells(linha - 1, 7)
                dout2 = Cells(linha - 1, col)
                linhasub = linha
                Do Until Cells(linhasub, col) <> Empty
                    If linhasub = 70453 Then
                        Exit Do
                    End If
                    linhasub = linhasub + 1
                Loop
                F3 = Cells(linhasub, 7)
                dout3 = Cells(linhasub, col)
                dout = TrianGular(F, F2, F3, dout2, dout3)
                Cells(linha, col) = dout
                Cells(linha, col).Interior.ColorIndex = 3
            End If
            linha = linha + 1
        Loop
        linha = 1
        col = col + 1
    Loop
End Sub

'---------------------------------'
Private Sub buttonsearch_Click()
    If TboxValor = Empty Then
        MsgBox ("Inserir Valor!")
        Exit Sub
    Else
        If TboxStart = Empty Then
            MsgBox ("Inserir Local de inicio da pesquisa!")
            Exit Sub
        Else
            Dim valor, ValorCelula As Double
            Dim start, EndCelula As String
            Dim VarFor1 As Integer
            Dim ls, cs As Single
            valor = CDbl(TboxValor.Value)
            start = TboxStart.Value
            For VarFor1 = Len(start) To 1 Step -1
                If Mid(start, VarFor1, 1) = "," Then
                    ls = CSng(Left(start, VarFor1 - 1))
                    cs = CSng(Right(start, Len(start) - VarFor1))
                End If
            Next VarFor1
            Do Until ls = 50000
                ValorCelula = CDbl(Cells(ls, cs).Value)
                If ValorCelula = valor Then
                    EndCelula = Cells(ls, cs).AddressLocal
                    MsgBox (EndCelula)
                    Cells(ls, cs).Activate
                    Exit Do
                End If
                ls = ls + 1
                'helper de search
                If ls = 16457 Then
                    ls = 16457
                End If
                'Caso não encontre o valor
                If ls = 50000 Then
                    MsgBox ("Valor não encontrado!")
                End If
                
            Loop
        End If
    End If
End Sub

Private Sub UserForm_Click()

End Sub

'---------------------------------'
