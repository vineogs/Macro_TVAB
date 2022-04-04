Attribute VB_Name = "Módulo1"
Sub Macro()

    'Formatação dos materiais da rede
    For GCR = 1 To 3000
        If Cells(GCR, 7).Value = "GCR" Then
            Rows(GCR).Font.Color = RGB(0, 176, 80)
        End If
    Next GCR
    
    'Formatação dos títulos
    
    Range("A1", "A9").EntireRow.Delete
    
    Dim tamanhoTitulo As Integer
    
    For Titulo = 1 To 3000
        tamanhoTitulo = Titulo + 19
        If Cells(Titulo, 8).Value = "LOC" Then
            Range("H" & Titulo, "H" & tamanhoTitulo).EntireRow.Interior.Color = RGB(0, 112, 192)
            Range("H" & Titulo, "H" & tamanhoTitulo).EntireRow.Font.Color = RGB(255, 255, 255)
        End If
    Next Titulo

    For Titulo2 = 1 To 3000
        If Cells(Titulo2, 15).Value = "PROGRAMA ATÉ 1 INTERVALO" Then
            If Cells((Titulo2 - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Titulo2 - 24), "O" & (Titulo2 - 3)).EntireRow.Delete
            End If
        End If
        
        If Cells(Titulo2, 15).Value = "PROGRAMA ATÉ 2 INTERVALO" Then
            If Cells((Titulo2 - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Titulo2 - 24), "O" & (Titulo2 - 3)).EntireRow.Delete
            End If
        End If
           
        If Cells(Titulo2, 15).Value = "PROGRAMA ATÉ 3 INTERVALO" Then
            If Cells((Titulo2 - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Titulo2 - 24), "O" & (Titulo2 - 3)).EntireRow.Delete
            End If
        End If
        
        If Cells(Titulo2, 15).Value = "PROGRAMA ATÉ 4 INTERVALO" Then
            If Cells((Titulo2 - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Titulo2 - 24), "O" & (Titulo2 - 3)).EntireRow.Delete
            End If
        End If
        
        If Cells(Titulo2, 15).Value = "PROGRAMA ATÉ 5 INTERVALO" Then
            If Cells((Titulo2 - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Titulo2 - 24), "O" & (Titulo2 - 3)).EntireRow.Delete
            End If
        End If
    Next Titulo2
    
    For Continuacao = 1 To 3000
        If Cells(Continuacao, 15).Value = "PROGRAMA ATÉ 1 INTERVALO" And Cells(Continuacao, 33).Value = "(CONTINUAÇÃO)" Then
            Range("F" & (Continuacao - 3), "F" & (Continuacao + 3)).EntireRow.Delete
        End If
    
        If Cells(Continuacao, 15).Value = "PROGRAMA ATÉ 2 INTERVALO" And Cells(Continuacao, 33).Value = "(CONTINUAÇÃO)" Then
            Range("F" & (Continuacao - 3), "F" & (Continuacao + 3)).EntireRow.Delete
        End If
        
        If Cells(Continuacao, 15).Value = "PROGRAMA ATÉ 3 INTERVALO" And Cells(Continuacao, 33).Value = "(CONTINUAÇÃO)" Then
            Range("F" & (Continuacao - 3), "F" & (Continuacao + 3)).EntireRow.Delete
        End If
        
        If Cells(Continuacao, 15).Value = "PROGRAMA ATÉ 4 INTERVALO" And Cells(Continuacao, 33).Value = "(CONTINUAÇÃO)" Then
            Range("F" & (Continuacao - 3), "F" & (Continuacao + 3)).EntireRow.Delete
        End If
    Next Continuacao
    
    For Enc = 1 To 3000
        If Cells(Enc, 15).Value = "PROGRAMA ATÉ ENCERRAMENTO" Then
            If Cells((Enc - 10), "Y").Value = "TÍTULO" Then
                Range("O" & (Enc - 24), "O" & (Enc - 3)).EntireRow.Delete
            End If
            Range("O" & (Enc - 3), "O" & (Enc - 2)).EntireRow.Delete
            Range("O" & (Enc), "O" & (Enc + 2)).EntireRow.Delete
        End If
   Next Enc
   
   For Enc2 = 1 To 3000
        If Cells(Enc2, "O").Value = "PROGRAMA ATÉ ENCERRAMENTO" And Cells(Enc2 + 6, "AB").Value = "CENTRAL DE DISTRIBUIÇÃO - ROTEIRO COMERCIAL" Then
            Range("O" & Enc2 - 2, "O" & Enc2 - 1).EntireRow.Delete
            Range("O" & Enc2 + 1, "O" & Enc2 + 3).EntireRow.Delete
        End If
   Next Enc2
    'Formatação do Calhau Canal
    
    Dim tempoCalhau As Integer
    tempoCalhau = 5
    For tempo = 1 To 60
        For Calhau = 1 To 3000
            If Cells(Calhau, 26).Value = "CALHAU CANAL " & tempoCalhau Then
                Rows(Calhau).Font.Color = RGB(255, 0, 0)
            End If
        Next Calhau
        tempoCalhau = tempoCalhau + 5
    Next tempo
    
    'Formatação de Fade (sem contar o tempo)
    
    For Fade = 1 To 3000
        For intervalo = 1 To 4
            
            If Cells(Fade, "O").Value = "PROGRAMA ATÉ " & intervalo & " INTERVALO" Then
                For comercial = Fade To (Fade + 5)
                    If Cells(comercial + 4, "Z").Font.Color = RGB(0, 0, 0) Then
                        Rows(comercial + 4).Copy
                        Rows(comercial + 4).Insert
                        Rows(comercial + 4).PasteSpecial
                        Rows(comercial + 4).RowHeight = 40
                        Rows(comercial + 4).ClearContents
                        Cells((comercial + 4), "Z").Value = "FADE "
                        Rows(comercial + 4).Font.Color = RGB(255, 0, 0)
                        Cells(comercial + 4, "Z").HorizontalAlignment = xlRight
                        Exit For
                    End If
                Next comercial
            End If
        Next intervalo
    Next Fade
    
    'Formatação das Chamadas da rede
    For chamada = 1 To 3000
        For intervalo2 = 1 To 5
            If Cells(chamada, "O").Value = "PROGRAMA ATÉ " & intervalo2 & " INTERVALO" Then
                Rows(chamada + 4).Copy
                Rows(chamada + 4).Insert
                Rows(chamada + 4).PasteSpecial
                Rows(chamada + 4).ClearContents
                Rows(chamada + 4).Font.Color = RGB(0, 176, 80)
                Cells(chamada + 4, "G").Value = "GCR"
                Cells(chamada + 4, "AI").Value = "CH"
            End If
        Next intervalo2
    Next chamada
End Sub

