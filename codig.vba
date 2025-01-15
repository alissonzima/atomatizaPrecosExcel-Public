Sub automatizarPrecificacao()

	Dim c As Range
	Dim firstAddress As String

	With Worksheets(1).Range("AP:AP")
		Set c = .Find("OFF GRID", LookIn:=xlValues)
		If Not c Is Nothing Then
			Do
				Rows(c.Row).Delete Shift:=xlUp
				Set c = .Find("OFF GRID", LookIn:=xlValues)
			Loop While Not c Is Nothing
		End If
	End With

	With Worksheets(1).Range("B:B")
		Set c = .Find("ALDO SOLAR ZERO GRID", LookIn:=xlValues)
		If Not c Is Nothing Then
			Do
				Rows(c.Row).Delete Shift:=xlUp
				Set c = .Find("ALDO SOLAR ZERO GRID", LookIn:=xlValues)
			Loop While Not c Is Nothing
		End If
	End With

	linhaFim = Range("AH1").End(xlDown).Row
	i = 2

	While i <= linhaFim
		If Cells(i, 34).Value <> "GROWATT" Then
			Cells(i, 34).EntireRow.Delete
			i = i - 1
			linhaFim = linhaFim - 1
		End If
		i = i + 1
	Wend

	linhaFim = Range("AR1").End(xlDown).Row
	i = 2

	While i <= linhaFim
		If Cells(i, 44).Value <> "PARAFUSO ESTRUTURAL MADEIRA" Then
			Cells(i, 44).EntireRow.Delete
			i = i - 1
			linhaFim = linhaFim - 1
		End If
		i = i + 1
	Wend

	Call deletaColuna(2, 1)
	Call deletaColuna(4, 2)
	Call deletaColuna(5, 18)
	Call deletaColuna(6, 7)
	Call deletaColuna(8, 7)
	Call deletaColuna(11, 71)
	Call deletaColuna(2, 2)
	Call deletaColuna(5, 1)
	
    With Worksheets(1).Range("G:G")
        Set c = .Find("FINAME/BNDES/MDA", LookIn:=xlValues)
        If Not c Is Nothing Then
            Do
                Rows(c.Row).Delete Shift:=xlUp
                Set c = .Find("FINAME/BNDES/MDA", LookIn:=xlValues)
            Loop While Not c Is Nothing
        End If
    End With		
	
	linhaFim = Range("A1").End(xlDown).Row
	
	Range("G2:G" & linhaFim).Copy
    Range("K1").PasteSpecial
    Application.CutCopyMode = False
    ActiveSheet.Range("$K$1:$K$" & linhaFim).RemoveDuplicates Columns:=1, Header:=xlNo
	
	linhaFim = Range("K1").End(xlDown).Row
    i = 1
    
    While i <= linhaFim
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = Sheets("Planilha1").Cells(i, 11)
        i = i + 1
    Wend
    
    Sheets("Planilha1").Range("K:K").Clear
	
    linha = 2
    
    While Sheets("Planilha1").Cells(linha, 1) <> ""
        Sheets("Planilha1").Range("A" & linha & ":G" & linha).Copy
        
        placa = Sheets("Planilha1").Cells(linha, 7)
        Sheets(placa).Select
        Range("A1000").End(xlUp).Offset(1, 0).PasteSpecial
        Application.CutCopyMode = False
        
        Sheets("Planilha1").Select
        
        linha = linha + 1
    Wend	
	
	Application.DisplayAlerts = False

    Sheets("Planilha1").Select
    ActiveWindow.SelectedSheets.Delete
	
    For Each aba In ThisWorkbook.Sheets
        linhaFim = aba.Range("A2").End(xlDown).Row
        aba.Range("A2:G" & linhaFim).Sort key1:=aba.Range("E2"), Order1:=xlAscending, key2:=aba.Range("B2"), Order2:=xlAscending
    Next	

    For Each aba In ThisWorkbook.Sheets
        i = 2
        linhaFim = aba.Range("A2").End(xlDown).Row
        If linhaFim > 1000 Then
            linhaFim = 2
        End If
        
        While i <= linhaFim
		
			aba.Range("B" & i).Value = Replace(aba.Range("B" & i), ".", "")
			aba.Range("B" & i).Value = CDbl(aba.Range("B" & i).Value)
            
            numAntesKW = InStr(aba.Range("C" & i), "KW ") - 5
            aba.Range("H" & i).Value = Mid(Mid(aba.Range("C" & i), numAntesKW, 5), InStr(Mid(aba.Range("C" & i), numAntesKW, 5), " "), 5)
            
            aba.Range("I" & i).Value = Mid(aba.Range("G" & i), InStr(aba.Range("G" & i), "W") - 3, 3)
			
			aba.Range("J" & i).Value = (aba.Range("E" & i) * 1000) / aba.Range("I" & i)
			
            aba.Range("K" & i).Value = Application.WorksheetFunction.RoundUp(((((aba.Range("H" & i) * 1000) * 1.4) / aba.Range("I" & i)) - aba.Range("J" & i)), 0)			
                                 
			If aba.Range("E" & i) = aba.Range("E" & i + 1) Then

                aba.Range("L" & i).Value = "M"
                
            Else
            
                If aba.Range("E" & i) = aba.Range("E" & i - 1) Then
                
                    aba.Range("L" & i).Value = "M"
                
                Else
                
                    aba.Range("L" & i).Value = "U"
                
                End If
            End If
			
            i = i + 1
            
        Wend
    Next
	
    For Each aba In ThisWorkbook.Sheets
        linhaFim = aba.Range("A2").End(xlDown).Row
        aba.Range("A2:L" & linhaFim).Sort key1:=aba.Range("E2"), Order1:=xlAscending, key2:=aba.Range("B2"), Order2:=xlAscending
    Next

    For Each aba In ThisWorkbook.Sheets
        i = 2
        linhaFim = aba.Range("A2").End(xlDown).Row
        If linhaFim > 1000 Then
            linhaFim = 2
        End If
        
        While i <= linhaFim
            
            If aba.Range("L" & i) = "U" Then
                aba.Range("M" & i).Value = "Sim"
            Else
                If aba.Range("E" & i) <> aba.Range("E" & i - 1) Then
                    If aba.Range("K" & i) <> 1 Then
                        aba.Range("M" & i).Value = "Sim"
                    Else
                        If aba.Range("E" & i) = aba.Range("E" & i + 1) Then
                            If aba.Range("E" & i + 1) <> 1 Then
                                aba.Range("M" & i).Value = "Nao"
                            End If
                        End If
                    End If
                Else
                    If aba.Range("M" & i - 1) = "Sim" Or aba.Range("M" & i - 1) = "Jatem" Then
                        aba.Range("M" & i).Value = "Jatem"
                    Else
                        If aba.Range("M" & i - 1) = "Quemsabe" And aba.Range("K" & i) <> 1 Then
                            aba.Range("M" & i) = "Sim"
                        End If
                        If aba.Range("K" & i) <> 1 Then
                        aba.Range("M" & i).Value = "Sim"
                        Else
                            If aba.Range("E" & i) = aba.Range("E" & i + 1) And aba.Range("K" & i + 1) <> 1 Then
                                aba.Range("M" & i).Value = "Nao"
                            Else
                                If aba.Range("M" & i - 1) = "Quemsabe" Then
                                    aba.Range("M" & i) = "Nao"
                                Else
                                    aba.Range("M" & i - 1).Value = "Talvez"
                                    aba.Range("M" & i).Value = "Quemsabe"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
                                                                  
            i = i + 1
        
        Wend
    Next
	
    For Each aba In ThisWorkbook.Sheets
        i = 2
        linhaFim = aba.Range("A2").End(xlDown).Row
        If linhaFim > 1000 Then
            linhaFim = 2
        End If
        
        While i <= linhaFim
            
            If aba.Range("M" & i).Value = "Jatem" Or aba.Range("M" & i).Value = "Nao" Then
                aba.Cells(i, 1).EntireRow.Delete
                i = i - 1
                linhaFim = linhaFim - 1
            End If
            
            i = i + 1
        
        Wend
    Next	
	
    For Each aba In ThisWorkbook.Sheets
        i = 2
        linhaFim = aba.Range("A2").End(xlDown).Row
        If linhaFim > 1000 Then
            linhaFim = 2
        End If
        
        While i <= linhaFim
            
            If aba.Range("M" & i).Value = "Quemsabe" Then
                If aba.Range("E" & i) = aba.Range("E" & i + 1) And aba.Range("M" & i + 1).Value = "Sim" Then
                    aba.Cells(i - 1, 1).EntireRow.Delete
                    aba.Cells(i - 1, 1).EntireRow.Delete
                Else
                    aba.Range("M" & i - 1).Value = "Sim"
                    aba.Cells(i, 1).EntireRow.Delete
                End If
                i = i - 1
                linhaFim = linhaFim - 1
            End If
            
            i = i + 1
        
        Wend
    Next
	
    For Each aba In ThisWorkbook.Sheets
    
        aba.Range("N1").Value = "KWP INICIAL"
        aba.Range("O1").Value = "KWP FINAL"
		aba.Range("P1").Value = "INVERSOR"
		aba.Range("Q1").Value = "CUSTO KIT"
		aba.Range("R1").Value = "CUSTO ART"
		aba.Range("S1").Value = "CUSTO ENGENHEIRO"
		aba.Range("T1").Value = "GANHO COMERCIAL"
		aba.Range("U1").Value = "GANHO INSTALACAO"
		aba.Range("V1").Value = "GANHO INSTALADOR"
		aba.Range("W1").Value = "ADICIONAL DESP+PROJ"
		aba.Range("X1").Value = "VALOR VENDA KIT"
		aba.Range("Y1").Value = "APP ->"
		aba.Range("Z1").Value = "$ W GERAL"
		aba.Range("AA1").Value = "$ SERVICO"
		aba.Range("AB1").Value = "LAJE ->"
		aba.Range("AC1").Value = "$ LAJE"
		aba.Range("AD1").Value = "SERV LAJE"
		aba.Range("AE1").Value = "SOLO ->"
		aba.Range("AF1").Value = "$ SOLO"
		aba.Range("AG1").Value = "SERV SOLO"
		aba.Range("AH1").Value = "SEM EST ->"
		aba.Range("AI1").Value = "$ SEM EST"
		aba.Range("AJ1").Value = "SERV SEM EST"
    
        i = 2
        linhaFim = aba.Range("A2").End(xlDown).Row
        If linhaFim > 1000 Then
            linhaFim = 2
        End If
        
        While i <= linhaFim
            
            aba.Range("N" & i).Value = aba.Range("E" & i).Value
            aba.Range("O" & i).Value = aba.Range("E" & i).Value
			aba.Range("P" & i).Value = "GROWATT " & Replace(aba.Range("H" & i).Value, ",", ".")
			aba.Range("Q" & i).Value = aba.Range("B" & i).Value
			aba.Range("R" & i).Value = CDbl("150")
			If aba.Range("O" & i) <= 25 Then
                aba.Range("S" & i).Value = CDbl("300")
            Else
                If aba.Range("O" & i) <= 50 Then
                    aba.Range("S" & i).Value = CDbl("500")
                Else
                    If aba.Range("O" & i) <= 75 Then
                        aba.Range("S" & i).Value = CDbl("700")
                    Else
                        aba.Range("S" & i).Value = CDbl("1000")
                    End If
                End If
            End If
			
            If aba.Range("O" & i) <= CDbl("10,5") Then
                    aba.Range("T" & i).Value = CDbl("0,45")
                Else
                    aba.Range("T" & i).Value = CDbl("0,35")
            End If
			
            If aba.Range("O" & i) <= CDbl("4,49") Then
                aba.Range("U" & i).Value = aba.Range("N" & i).Value * 1000 * CDbl("0,3")
            Else
                If aba.Range("O" & i) <= CDbl("10,35") Then
                    aba.Range("U" & i).Value = aba.Range("N" & i).Value * 1000 * CDbl("0,2")
                Else
                    aba.Range("U" & i).Value = aba.Range("N" & i).Value * 1000 * CDbl("0,15")
                End If
            End If			
			
			If aba.Range("O" & i) <= CDbl("4,49") Then
					aba.Range("V" & i).Value = aba.Range("N" & i).Value * 1000 * CDbl("0,3")
				Else
					aba.Range("V" & i).Value = aba.Range("N" & i).Value * 1000 * CDbl("0,15")
			End If		

			aba.Range("W" & i).Value = CDbl("1115")
			
            valor = aba.Range("U" & i).Value + aba.Range("V" & i).Value + aba.Range("S" & i).Value + aba.Range("R" & i).Value + aba.Range("Q" & i).Value + aba.Range("W" & i).Value + (aba.Range("T" & i).Value * aba.Range("Q" & i).Value)
            aba.Range("X" & i).Value = Round(valor + (valor * CDbl("0,06")), 2)
			
			aba.Range("Y" & i).Value = "APP ->"
			aba.Range("Z" & i).Value = Round(aba.Range("Q" & i).Value / aba.Range("O" & i).Value / 1000, 2)
			aba.Range("AA" & i).Value = Round(aba.Range("X" & i).Value - aba.Range("Q" & i).Value, 2)
			aba.Range("AB" & i).Value = "LAJE ->"
			aba.Range("AC" & i).Value = aba.Range("Z" & i).Value + CDbl("0,15")
			
            novoCusto = aba.Range("AC" & i).Value * aba.Range("O" & i).Value * 1000
            valor = aba.Range("U" & i).Value + aba.Range("V" & i).Value + aba.Range("S" & i).Value + aba.Range("R" & i).Value + novoCusto + aba.Range("W" & i).Value + (aba.Range("T" & i).Value * novoCusto)
            valorVenda = Round(valor + (valor * CDbl("0,06")), 2)
            aba.Range("AD" & i).Value = valorVenda - novoCusto
			
			aba.Range("AE" & i).Value = "SOLO ->"
			aba.Range("AF" & i).Value = aba.Range("Z" & i).Value + CDbl("0,31")
			
            novoCusto = aba.Range("AF" & i).Value * aba.Range("O" & i).Value * 1000
            valor = aba.Range("U" & i).Value + aba.Range("V" & i).Value + aba.Range("S" & i).Value + aba.Range("R" & i).Value + novoCusto + aba.Range("W" & i).Value + (aba.Range("T" & i).Value * novoCusto)
            valorVenda = Round(valor + (valor * CDbl("0,06")), 2)
            aba.Range("AG" & i).Value = valorVenda - novoCusto			
			
			aba.Range("AH" & i).Value = "SEM EST ->"
			aba.Range("AI" & i).Value = aba.Range("Z" & i).Value - CDbl("0,15")
			
            novoCusto = aba.Range("AI" & i).Value * aba.Range("O" & i).Value * 1000
            valor = aba.Range("U" & i).Value + aba.Range("V" & i).Value + aba.Range("S" & i).Value + aba.Range("R" & i).Value + novoCusto + aba.Range("W" & i).Value + (aba.Range("T" & i).Value * novoCusto)
            valorVenda = Round(valor + (valor * CDbl("0,06")), 2)
            aba.Range("AJ" & i).Value = valorVenda - novoCusto

            i = i + 1
            
        Wend
    Next	

    For Each aba In ThisWorkbook.Sheets
    
    aba.Columns("N:AJ").AutoFit

    Next
	
End Sub

Sub deletaColuna(x, y)
i = 0
While i < y
	Columns(x).EntireColumn.Delete
	i = i + 1
Wend

End Sub

