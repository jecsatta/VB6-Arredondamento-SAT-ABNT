'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'FUNÇAO PARA FAZER O ARREDONDAMENTO DE VALORES, BASEADO NAS REGRAS DE ARREDONDAMENTO DA NORMA ABNT NBR 5891 DE 1977
'TRABALHA COM 4 DIGITOS NA DECIMAL DE ENTRADA
'DEVOLVERÁ O VALOR ARREDONDADO COM 2 DECIMAIS
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function Arredondamento_ABNT_NBR5891(Valor As Currency) As String
    On Error GoTo Trata_Erros
  
  
   'TRANSFORMA E FORMATA O VALOR PARA STRING E 4 DECIMAIS
   Dim StrValor_Trabalhar As String
   StrValor_Trabalhar = Format(Valor, "############0.0000")

   'DESCOBRE A POSIÇAO DA VIRGULA
   Dim Posicao_Virgula As Integer
   Posicao_Virgula = InStr(1, CStr(StrValor_Trabalhar), ",")
   Dim StrDecimal As String
   StrDecimal = Mid(StrValor_Trabalhar, Posicao_Virgula + 1, Len(StrValor_Trabalhar))
  
   'VERIFICA SE NA DECIMAL OS 2 ULTIMOS DIGITOS SAO IGUAIS A "00", SE FOREM, NAO SERÁ NECESSÁRIO ARREDONDAR
   'POR EXEMPLO 2,5500
   If Mid(StrDecimal, 3, 2) = "00" Then
      Arredondamento_ABNT_NBR5891 = Format(CCur(StrValor_Trabalhar), "############0.00")
      Exit Function
   End If
  
  
   'DEFAULT
   Dim StrValor_Retornar As String
   StrValor_Retornar = CStr(Format(Valor, "#############0.00"))
  
  
   '********************************************************************************************************************************************
   '1- Quando o algarismo seguinte a 2S CASA for INFERIOR a 5, A 2S CASA permanecerá SEM modificaçao.
   'ENTAO SE NA 3S CASA O NUMERO FOR < 5 (MENOR QUE 5) ENTAO NAO ARREDONDA, MANTEM O VALOR ORIGINAL
   'EXEMPLO 2,5501 FICARÁ SOMENTE 2,55 POIS A TERCEIRA CASA (0) É MENOR QUE 5
   '********************************************************************************************************************************************
   If CInt(Mid(StrDecimal, 3, 1)) < 5 Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5501  REMOVERÁ O 01 DO FINAL, RETORNANDO SOMENTE O 2,55
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
      
   End If
  
   '********************************************************************************************************************************************
   '2 - Quando o algarismo seguinte A 2S CASA for SUPERIOR a 5 ENTAO AUMENTARA EM UMA UNIDADE A 2S CASA, EXEMPLO: 2,556 (FICA 2,56)
   '********************************************************************************************************************************************
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   'VERIFICA SE A TERCEIRA CASA É MAIOR QUE 5
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If CInt(Mid(StrDecimal, 3, 1)) > 5 Then
      'SE FOR MAIOR QUE 5, ENTAO ARREDONDA PRA MAIS O VALOR, EXEMPLO: 2,556 FICARÁ 2,56
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5501  REMOVERÁ O 01 DO FINAL, RETORNANDO SOMENTE O 2,55
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
  
   '************************************************************************************************************************************************************************
   '3 - Quando a TERCEIRA CASA É IGUAL A CINCO, TEREMOS 2 OPCOES (A e B):
   '************************************************************************************************************************************************************************
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '(A) - SE A SEGUNDA CASA FOR IMPAR ENTAO ARREDONDA PRA MAIS O VALOR, EXEMPLO: 2,3751 (o 7 dos 37 centavos é IMPAR, neste caso arredonda pra mais)
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   If EImpar(CLng(Mid(StrDecimal, 2, 1))) = True Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,3751  REMOVERÁ O 51 DO FINAL, RETORNANDO SOMENTE O 2,37
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '(B) - SE A SEGUNDA CASA FOR PAR, ENTAO:
   'SE A QUARTA CASA FOR ALGARISMO ZERO, NAO HAVERÁ ALTERAÇAO NAS DECIMAIS, RETORNANDO O VALOR SEM ARREDONDAR, EXEMPLO: 2,5450 (FICARA 2,54)
   'SE A QUARTA CASA FOR ALGARISMO DIFERENTE DE ZERO, A 2S CASA  deverá ser AUMENTADA EM UMA unidade, EXEMPLO: 2,5451 (FICARÁ 2,55)
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  
   'SE A QUARTA CASA FOR IGUAL A ZERO
   If CInt(Mid(StrDecimal, 4, 1)) = 0 Then
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,5450  REMOVERÁ O 50 DO FINAL, RETORNANDO SOMENTE O 2,54
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
  
   'SE A QUARTA CASA FOR MAIOR QUE ZERO, ACRESCENTA EM 0,01 ARREDONDANDO PRA MAIS O VALOR DECIMAL COM 2 CASAS
   Else
      StrValor_Retornar = Mid(StrValor_Trabalhar, 1, Len(StrValor_Trabalhar) - 2) 'PEGA O VALOR SEM AS 2 ULTIMAS CASAS, EX: 2,3451  REMOVERÁ O 51 DO FINAL, RETORNANDO SOMENTE O 2,34
      StrValor_Retornar = CCur(StrValor_Retornar) + CCur("0,01")  'SOMA MAIS 1 CENTAVO
      Arredondamento_ABNT_NBR5891 = Format(StrValor_Retornar, "############0.00")
      Exit Function
   End If
  
  
Trata_Erros:
   If err.Number <> 0 Then
      MsgBox "Erro na funcao de ARREDONDAMENTO ABNT NBR 5891: " & err.Source & " " & err.Description, vbCritical
      Exit Function
   End If
End Function