<% 





'#############################
'Novo sistema utiliza criptografia extrema em 64 bits usando o algoritimo BASE64 da RSA
'#############################

Function Cripto(StrCripto, BolAcao) 'Inнcio de da funзгo de criptografia. Aonde o parвmetro String й o valor que serб criptografado ou descriptografado. E o parвmetro BolAcao й um valor booleano (True ou False) para indicar se deve ser criptografado (True) ou descriptografado (False).

if application("Cripto_Ativa") = "Sim" then
If BolAcao Then
Cripto = EncodeBase64(StrCripto)
       Else
Cripto = DecodeBase64(StrCripto)
End If
else
Cripto = StrCripto
end if

End Function


Function EncodeBase64(inData)
On Error Resume Next
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut
  Dim sOut
  Dim I
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    nGroup = Oct(nGroup)
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    sOut = sOut + pOut
    If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1:
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2:
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  EncodeBase64 = sOut
End Function


Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function
'==============================================================


'==============================================================
Function DecodeBase64(ByVal base64String)
On Error Resume Next
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength
  Dim sOut
  Dim groupBegin
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "VirtuaStore OPEN", "String de criptografia com problemas. " & VBNewline & "Contate nosso suporte tйcnico pelo telefone (0xx51) 475-7545."
    Exit Function
  End If

  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes
    Dim CharCounter
    Dim thisChar
    Dim thisData
    Dim nGroup
    Dim pOut
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      thisChar = Mid(base64String, groupBegin + CharCounter, 1)
      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(Base64, thisChar) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Loja Virtual", "String de criptografia com problemas. " & VBNewline & "Contate nosso suporte tйcnico pelo telefone (0xx12) 9728-1573"
        Exit Function
      End If
      nGroup = 64 * nGroup + thisData
    Next
    nGroup = Hex(nGroup)
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    sOut = sOut & Left(pOut, numDataBytes)
  Next
  DecodeBase64 = sOut
End Function


%>