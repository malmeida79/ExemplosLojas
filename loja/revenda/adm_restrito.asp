<%'-----------------------------------------------------------------------------------
  '####### 'VERIFICA SE O ADMINISTRADOR EST� LOGADO, CASO POSITIVO, ABRE A PAGINA
  '-----------------------------------------------------------------------------------
  IF Session("admin")<> "logado" THEN
  Response.Redirect "default.asp"
  END IF%>