<!--#include file="./khorClass.asp"-->
<!--#include file="./coboClass.asp"-->
<!--#include file="./scatterChartCommon.asp"-->
<%
  server.ScriptTimeout = 999999999
  IdPeriodo1 = reqn("IdPeriodo1")
  IdFactor1 = reqn("IdFactor1")
  Variable1 = request("Variable1")
  IdPeriodo2 = reqn("IdPeriodo2")
  IdFactor2 = reqn("IdFactor2")
  Variable2 = request("Variable2")
  IdEncuesta = reqn("IdEncuesta")
  filtroListado = request("filtroListado")
  'Agrupador de variable de configuración
  ECCOAgrupador=ECCOAgrupadoresReporte()

  infoCompleta = (IdPeriodo1<>0) AND (IdPeriodo2<>0) AND _
                 (IdFactor1<>0) AND (IdFactor2<>0) AND _
                 (Variable1<>"") AND (Variable2<>"") AND _
                 filtroListado<>""

  salida = "{"&""""&"animation"&""""&": { "&""""&"duration"&""""&": 10000}, "&""""&"datasets"&""""&": ["
  IF infoCompleta THEN
    set regMapa = new coboItemMapa
    regMapa.inicializa IdPeriodo1, IdPeriodo2, IdFactor1, IdFactor2, Variable1, Variable2, IdEncuesta, filtroListado, true
    set rs = getrs( conn, regMapa.getQuery(0) )	
	'response.write regMapa.getQuery(0)
	set puntosMapa = new AssocArrayClass	
	set puntosMapaPromedio = new AssocArrayClass 
	puntosMApaPromedio(CStr("total"))=0
    while not rs.EOF
      regMapa.getFromRS rs
	  cuadrante=regMapa.cuadrante(false)  	  
	  if VarType( puntosMapa(CStr(cuadrante))(CStr(CInt(regMapa.Dato1+.5)))(CStr(CInt(regMapa.Dato2+.5))))=9 then
		puntosMapa(CStr(cuadrante))(CStr(CInt(regMapa.Dato1+.5)))(CStr(CInt(regMapa.Dato2+.5)))=0		
	  end if
	  puntosMapa(CStr(cuadrante))(CStr(CInt(regMapa.Dato1+.5)))(CStr(CInt(regMapa.Dato2+.5)))= puntosMapa(CStr(cuadrante))(CStr(CInt(regMapa.Dato1+.5)))(CStr(CInt(regMapa.Dato2+.5)))+1
      if VarType(puntosMApaPromedio(CStr(cuadrante))(CStr("x")))=9 then 
		puntosMApaPromedio(CStr(cuadrante))(CStr("x"))=0
		puntosMApaPromedio(CStr(cuadrante))(CStr("y"))=0
		puntosMApaPromedio(CStr(cuadrante))(CStr("total"))=0
	  end if
	  puntosMApaPromedio(CStr(cuadrante))(CStr("x"))=puntosMApaPromedio(CStr(cuadrante))(CStr("x"))+regMapa.dato1
      puntosMApaPromedio(CStr(cuadrante))(CStr("y"))=puntosMApaPromedio(CStr(cuadrante))(CStr("y"))+regMapa.dato2
      puntosMApaPromedio(CStr(cuadrante))(CStr("total"))=puntosMApaPromedio(CStr(cuadrante))(CStr("total"))+1	
      puntosMApaPromedio(CStr("total"))=puntosMApaPromedio(CStr("total"))+1	  
	 rs.movenext
    wend
	rs.close
    set rs = nothing
    set regMapa = nothing
  END IF
  coma=""
  for i=1 to 4
    if i=1 then
	   cuadranteStr="I"
	   colorHTML="rgba(252,13,13,.5)"
	elseif i=2 then
	   cuadranteStr="II"
	   colorHTML="rgba(255, 12, 251,.5)"
	elseif i=3 then
	   cuadranteStr="III"	   
	   colorHTML="rgba(34, 37, 148,.5)"
	elseif i=4 then
		cuadranteStr="IV"
		colorHTML="rgba(36,148,34,.5)"
    end if
	CadenaExtra="-"&formatNumber((puntosMApaPromedio(cuadranteStr)(CStr("total"))*100)/puntosMApaPromedio(CStr("total")),2)&"% / "&puntosMApaPromedio(cuadranteStr)(CStr("total"))&" "
	salida=salida&coma&"{"&""""&"label"&""""&": "&""""&"Cuadrante "&cuadranteStr&CadenaExtra&""""&","&""""&"backgroundColor"&""""&": "&""""&colorHTML&""""&","&""""&"borderColor"&""""&": "&""""&colorHTML&""""&","&""""&"borderWidth"&""""&": 1, "&""""&"data"&""""&": ["
	coma2=""
	
	        if VarType( puntosMApaPromedio(CStr(cuadranteStr))(CStr("x")) )<>9 then
			    
			    promedioX=puntosMApaPromedio(CStr(cuadranteStr))(CStr("x")) / puntosMApaPromedio(CStr(cuadranteStr))(CStr("total"))
				promedioY=puntosMApaPromedio(CStr(cuadranteStr))(CStr("y")) / puntosMApaPromedio(CStr(cuadranteStr))(CStr("total"))
				salida = salida & coma2 &"{"&""""&"x"&""""&":"&formatNumber(promedioX,2)&","&""""&"y"&""""&":"&formatNumber(promedioY,2)&","&""""&"r"&""""&":"&formatNumber((puntosMApaPromedio(cuadranteStr)(CStr("total"))*100)/puntosMApaPromedio(CStr("total")),2)&"}"
				coma2=","
			end if
    coma= ","
	salida=salida&"]}"
  next
  salida = salida & "]}"
  Response.write salida
  coboClean

  conn.close
  set conn = nothing
%>