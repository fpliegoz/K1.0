<!--#include file="./khorClass.asp"-->
<!--#include file="./coboClass.asp"-->
<!--#include file="./coboAuxReportes.asp"-->
<%
  Server.ScriptTimeout = 999999
  conn.CommandTimeout = 200
  thispageid="ECCORepGral"
  thispage="ECCORepGral.asp"
  '====ECCOVARS
  '--Colores especiales de CLIMA
  
  ECCOcolorhtml1="rgba(216,45,45,1)" 'Inicial
  ECCOcolorhtml2="rgba(217,201,45,1)" 'Intermedio
  ECCOcolorhtml3="rgba(44,171,13,1)" 'Final
  ECCOlimit1=2
  ECCOlimit2=3
  ECCOColorHTMLMedia="#FF6400"
  ECCOColorHTMLModa="#008D00"
  ECCOColorHTMLDE="#0A44ED"
  ECCOColorHTMLIM="#FC0101"
  ECCOGraficaGrl=lblECCO_GraficaGeneral
  ECCODetalleFac=lblECCO_DetalleXfactor
  ECCODetalleImp="Detalle de Importancia"
  'Variable sucia para habilitar la agrupación de la grafica en 3 grupos, se debe tener el tamaño de la escala en 5 para que esto funcione
  confAgrupaVal=True
  
  '=== Inicializacion
  
  
  errmsg=""
  muestraNumeros = khorConfigValueWithDefault(596,true,0)<>0
  muestraPorcentajeMedias = khorConfigValueWithDefault(622,true,0)<>0
  mov=ucase(reqplain("mov"))
  tipoGrafica = ucase(reqs("tipoGrafica"))
  if tipoGrafica = "" THEN tipoGrafica = "BARRAS"
  if mov="PDF" then
    sesionFromRequest
    childwin = true
  else
    childwin = (reqplain("childwin") = thispageid)
  end if
  checaSesion ses_super&","&ses_adminid, "", ""
  modulosActivos = khorModulosActivos()
  validaEntrada khorPermisoModulo(Modulo_COBO,modulosActivos) OR khorPermisoModulo(Modulo_COBOConsulta,modulosActivos), "", thispageid
  tamEscala = coboEscalaLen()
  '=== Obtiene Empresa seleccionada o del contexto del usuario
  idf = reqn("idempresafiltro")
  if idf=0 then idf=khorEmpresaUsuario(0)
  '=== Obtiene parametros
  IdPeriodo = reqn("IdPeriodo")
  IdEncuesta = reqn("IdEncuesta")
  IdPersonal = request("IdPersonal")
  if muestraPorcentajeMedias then conMediaPercent = reqn("conMediaPercent")=1
  if IdPersonal<>"*" AND IdPersonal<>"" AND IdPersonal<>"BUSCADOR" then IdPersonal=reqCSV("IdPersonal")
  '=== Inicializa filtros
  filtroAuxParticipantes = strAdd( "Status=" & COBO_APLICADA, " AND ", iif(IdPeriodo>0,"IdPeriodo=" & IdPeriodo,"") )
  filSucursal = coboFiltroSucursalPar(idf)
  if coAnonima AND filSucursal<>"" then
    filtroAuxParticipantes = strAdd( filtroAuxParticipantes, " AND ", "IdSucursalPar IN (" & filSucursal & ")" )
  end if
  '=== Filtro de encuestas y validacion de la seleccionada
  filtroEncuestas = "EXISTS(SELECT * FROM co_Participante WHERE " & filtroAuxParticipantes & " AND co_Participante.IdEncuesta=co_Encuesta.IdEncuesta)"
  IdEncuesta = getBDnum("IdEncuesta","SELECT IdEncuesta FROM co_Encuesta WHERE IdEncuesta=" & IdEncuesta & " AND " & filtroEncuestas)
  if IdEncuesta>0 then
    filtroAuxParticipantes = strAdd( filtroAuxParticipantes, " AND ", "IdEncuesta=" & IdEncuesta )
  end if
  '=== Inicializa buscador
  coboInicializaBuscador
  '=== Obtiene personas seleccionadas: lista de id's seleccionados, BUSCADOR, ó *=todas las que aplicaron
  IF IdPeriodo<>0 THEN
    '=== Crea el filtro general de datos y la leyenda descriptiva con el numero de personas seleccionadas
    IF coAnonima THEN
      filtroListado = strAdd( filtroAuxParticipantes, " AND ", "IdPersonal=0" )
      if IdPersonal="BUSCADOR" then
        filtroListado = strAdd( filtroListado, " AND ", bobj.condicion("origenParticipante.IdParticipante") )
        sq = "SELECT COUNT(*) AS cuantos FROM co_Participante origenParticipante WHERE " & strAdd( "Status=" & COBO_APLICADA, " AND ", filtroListado )
        strPersonal = getBDnum("cuantos",sq) & ": " & lblKHOR_UsandoCondicionesBusqueda
      elseif IdPersonal="*" then
        sq = "SELECT COUNT(*) AS cuantos FROM co_Participante WHERE " & strAdd( filtroListado, " AND ", "Status=" & COBO_APLICADA )
        strPersonal = getBDnum("cuantos",sq) & ": " & lblCobo_TodasLasQueRespondieron
      else
        strPersonal = strLang( lblFRS_SeHa_nSeleccionado_N_X, "0||" & lblKHOR_Persona_s )
      end if
    ELSE
      '-- Filtro auxiliar de participantes con encuesta aplicada
      filtroAuxAplicados = "EXISTS(SELECT * FROM co_Participante WHERE " & filtroAuxParticipantes & " AND Status=" & COBO_APLICADA & " AND co_Participante.IdPersonal=Personal.IdPersonal)"
      '-- Inicializa filtro general de datos con condiciones de personas
      if idf<>0 then
        filtroListado = "Personal.IdSucursal=" & idf
      else
        filtroListado = khorCondicionSucursal("Personal.IdSucursal")
      end if
      '-- Crea el filtro de personas usado para la seleccion individual
      filtroPersonas = strAdd(filtroListado, " AND ", filtroAuxAplicados)
      '-- Modalidades de busqueda
      if IdPersonal="BUSCADOR" then
        '-- Agrega condiciones del buscador al filtro general de datos
        bobj.getRequest ""
        filtroListado = strAdd( filtroListado, " AND ", bobj.condicion("Personal") )
        sq = "SELECT COUNT(*) AS cuantos FROM Personal WHERE " & strAdd( filtroListado, " AND ", filtroAuxAplicados )
        strPersonal = getBDnum("cuantos",sq) & ": " & lblKHOR_UsandoCondicionesBusqueda
      elseif IdPersonal="*" then
        sq = "SELECT COUNT(*) AS cuantos FROM Personal WHERE " & strAdd( filtroListado, " AND ", filtroAuxAplicados )
        strPersonal = getBDnum("cuantos",sq) & ": " & lblCobo_TodasLasQueRespondieron
      else
        if IdPersonal<>"" then
          '-- Verifica que las personas seleccionadas cumplan con las demas condiciones
          IdPersonal = getBDlist( "IdPersonal", "SELECT IdPersonal FROM Personal WHERE " & strAdd( "Personal.IdPersonal IN (" & IdPersonal & ")", " AND ", filtroAuxAplicados ), false )
          '-- Agrega personas seleccionadas al filtro general de datos
          filtroListado = strAdd( filtroListado, " AND ", "Personal.IdPersonal IN (" & IdPersonal & ")" )
        end if
        strPersonal = strLang( lblFRS_SeHa_nSeleccionado_N_X, lenCSV(IdPersonal) & "||" & lblKHOR_Persona_s)
      end if
      '-- Agrega filtros para la vista
      filtroListado = strAdd( filtroAuxParticipantes, " AND ", filtroListado )
    END IF
  ELSE
    IdEncuesta = 0
    IdPersonal = ""
  END IF
  filtroPeriodo = "EXISTS(SELECT * FROM co_Participante WHERE Status=" & COBO_APLICADA & " AND co_Participante.IdPeriodo=ed_periodo.IdPeriodo)"
  '=== Datos del reporte
  numitems = 0
  redim itemRep(numitems)
  IF IdPeriodo>0 AND IdPersonal<>"" THEN
    '=== Incializa menu auxiliar de Factores-Grupos para la seleccion
    auxlst = "" '-- Lista de id's de todos
    menuFijoFactores = ""
    sq = "SELECT DISTINCT IdFactor, Grupo, Factor, (CASE WHEN Factor=Grupo THEN Factor ELSE (Grupo" & db_concat &  "':'" & db_concat &  "Factor) END) AS Descripcion" & _
        " FROM vco_encuesta WHERE " & replace( filtroEncuestas, "co_Encuesta", "vco_Encuesta" ) & " ORDER BY grupo, factor"
    set rs = getrs(conn,sq)
    while not rs.eof
      menuFijoFactores = menuFijoFactores & rsNum(rs,"IdFactor") & ":" & rs("Descripcion") & "|"
      auxlst = strAdd( auxlst, ",", rsNum(rs,"IdFactor") )
      rs.movenext
    wend
    rs.close
    set rs = nothing
    '=== Si no hay Factor(es) seleccionados, selecciona todos
    IdFactor = reqCSV("IdFactor")
    if IdFactor<>"" AND auxlst <> "" then
      IdFactor = getBDlist( "IdFactor", "SELECT IdFactor FROM co_Factor WHERE IdFactor IN ("&reqCSV("IdFactor")&") AND IdFactor IN ("&auxlst&")", false )
    end if
    if IdFactor="" then
      if auxlst = "" then
        IdFactor = 0
      else
        IdFactor = auxlst
      end if
    end if
    '=== Inicializa arreglo de resultados y leyenda de seleccion de factor(es)
    conIM = false
    numFactoresSelected = lenCSV(IdFactor)
    DesvStdFunction = db_DesvStd()
    if numFactoresSelected=1 then
      strFactor = descFromMenuFijo(menuFijoFactores,IdFactor)
      strItem = lblKHOR_Pregunta
      datoModa = "IdReactivo"
      sq = coboQueryReactivos( IdFactor, filtroListado )
    else
      strFactor = numFactoresSelected & " " & lblCOBO_FactoresSeleccionados
      strItem = lblCOBO_Factor
      datoModa = "IdFactor"
      sq = coboQueryFactores( IdFactor, filtroListado )
    end if
    set rs = getrs(conn,sq)
    while not rs.eof
      numitems = numitems + 1
      redim preserve itemRep(numitems)
      set itemRep(numitems) = new coboItemReporte
      itemRep(numitems).getFromRS rs
      'itemRep(numitems).getModas datoModa, filtroListado
      if itemRep(numitems).Escalas = COBO_Ambas then conIM = true
      rs.movenext
    wend
    rs.close
    set rs = nothing
    '=== Agrega renglon de Resultado General
    sqResGral = ""
    if numFactoresSelected=1 then
      '--- Se listan Reactivos (un solo Factor), el Resultado General es el del Factor.
      sqResGral = coboQueryFactores( IdFactor, filtroListado )
    else
      '--- Se listan varios Factores
      IdGrupo = getBDlist( "IdGrupo", "SELECT DISTINCT IdGrupo FROM co_Factor WHERE IdFactor IN ("&IdFactor&")", false )
      if lenCSV(IdGrupo)=1 then
        '-- Solo si TODOS los Factores son del mismo Grupo
        if numFactoresSelected = getBDnum("cuantos","SELECT COUNT(*) AS cuantos FROM co_Factor WHERE IdGrupo="&IdGrupo) then
          '-- y estan seleccionados TODOS los Factores de ese Grupo
          '-- el Resultado General es el del Grupo
          sqResGral = coboQueryGrupos( IdGrupo, filtroListado )
        end if
      end if
    end if
    if sqResGral<>"" then
      set rs = getrs(conn,sqResGral)
      if not rs.eof then
        numitems = numitems + 1
        redim preserve itemRep(numitems)
        set itemRep(numitems) = new coboItemReporte
        itemRep(numitems).getFromRS rs
        itemRep(numitems).Id = 0  'Se utiliza cero para marcar el renglon de Resultado General
        itemRep(numitems).Descripcion = uCase(lblFRS_General) & ":" & itemRep(numitems).Descripcion
      end if
      rs.close
      set rs = nothing
    end if
    if ubound(itemRep)=0 then
      IdFactor=""
      errmsg=lblCobo_NoHayEvalConCaracIndicadas
    end if
  END IF
  '--- Titulos
  titulo=lblCOBO_ClimaOrganizacional
  tit1=lblCOBO_RepGeneral
  tit2=getDescripcion("ED_Periodo","IdPeriodo","Periodo",IdPeriodo)

'================================================================================'
%>
<HTML>
<HEAD>
<TITLE><%=khorAppName()%> - <%=titulo%> - <%=tit1%></TITLE>

<% includeJS %>
<script src="./js/Chart.js"></script>
<script src="./js/Chart.bundle.js"></script>

<%
IF mov<>"PDF" THEN
%>
<script language="JavaScript">
<!--
function cambiaGrafica(tipo) {
  sendval('','mov','','tipoGrafica',tipo);
}

function cambiaLista() {
  sendval('','mov','');
} <%
  IF IdPeriodo>0 THEN
    coboBuscadorJS
    IF NOT coAnonima THEN %>
function seleccionaPersona() {
  var presel = getValor('IdPersonal');
  if (presel == '*' || presel == 'BUSCADOR') presel = '';
  abreSeleccion('PERSONA',true,presel,null,null,null,null,'<%=setSelectionFilter("PERSONA",filtroPersonas)%>');
} <%
    END IF %>
function setSeleccion(tipo,lista) {
  if (tipo == 'FACTOR') {
    sendval('','IdFactor',lista);
  } else {
    sendval('','IdPersonal',lista);
  }
} <%
    IF IdPersonal<>"" THEN
      IF pdf_enabled() THEN %>
function myPrintPage() { <%
  pdfkey = initPDFurl( thispageid & "_" & IdPeriodo & "_" & IdEncuesta & "_" & getSesionId(), _
                        pdf_URL() & thispage & "?mov=pdf&idempresafiltro=" & idf & "&IdPeriodo=" & IdPeriodo & "&IdEncuesta=" & IdEncuesta & "&IdPersonal=" & IdPersonal & "&IdFactor=" & IdFactor & "&" & reqbuscador & "=" & buscador & "&" & bobj.keyBuscadorSesion(buscador) & "=" & encrypt(bobj.condicionURL()) ) %>
  openPDFjob('<%=pdfkey%>');
} <%
      END IF %>
var tabcount=3;
var tab;
function tabSel(t) {
  if (t!=tab) {
    for (var i=0; i<tabcount; i++) {
      //setVisible('tabC'+i,i==t);
      ot=MM_findObj('tab'+i);
      ot.className=(i==t)?'tabOpen':'tabClose';
    }
    tab=t;
  }
}

function oculta(paso){  
   if(paso==2){
		$("#grafica").hide("slow");
		$("#grafica3").hide("slow");
		$("#grafica2").show("slow");
   }
   if(paso==1){
		$("#grafica").show("slow");
		$("#grafica3").hide("slow");
		$("#grafica2").hide("slow");
   }
   if(paso==3){
		$("#grafica").hide("slow");
		$("#grafica2").hide("slow");
		$("#grafica3").show("slow");
   }
}
function seleccionaFactor() {
  abreSeleccionSimple('FACTOR',true,getValor('IdFactor'),'<%=strJS(lblCOBO_Factor_es)%>','<%=menuFijoFactores%>')
} <%
    END IF
  END IF %>
//-->
</script> <%
END IF %>
</HEAD>
<BODY>
<!--#include file="khorHeader.asp"-->
<table class="pagetable" cellspacing="5" width="<%=khorWinWidth()%>">
  <tr>
    <td class="pagetitle">
      <%call ponEncabezado(titulo,tit1,tit2)%>
    </td>
  </tr>
  <tr>
    <td class="bordegris">
    <%IF errmsg<>"" THEN%><div class="alerta"><%=errmsg%></div><%END IF%>
    <%ponLigaBottom true%>

<%IF mov<>"PDF" THEN%>
    <form name="TrueForm" action="<%=thispage%>" method="POST" class="plana" onSubmit="return false;">
<%END IF%>
      <table border=0 cellspacing=0 cellpading=5 align="center">
        <%
          prehtml = "<tr><td><b>" & lblKHOR_Sucursal & ":</b></td><td colspan=""3"">"
          posthtml = "</td></tr>"
          coboSeleccionSucursal idf, "cambiaLista()"" style=""font-size:90%;width:auto;""", prehtml, posthtml
        %>
        <tr>
          <td><B><%=lblKHOR_Periodo%>:</B></td>
          <td colspan="3"> <%
            IF mov<>"PDF" THEN %>
            <select name="IdPeriodo" onChange="cambiaLista()" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
            <option value="0">- <%=lblFRS_Seleccione_ & lblKHOR_Periodo%> -</option>
            <% optionFromCatObj CAT_PERIODO, idf, IdPeriodo, filtroPeriodo %>
            </select> <%
            ELSE
              response.write descFromTable("ED_Periodo","IdPeriodo","Periodo",IdPeriodo)
            END IF %>
          </td>
        </tr>
        <%IF IdPeriodo>0 THEN%>
        <tr>
          <td><B><%=lblCOBO_Encuesta%>:</B></td>
          <td colspan="3"> <%
            IF mov<>"PDF" THEN %>
            <select name="IdEncuesta" onChange="cambiaLista()" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
            <option value="0">- <%=lblFRS_Cualquiera%> -</option>
            <% optionFromCat "co_Encuesta","IdEncuesta","Encuesta",false,filtroEncuestas,IdEncuesta %>
            </select> <%
            ELSEIF IdEncuesta=0 THEN
              response.write lblFRS_Cualquiera
            ELSE
              response.write descFromTable("co_Encuesta","IdEncuesta","Encuesta",IdEncuesta)
            END IF %>
          </td>
        </tr>
        <TR>
          <TD><B><%=iif(coAnonima,lblKHOR_Personas,lblKHOR_Persona_s)%>:</B></TD>
          <TD> <%
            IF mov<>"PDF" THEN %>
            <INPUT  class="quitabordes" style="WIDTH:300px;" readOnly="true" name=strPersonal value="<%=strPersonal%>" >
            <INPUT type=hidden name="IdPersonal" value="<%=IdPersonal%>"> <%
            ELSE
              response.write strPersonal
            END IF %>
          </TD> <%
            IF coAnonima OR mov="PDF" THEN %>
          <TD colspan="2"><%
            ELSE %>
          <TD>
            <INPUT type=button value="<%=lblFRS_Seleccionar%>" onclick="seleccionaPersona();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
          </TD>
          <TD> <%
            END IF
            IF mov<>"PDF" THEN %>
            <INPUT type=button value="<%=lblFRS_Buscar%>" onclick="callBuscador();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
            <INPUT type=button value="<%=lblFRS_Todas%>" onclick="setSeleccion('PERSONA','*');" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);> <%
            END IF %>
          </TD>
        </TR> <%
            IF IdPersonal="BUSCADOR" THEN %>
        <TR>
          <TD><B><%=lblCOBO_Condiciones%>:</B></TD>
          <TD colspan="3">
            <div class="bordeGris" style="margin:1px;">
              <%=bobj.explicacion("<BR>")%>
            </div>
          </TD>
        </TR> <%
            END IF
            IF IdPersonal<>"" THEN %>
        <TR>
          <TD><B><%=lblCOBO_Factor_es%>:</B></TD>
          <TD<%=iif(mov="PDF"," colspan=2","")%>> <%
            IF mov<>"PDF" THEN %>
            <INPUT  class="quitabordes" style="WIDTH:300px;" readOnly="true" name=strFactor value="<%=strFactor%>" >
            <INPUT type=hidden name="IdFactor" value="<%=IdFactor%>"> <%
            ELSE
              response.write strFactor
            END IF %>
          </TD> <%
            IF mov<>"PDF" THEN %>
          <TD>
            <INPUT type=button value="<%=lblFRS_Seleccionar%>" onclick="seleccionaFactor();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
          </TD>
          <TD> <%
              if numFactoresSelected < lenMenuFijo(menuFijoFactores) then %>
            <INPUT type=button value="<%=lblFRS_Todos%>" onclick="setSeleccion('FACTOR','');" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);> <%
              end if %>&nbsp;
          </TD> <%
            END IF %>
        </TR> <%
            END IF 
            IF muestraPorcentajeMedias THEN%>
        <tr>
          <td align="right"><%inputCheckbox "conMediaPercent",1,conMediaPercent,"","onChange='sendval("""");'"%></td>
          <td colspan="3"> 
            <%=lblCOBO_MostrarMediaEnPorcentaje%>
          </td>
        </tr>
          <%END IF
        END IF%>
      </table>
	  </br>
      <%IF IdPeriodo>0 AND IdPersonal<>"" AND IdFactor<>"" THEN%>
        <!--==================== GRAFICAS ====================-->
		
      <%IF mov<>"PDF" THEN%>
      <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
        <tr>
          <td>
            <table border="0" cellspacing="0" cellpadding="5" width="100%">
              <tr>
                <td id="tab0" nowrap onclick="oculta(1);tabSel(0);"><b><%=ECCOGraficaGrl%></b></td>
                <td id="tab1" nowrap onclick="oculta(2);tabSel(1);"><b><%=ECCODetalleFac%></b></td>
				<%if conIM then%>
			    <td id="tab2" nowrap onclick="oculta(3);tabSel(2);"><b><%=ECCODetalleImp%></b></td>
				<%end if%>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td class="tabC" valign="top">
      <%END IF%>
		
	      <center>  <div id="grafica" width="100%" height="450px">
			<div class="wrapper">
			<canvas id="chart-0" width="900px" height="450px"></canvas>
		    </div><center>
			<table border="0" cellspacing="1" class="tsmall" width="600px">
			  <tr class="celdaTit" style="text-align:center;">
				<td colspan="2">Detalle General</td>
				<td>Media</td>
				<td>Desviación estandar superior</td>
				<td>Desviación estandar inferior</td>
				<td>Moda</td>
				<%if conIM then%>
				<td>Importancia</td>
				<%end if%>
			 </tr>
			<%
			'Nueva Tabla del detalle con porcentajes por cada elemento
			 for z=1 to numitems
			 
			%>          			
             <tr class="celdaDark" style="text-align:center;">
				<td><%=numRomano(z)%></td>
				<td><%=itemRep(z).Descripcion%></td>
				<td><%=coboFormatNum(itemRep(z).SA)%></td>
				<td><%=coboFormatNum(itemRep(z).SA+itemRep(z).SAdev)%></td>
				<td><%=coboFormatNum(itemRep(z).SA-itemRep(z).SAdev)%></td>
				<td><%=coboFormatNum(itemRep(z).SAmoda)%></td>
				<%if conIM then%>
				<td><%=coboFormatNum(itemRep(z).IM)%></td>
				<%end if%>
			 </tr>
			 <%
			 next
			%>
             </table>			 
			</div></center>
			</div>
			<div <%if mov<>"PDF" then%>hidden<%end if%> id="grafica2" width="800px" height="650px">		
			<div class='wrapper'><canvas id='chart-z' width='700px' height='650px'></canvas></div>
			</br><center>
			 <table border="0" cellspacing="1" class="tsmall" width="600px">
			  <tr class="celdaTit" style="text-align:center;">
				<td colspan="2">Detalle por elemento</td>
				<%
				for colorestotales=1 to tamescala				
					cadenacolores=cadenacolores&","
                       coloragregar=ECCOcolorhtml3					
					if colorestotales <=ECCOlimit2 then 
					   coloragregar=ECCOcolorhtml2
					end if 
					if colorestotales <=ECCOlimit1 then 
					   coloragregar=ECCOcolorhtml1
					end if
                    response.write "<td style='background-color:"&coloragregar&";'>"&colorestotales&"</td>"				
	            next  
				%>
			 </tr>
			<%
			'Nueva Tabla del detalle con porcentajes por cada elemento
			 for z=1 to numitems
			 
			%>          			
             <tr class="celdaDark" style="text-align:center;">
				<td><%=numRomano(z)%></td>
				<td><%=itemRep(z).Descripcion%></td>
				<%
					for y=1 to tamescala
					numero=iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(z).fSA(y)*100,1), formatNumber(div( itemRep(z).fSA(y), itemRep(z).fTot )*100,1))
				%>
				<td><%=numero%>%</td>
				<%next%>
			 </tr>
			 <%
			 next
			%>
             </table>			 
			</div></center>
			<%if conIM then%>
			<div <%if mov<>"PDF" then%>hidden<%end if%> id="grafica3" width="800px" height="650px">		
			<div class='wrapper'><canvas id='chart-z2' width='700px' height='650px'></canvas></div>
			</br>
			<center>
			 <table border="0" cellspacing="1" class="tsmall" width="600px">
			  <tr class="celdaTit" style="text-align:center;">
				<td colspan="2"><%=lblECCODetalleImp%></td>
				<%
				 cadenacolores=""""&""""
		        for colorestotales=1 to tamescala				
					cadenacolores=cadenacolores&","			
					if colorestotales =1 then 
					   coloragregar="#00FA9A"
					end if 
					if colorestotales =2 then 
					   coloragregar="#00CED1"
					end if
					if colorestotales =3 then 
					   coloragregar="#00BFFF"
					end if 
					if colorestotales =4 then 
					   coloragregar="#008B8B"
					end if 
					if colorestotales =5 then 
					   coloragregar="#008080"
					end if
                    response.write "<td style='background-color:"&coloragregar&";'>"&colorestotales&"</td>"						
	            next   	
				%>
			 </tr>
			<%
			'Nueva Tabla del detalle con porcentajes por cada elemento
			 for z=1 to numitems
			 
			%>          			
             <tr class="celdaDark" style="text-align:center;">
				<td><%=numRomano(z)%></td>
				<td><%=itemRep(z).Descripcion%></td>
				<%
					for y=1 to tamescala
					numero=iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(z).fSA(y)*100,1), formatNumber(div( itemRep(z).fSA(y), itemRep(z).fTot )*100,1))
				%>
				<td><%=numero%>%</td>
				<%next%>
			 </tr>
			 <%
			 next
			%>
             </table>			 
			</div></center>
			<%end if%>
            <div id="tabC0" style="display:<%=displayStyle(mov="PDF2")%>;text-align:center;"> 
			<style>
 
   .graph_container{
    display:block;
    width:550px;
  }
  </style>
		<script>
		var presets = window.chartColors;
		var data = {
			labels: [<%
			         cadena=""
			         coma=""
			         for z=1 to numitems
						pivot=InStr(itemRep(z).Descripcion,":")
						total=Len(itemRep(z).Descripcion)
						extract=total-pivot
			            cadena=cadena&coma&""""&numRomano(z)&" "&Replace(Right(itemRep(z).Descripcion,extract),Chr(34),"'")&""""
						coma=","
			         next
					 response.write strJS(cadena)
					 %>],
			datasets: [{
				borderColor: "<%=ECCOColorHTMLMedia%>",
				steppedLine: false,
				data: [<%
				      cadena=""
			         coma=""
			         for z=1 to numitems
					    numero= coboFormatNum(itemRep(z).SA)
					    cadena=cadena&coma&""""&numero&""""
						coma=","
			         next
				      response.write cadena
					  %>
					  ],
				label: 'Media',
				
				fill:false
				
			}, {
				borderColor: "<%=ECCOColorHTMLDE%>",
				data: [<%
				      
				      cadena=""
			         coma=""
			         for z=1 to numitems
					    numero= coboFormatNum(itemRep(z).SA+itemRep(z).SAdev)
					    cadena=cadena&coma&""""&numero&""""
						coma=","
			         next
				      response.write cadena
					  
				       %>
				],
				label: 'Desviación estandar Superior',
				fill: '-1'
				
			}, {
				borderColor: "<%=ECCOColorHTMLDE%>",
				data: [<%
				      
				      cadena=""
			         coma=""
			         for z=1 to numitems
					    numero= coboFormatNum(itemRep(z).SA-itemRep(z).SAdev)
					    cadena=cadena&coma&""""&numero&""""
						coma=","
			         next
				      response.write cadena
					  
				       %>
				],
				label: 'Desviación estandar Inferior',
				
				fill: '-1'
			}, {
				borderColor: "<%=ECCOColorHTMLModa%>",
				data: [<%
				       cadena=""
			         coma=""
			         for z=1 to numitems
					    cadena=cadena&coma&""""&coboFormatNum(itemRep(z).SAmoda)&""""
						coma=","
			         next
				      response.write cadena
				       %>
				],
				label: 'Moda',
				fill:false
			}<%if conIM then%>,{
				borderColor: "<%=ECCOColorHTMLIM%>",
				data: [<%
				       cadena=""
			         coma=""
			         for z=1 to numitems
					    cadena=cadena&coma&""""&coboFormatNum(itemRep(z).IM)&""""
						coma=","
			         next
				      response.write cadena
				       %>
				],
				label: 'Importancia',
				fill:false
			}
			
			
			<%end if%>]
		};

		var options = {
			
			responsive:false,
			lineHeight:'1.0',
			<% if mov="PDF" then %>
			 animation: false,
			<%end if%>
			elements: {
				line: {
					tension: 0.000001
				}
			},
			scales: {
				yAxes: [{
					stacked: false
				}],
				xAxes: [{
						ticks: {
							autoSkip: false
						}
				}]
			},
			tooltips: {
                    mode: 'index',
                    intersect: false
                },
			hover: {
                    mode: 'nearest',
                    intersect: true
                },
			plugins: {
				filler: {
					propagate: false
				},
				samples_filler_analyser: {
					target: 'chart-analyser'
				}
			}
		};

		$( document ).ready(function() {
			var chart = new Chart('chart-0', {
				type: 'line',
				data: data,
				options: options
			});
		});
		
		
		var options2 = {
    tooltips: {
        enabled: true,
		 callbacks: {
			label: function(tooltipItem, data) { 
					var dataset = data.datasets[tooltipItem.datasetIndex];
					return dataset.data[tooltipItem.index] + "%"
				   }
			}
    },
	responsive:false,
    
    scales: {
        xAxes: [{
            ticks: {
                beginAtZero:true,
                fontFamily: "'Open Sans Bold', sans-serif",
                fontSize:11,
			    autoSkip: true,
                max: 100
                
            },
            scaleLabel:{
                display:false
            },
            gridLines: {
            }, 
            stacked: true
        }],
        yAxes: [{
            gridLines: {
                display:false,
                color: "#fff",
                zeroLineColor: "#fff",
                zeroLineWidth: 0
            },
            ticks: {
                fontFamily: "'Open Sans Bold', sans-serif",
                fontSize:11
            },
            stacked: true
        }]
    },
    legend:{
        display:false
    },
    <% if mov="PDF" then %>
			 animation: false,
	<%else%>
    animation: {
        onComplete: function () {
            var chartInstance = this.chart;
            var ctx = chartInstance.ctx;
            ctx.textAlign = "left";
            ctx.font = "9px Open Sans";
            ctx.fillStyle = "#fff";
<%IF confAgrupaVal = true then %>
  Chart.helpers.each(this.data.datasets.forEach(function (dataset, i) {
                var meta = chartInstance.controller.getDatasetMeta(i);
                Chart.helpers.each(meta.data.forEach(function (bar, index) {
                    data = dataset.data[index];
                    if(i==0){
					   // alert(data);
                       // ctx.fillText(data,  bar._model.x-80 ,bar._model.y-7);
                    } else if (i==1){ 
					   // ctx.fillText(data,  bar._model.x-80 ,bar._model.y-7);
					} else if (i==2){ 
					   // ctx.fillText(data,  bar._model.x-30 ,bar._model.y-7);
					}else{
                       // ctx.fillText(data, bar._model.x-20, bar._model.y-7);
                    }
                }),this)
            }),this);
        }
    },

<%ELSE%>
            Chart.helpers.each(this.data.datasets.forEach(function (dataset, i) {
                var meta = chartInstance.controller.getDatasetMeta(i);
                Chart.helpers.each(meta.data.forEach(function (bar, index) {
                    data = dataset.data[index];
                    if(i==0){
					   // alert(data);
                        //ctx.fillText(data,  bar._model.x-8 ,bar._model.y-7);
                    } else if (i==1){ 
					    //ctx.fillText(data,  bar._model.x-2 ,bar._model.y-7);
					} else if (i==2){ 
					    //ctx.fillText(data,  bar._model.x ,bar._model.y-7);
					}else{
                        //ctx.fillText(data, bar._model.x-20, bar._model.y-7);
                    }
                }),this)
            }),this);
        }
    },
<%END IF%>
<%end if%>
    pointLabelFontFamily : "Quadon Extra Bold",
    scaleFontFamily : "Quadon Extra Bold",
};
	
                $( document ).ready(function() {
		<%      cadenacolores=""""&""""
		        for colorestotales=1 to tamescala				
					cadenacolores=cadenacolores&","
                       coloragregar=ECCOcolorhtml3					
					if colorestotales <=ECCOlimit2 then 
					   coloragregar=ECCOcolorhtml2
					end if 
					if colorestotales <=ECCOlimit1 then 
					   coloragregar=ECCOcolorhtml1
					end if
                    cadenacolores=cadenacolores&""""&coloragregar&""""					
	            next   	
		%>
				    var colors = [<%=cadenacolores%>];
					var chart = new Chart('chart-z', {
							type: 'horizontalBar',
							data: { labels:[
        <%  cadena =""
		    coma =""
			
		    for j=1 to  numitems   
                        cadena=cadena&coma&""""&Replace(iif(Len(itemRep(j).Descripcion)>75,Left(itemRep(j).Descripcion,63)&"...",itemRep(j).Descripcion),Chr(34),"'")&""""
						coma=","		
            next
			response.write strJS(cadena)
			coma=""
			coma2=""%>
			                       ],
						             datasets: [
									 <%	
                   
				      for y=1 to tamescala%>		
                                            <%=coma2%>{					  
											data: [<%  cadena =""
													coma =""
													
													for j=1 to  numitems
                            if tamescala =5 AND confAgrupaVal THEN
                                if y=2 OR y=5 then
                                  suma = CDbl(iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fSA(y)*100,1), formatNumber(div( itemRep(j).fSA(y), itemRep(j).fTot )*100,1)))+CDbl(iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fSA(y-1)*100,1), formatNumber(div( itemRep(j).fSA(y-1), itemRep(j).fTot )*100,1)))
                                  cadena=cadena&coma&""""&suma&""&""""	
                                  coma=","
                                elseif y=3 then
                                  cadena=cadena&coma&""""&iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fSA(y)*100,1), formatNumber(div( itemRep(j).fSA(y), itemRep(j).fTot )*100,1))&""""						
                                  coma=","    
                                end if
                            else
                                cadena=cadena&coma&""""&iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fSA(y)*100,1), formatNumber(div( itemRep(j).fSA(y), itemRep(j).fTot )*100,1))&""""						
                                coma=","
                            end if
													next
													response.write cadena
													%>													
													],
										
											backgroundColor: colors[<%=y%>],
											hoverBackgroundColor: colors[<%=y%>]
											<% 
											    coma2=","
											
											
											%>
											}
			        <%         
			    		next
                   %>
									]},
							options: options2
						});
		<%if conIM then%>
				<%      cadenacolores=""""&""""
		        for colorestotales=1 to tamescala				
					cadenacolores=cadenacolores&","			
					if colorestotales =1 then 
					   coloragregar="#00FA9A"
					end if 
					if colorestotales =2 then 
					   coloragregar="#00CED1"
					end if
					if colorestotales =3 then 
					   coloragregar="#00BFFF"
					end if 
					if colorestotales =4 then 
					   coloragregar="#008B8B"
					end if 
					if colorestotales =5 then 
					   coloragregar="#008080"
					end if
                    cadenacolores=cadenacolores&""""&coloragregar&""""					
	            next   	
		%>
				    var colors2 = [<%=cadenacolores%>];
				//////Cambios grafica 3
				var chart2 = new Chart('chart-z2', {
							type: 'horizontalBar',
							data: { labels:[
        <%  cadena =""
		    coma =""
			
		    for j=1 to  numitems   
                        cadena=cadena&coma&""""&Replace(iif(Len(itemRep(j).Descripcion)>75,Left(itemRep(j).Descripcion,63)&"...",itemRep(j).Descripcion),Chr(34),"'")&""""
						coma=","		
            next
			response.write strJS(cadena)
			coma=""
			coma2=""%>
			                       ],
						             datasets: [
									 <%	
                   
				      for y=1 to tamescala%>		
                                            <%=coma2%>{					  
											data: [<%  cadena =""
													coma =""
													
													for j=1 to  numitems
                            if tamescala =5 AND FALSE THEN
                                if y=2 OR y=5 then
                                  suma = CDbl(iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fIM(y)*100,1), formatNumber(div( itemRep(j).fIM(y), itemRep(j).fTot )*100,1)))+CDbl(iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fIM(y-1)*100,1), formatNumber(div( itemRep(j).fIM(y-1), itemRep(j).fTot )*100,1)))
                                  cadena=cadena&coma&""""&suma&""&""""	
                                  coma=","
                                elseif y=3 then
                                  cadena=cadena&coma&""""&iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fIM(y)*100,1), formatNumber(div( itemRep(j).fIM(y), itemRep(j).fTot )*100,1))&""""						
                                  coma=","    
                                end if
                            else
                                cadena=cadena&coma&""""&iif(muestraNumeros and sqResGral<>"",formatNumber(itemRep(j).fIM(y)*100,1), formatNumber(div( itemRep(j).fIM(y), itemRep(j).fTot )*100,1))&""""						
                                coma=","
                            end if
													next
													response.write cadena
													%>													
													],
										
											backgroundColor: colors2[<%=y%>],
											hoverBackgroundColor: colors2[<%=y%>]
											<% 
											    coma2=","
											
											
											%>
											}
			        <%         
			    		next
                   %>
									]},
							options: options2
						});
					
		<%end if%>
    });			
	</script>	
			
			
			<%
              if tipoGrafica = "PIE" and sqResGral<>"" then
                ValoresAS = ""
                ValoresIM = ""
                EtiqAS = ""
                EtiqIM = ""
                Color = ""
                for j=1 to tamEscala
                  Color = strAdd(Color ,",", ListaColores(j-1))
                    ValoresAS = strAdd(ValoresAS,",",itemRep(numitems).fSA(j))
                    EtiqAS = strAdd(EtiqAS,",",j&"("&coboFormatPorcentaje(div( itemRep(numitems).fSA(j), itemRep(numitems).fTot))&")")
                    if conIM then 
                      ValoresIM = strAdd(ValoresIM,",",itemRep(numitems).fIM(j))
                      EtiqIM = strAdd(EtiqIM,",",j&"("&coboFormatPorcentaje(div( itemRep(numitems).fIM(j), itemRep(numitems).fTot))&")")
                    end if
                next%>
                <table align="center">
                  <tr>
                    <td align="center"><%=lblCOBO_SituacionActual%></td><%
                    if conIM then%>
                    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td  align="center"><%=lblCOBO_Importancia%></td><%
                    end if%>
                  </tr>
                  <tr>
                    <td>
                      <img src="<%="./charts/PieChart.ashx?titles="&EtiqAS&"&values="&ValoresAS&"&colors="&Color%>" border="0" style="width:200px; height:auto;"/>
                    </td><%
                    if conIM then%>
                    <td></td>
                    <td>
                      <img src="<%="./charts/PieChart.ashx?titles="&EtiqIM&"&values="&ValoresIM&"&colors="&Color%>" border="0" style="width:200px; height:auto;"/>
                    </td><%
                    end if%>
                  </tr>
                </table><%
              else
                for i=1 to numitems
                  if itemRep(i).Id=0 then
                    Etiquetas = strAdd( Etiquetas, ",", "G" )
                  else
                    Etiquetas = strAdd( Etiquetas, ",", numRomano(i) )
                  end if
                  auxs = coboFormatNum(itemRep(i).SA)
                  if conIM then auxs = auxs & "," & coboFormatNum(itemRep(i).IM)
                  Valores = strAdd( Valores, "@", auxs )
                next
                Colores = ListaColores(coboColorSA)
                Detalle = lblCOBO_SituacionActual_Abrev
                if conIM then
                  Colores = Colores & "," & ListaColores(coboColorIM)
                  Detalle = Detalle & "," & lblCOBO_Importancia_Abrev
                end if
                GraficaBarras "graficaRes", khorWinWidthPix(), "", Etiquetas, Valores, Colores, Detalle, 1, tamEscala, iif (tipoGrafica = "PIE","",tipoGrafica), "", "", "", "" 
              end if%>
            </div>
            <div id="tabC1" style="display:<%=displayStyle(mov="PDF2")%>;text-align:justify;"> <%
              if tipoGrafica = "PIE" and sqResGral<>"" then
                ValoresAS = ""
                sumIncluye = 0
                sumNoIncluye = 0
                EtiqAS = ""
                EtiqIncluye = ""
                EtiqNoIncluye = ""
                for j=1 to tamEscala
                    if inCSV(coboItemsFrec,j)>=0 then
                      sumIncluye = sumIncluye + itemRep(numitems).fSA(j)
                      EtiqIncluye = strAdd(EtiqIncluye,"-",j)
                    else
                      sumNoIncluye = sumNoIncluye + itemRep(numitems).fSA(j)
                      EtiqNoIncluye = strAdd(EtiqNoIncluye,"-",j)
                    end if
                    ValoresAS = sumIncluye & "," & sumNoIncluye
                    EtiqAS = EtiqIncluye & "(" & coboFormatPorcentajecoboFormatPorcentaje(div( sumIncluye, itemRep(numitems).fTot )) & ")" & "," & EtiqNoIncluye & "(" & coboFormatPorcentaje(div( sumNoIncluye, itemRep(numitems).fTot)) & ")"
                next%>
                <table align="center">
                  <tr>
                    <td align="center"><%=lblCOBO_SituacionActual%></td>
                  </tr>
                  <tr>
                    <td>
                      <img src="<%="./charts/PieChart.ashx?titles="&EtiqAS&"&values="&ValoresAS&"&colors="&ListaColores(coboColorFrec)&","&ListaColores(coboColorFrecResto)%>" border="0" style="width:200px; height:auto;"/>
                    </td>
                  </tr>
                </table><%
              else
                Valores = ""
                for i=1 to numitems
                  suma = 0
                  for j=1 to tamEscala
                    if inCSV(coboItemsFrec,j)>=0 then
                      suma = suma + itemRep(i).fSA(j)
                    end if
                  next
                  suma = div( suma, itemRep(i).fTot ) * 100
                  Valores = strAdd( Valores, "@", coboFormatNum(suma) )
                next
                GraficaBarras "graficaFrec", khorWinWidthPix(), "", Etiquetas, Valores, ListaColores(coboColorFrec), "", 10, 100, iif (tipoGrafica = "PIE","",tipoGrafica), "", "", "", "" 
              end if%>
            </div>
      <%IF mov<>"PDF" THEN%>
          </td>
        </tr>
      </table>
      <script language="JavaScript">tabSel(0); </script>
      <%END IF%>
      <DIV align="center" class="tsmall"> <%
        response.write "<b>" & lblCOBO_SituacionActual & ":</b> " & replace(menuFijoCOsa,"|",", ")
        if conIM then
          response.write "<br /><b>" & lblCOBO_Importancia & ":</b> " & replace(menuFijoCOim,"|",", ")
        end if %>
      </DIV>
      <br>
        <!--==================== RESULTADOS ====================-->	
		<div>
     <!-- <TABLE border=0 cellspacing=1 cellpadding=1 align="center">
        <tr class="celdaTit" align="center">
          <td rowspan="2" width="2">#</td>
          <td rowspan="2"><%=strItem%></td>
          <td colspan="<%=(3+tamEscala)%>" class="tsmall" align="center"><%=lblCOBO_SituacionActual%></td> <%
          if conIM then %>
          <td colspan="<%=(3+tamEscala)%>" class="tsmall" align="center"><%=lblCOBO_Importancia%></td>
          <td rowspan="2"><%=lblCOBO_Importancia_Abrev%> - <%=lblCOBO_SituacionActual_Abrev%></td>
          <td rowspan="2"><%=lblFRS_Recomendacion%></td> <%
          end if %>
        </tr>
        <tr class="celdaTit" align="center"> <%
          for j=1 to tamEscala
            if inCSV(coboItemsFrec,j)>=0 then
              modstyle = " style=""filter:;background:#" & ListaColores(coboColorFrec) & "; color:#" & colorEtiqueta(coboColorFrec) & ";"""
            else
              modstyle = ""
            end if %>
          <td class="tsmall" align="center"<%=modstyle%>><%=j%></td> <%
          next %>
          <td class="tsmall" align="center" style="filter:;background:#<%=ListaColores(coboColorSA)%>; color:#<%=colorEtiqueta(coboColorSA)%>"><%=lblFRS_Media%></td>
          <td class="tsmall" align="center"><%=lblFRS_DesvStd_Abrev%></td>
          <td class="tsmall" align="center"><%=lblFRS_Moda%></td> <%
          if conIM then
            for j=1 to tamEscala %>
          <td class="tsmall" align="center"><%=j%></td> <%
            next %>
          <td class="tsmall" align="center" style="filter:;background:#<%=ListaColores(coboColorIM)%>; color:#<%=colorEtiqueta(coboColorIM)%>">Media</td>
          <td class="tsmall" align="center"><%=lblFRS_DesvStd_Abrev%></td>
          <td class="tsmall" align="center"><%=lblFRS_Moda%></td> <%
          end if %>
        </tr> <%
          totalfTot = 0
          redim totalfSA(tamEscala)
          redim totalfIM(tamEscala)
          totalModaIM = 0
          totalModaSA = 0
          for i=1 to numitems
            color = itemRep(i).semaforoRecomendacion(true)
            if itemRep(i).Id=0 then
              estilo = "celdaTit"
              clave = "G"
            else
              estilo = switchEstilo(estilo)
			  'David
              clave = numRomano(i)
            end if %>
        <tr class="<%=estilo%>">
          <td align="center"><%=clave%></td>
          <td class="tsmall"><%
            if numFactoresSelected=1 OR itemRep(i).Id=0 OR mov="PDF" then %>
            <%=itemRep(i).Descripcion%><%
            else %>
            <a href="#" onClick="setSeleccion('FACTOR',<%=itemRep(i).Id%>);"><%=itemRep(i).Descripcion%></a><%
            end if %>
          </td> <%
            totalfTot = totalfTot + itemRep(i).fTot
            for j=1 to tamEscala 
              totalfSA(j) = totalfSA(j) + itemRep(i).fSA(j)
              totalfIM(j) = totalfIM(j) + itemRep(i).fIM(j)%>
          <td class="tsmall" align="center" <%=iif(tipoGrafica<>"PIE" OR clave<>"G",""," style=""background-color:" & ListaColores(j-1) & ";text-color:" & colorEtiqueta(j-1) & """")%>><%=iif(muestraNumeros and sqResGral<>"",itemRep(i).fSA(j),coboFormatPorcentaje( div( itemRep(i).fSA(j), itemRep(i).fTot ) ))%></td> <%
            next %>
          <td align="center"><b><%=iif(muestraPorcentajeMedias and conMediaPercent,coboFormatPorcentaje(div( itemRep(i).SA, tamEscala)) ,coboFormatNum(itemRep(i).SA))%></b></td>
          <td align="center"><%=coboFormatNum(itemRep(i).SAdev)%></td>
          <td align="center"><%=itemRep(i).SAmoda%></td> <%
            if conIM then
              if itemRep(i).Escalas=COBO_SoloSA then %>
          <td colspan="<%=(4+tamEscala)%>" align="center"><i><%=lblFRS_N_A%></i></td> <%
              else
                for j=1 to tamEscala %>
          <td class="tsmall" align="center" <%=iif(tipoGrafica<>"PIE" OR clave<>"G",""," style=""background-color:" & ListaColores(j-1) & ";text-color:" & colorEtiqueta(j-1) & """")%>><%=iif(muestraNumeros and sqResGral<>"",itemRep(i).fIM(j),coboFormatPorcentaje( div( itemRep(i).fIM(j), itemRep(i).fTot ) ) )%></td> <%
                next %>
          <td align="center"><b><%=iif(muestraPorcentajeMedias and conMediaPercent,coboFormatPorcentaje(div( itemRep(i).IM, tamEscala)) ,coboFormatNum(itemRep(i).IM))%></b></td>
          <td align="center"><%=coboFormatNum(itemRep(i).IMdev)%></td>
          <td align="center"><%=itemRep(i).IMmoda%></td>
          <td class="tsmall" align="center" nowrap><%=coboFormatNum(itemRep(i).IM-itemRep(i).SA)%></td> <%
              end if %>
          <td class="tsmall" style="color:#<%=colorEtiquetaRGB(color)%>;background:#<%=color%>;"><%=itemRep(i).semaforoRecomendacion(false)%>&nbsp;</td> <%
            end if %>
        </tr> <%
          next 
          'COBO_muestraEstadisticasTodosFactores = true'En el Reporte por Factores muestra las estadísticas de todos los factores cuando es verdadero y hay más de un factor seleccionado.
          if COBO_muestraEstadisticasTodosFactores and sqResGral=""then%>
        <tr class="celdaTit" align="center"> 
          <td class="tsmall" align="center">-</td>
          <td class="tsmall" align="center"><%=lblFRS_Todos%></td><%
          AuxMediaSA = 0
          AuxMediaIM = 0
          for j=1 to tamEscala
            AuxMediaSA = AuxMediaSA + totalfSA(j)*j
            sumaTotTotSA = sumaTotTotSA + div(totalfSA(j),totalfTot)
            if totalModaSA <= totalfSA(j) then 
              modaRealSA = j
              totalModaSA =  totalfSA(j)
            end if%>
          <td class="tsmall" align="center"><%=iif(muestraNumeros and sqResGral<>"",totalfSA(j),coboFormatPorcentaje(div(totalfSA(j),totalfTot)))%></td> <%
          next 
          mediaRealSA = div(AuxMediaSA,totalfTot)
          AuxDesvSA = 0
          for j=1 to tamEscala
            AuxDesvSA = AuxDesvSA + totalfSA(j)*(j - mediaRealSA)*(j- mediaRealSA)
          next%>
          <td class="tsmall" align="center"><%=iif(conMediaPercent,coboFormatPorcentaje(div(mediaRealSA,tamEscala)),coboFormatNum(mediaRealSA))%></td>
          <td class="tsmall" align="center"><%=coboFormatNum(Sqr(div(AuxDesvSA,totalfTot-1)))%></td>
          <td class="tsmall" align="center"><%=modaRealSA%></td> <%
          if conIM then
            for j=1 to tamEscala 
              AuxMediaIM = AuxMediaIM + totalfIM(j)*j 
              sumaTotTotIM = sumaTotTotIM + div(totalfIM(j),totalfTot)
              if totalModaIM <= totalfIM(j) then 
                modaRealIM = j
                totalModaIM = totalfIM(j)
              end if%>
          <td class="tsmall" align="center"><%=iif(muestraNumeros and sqResGral<>"",totalfIM(j),coboFormatPorcentaje(div(totalfIM(j),totalfTot)))%></td> <%
            next 
            mediaRealIM = div(AuxMediaIM,totalfTot)
            AuxDesvIM = 0
            for j=1 to tamEscala
              AuxDesvIM = AuxDesvIM + totalfIM(j)*(j - mediaRealIM)*(j - mediaRealIM)
            next%>
          <td class="tsmall" align="center"><%=iif(conMediaPercent,coboFormatPorcentaje(div(mediaRealIM,tamEscala)),coboFormatNum(mediaRealIM))%></td>
          <td class="tsmall" align="center"><%=coboFormatNum(Sqr(div(AuxDesvIM,totalfTot-1)))%></td>
          <td class="tsmall" align="center"><%=modaRealIM%></td>
          <td class="tsmall" align="center"><%=coboFormatNum(mediaRealIM-mediaRealSA)%></td>
          <td class="tsmall" align="center">-</td> <%
          end if %>
        </tr>
      <%end if%>
      </TABLE>-->
	  
      <%END IF%>
<%IF mov<>"PDF" THEN%>
      <%ponLigaTop true%>
      <div align="center" nowrap>
        <%ponLigaRegreso(thispageid)%>
      </div>
      <input type="hidden" name="mov" value="">
      <input type="hidden" name="<%=reqbuscador%>" value="<%=serverHTMLencode(buscador)%>">
      <input type="hidden" name="tipoGrafica" value=<%=tipoGrafica%>>
    </form>
<%END IF%>
    </td>
  </tr>
</table>
</div>
<!--#include file="khorFooter.asp"-->
</BODY>
</HTML>
<%
  coboClean
  for i=1 to numitems
    set itemRep(i) = nothing
  next
  set bobj=nothing
'================================================================================'
  conn.close
  set conn=nothing
'================================================================================'
%>
