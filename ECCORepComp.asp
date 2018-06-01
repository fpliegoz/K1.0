<!--#include file="./khorClass.asp"-->
<!--#include file="./coboClass.asp"-->
<!--#include file="./coboAuxReportes.asp"-->
<!--#include file="./coboEspecial.asp"-->
<%
  Server.ScriptTimeout = 9999
  thispageid="ECCORepComp"
  thispage="ECCORepComp.asp"
  graficaDinamica = khorConfigValueWithDefault(597,true,0)<>0
  prom1 = reqn("Promedio1")
  prom2 = reqn("Promedio2")
  ubic1 = reqs("Ubicacion1")
  ubic2 = reqs("Ubicacion2")
  'Agrupador de variable de configuración
  ECCOAgrupador=ECCOAgrupadoresReporte()
  'VAriables sucias
  ECCOcuadrantes=4
  ECCOcuadranteOrigenX=4
  ECCOcuadranteOrigenY=3
  if coboRepComp_valorInicialGraficaEnUno then
    mensajeLimites=strLang( lblCOBO_ElValorDeLaUbicacionDebeSerEntre_X_y_, "1" )
    if ubic1 = "" then
    ubic1 = "1"
    end if
    if ubic2 = "" then
      ubic2 = "1"
    end if
  else
    mensajeLimites=strLang( lblCOBO_ElValorDeLaUbicacionDebeSerEntre_X_y_, "0" )
  end if
  
  
  '=== Inicializacion
  checaSesion ses_super&","&ses_adminid, "", ""
  modulosActivos = khorModulosActivos()
  validaEntrada khorPermisoModulo(Modulo_COBO,modulosActivos) OR khorPermisoModulo(Modulo_COBOConsulta,modulosActivos), "", thispageid
  errmsg=""
  mov=reqplain("mov")
  '=== Obtiene Empresa seleccionada o del contexto del usuario
  idf = reqn("idempresafiltro")
  if idf=0 then idf=khorEmpresaUsuario(0)
  '=== Obtiene parametros
  IdEncuesta = reqn("IdEncuesta")
  IdPersonal = request("IdPersonal")
  '=== Condiciones de periodo por eje
  IdPeriodo1 = reqn("IdPeriodo1")
  filtro1 = ""
  if IdPeriodo1<>0 then
    filtro1 = "origenParticipante.IdPeriodo=" & IdPeriodo1
  end if
  if coAnonima then
    IdPeriodo2 = IdPeriodo1
    filtro2 = filtro1
  else
    IdPeriodo2 = reqn("IdPeriodo2")
    filtro2 = ""
    if IdPeriodo2<>0 then
      filtro2 = "origenParticipante.IdPeriodo=" & IdPeriodo2
    end if
  end if
  if filtro1=filtro2 then
    filtroAuxParticipantes = filtro1
  else
    filtroAuxParticipantes = strAdd( filtro1, " OR ", filtro2 )
    if filtroAuxParticipantes<>"" then filtroAuxParticipantes = "(" & filtroAuxParticipantes & ")"
  end if
  filSucursal = coboFiltroSucursalPar(idf)
  if coAnonima AND filSucursal<>"" then
    filtroAuxParticipantes = strAdd( filtroAuxParticipantes, " AND ", "origenParticipante.IdSucursalPar IN (" & filSucursal & ")" )
  end if
  '=== Filtro de encuestas y validacion de la seleccionada
  filtroEncuestas = "EXISTS(SELECT * FROM co_Participante origenParticipante WHERE " & strAdd( filtroAuxParticipantes, " AND ", "Status=" & COBO_APLICADA & " AND origenParticipante.IdEncuesta=co_Encuesta.IdEncuesta)" )
  IdEncuesta = getBDnum("IdEncuesta","SELECT IdEncuesta FROM co_Encuesta WHERE IdEncuesta=" & IdEncuesta & " AND " & filtroEncuestas)
  '=== Inicializa buscador
  coboInicializaBuscador
  if IdEncuesta<>0 then
    filtroAuxParticipantes = strAdd( filtroAuxParticipantes, " AND ", "origenParticipante.IdEncuesta="&IdEncuesta )
  end if
  '=== Crea el filtro general de datos y la leyenda descriptiva con el numero de personas seleccionadas
  IF coAnonima THEN
    filtroListado = strAdd( filtroAuxParticipantes, " AND ", "origenParticipante.IdPersonal=0 AND Status=" & COBO_APLICADA )
    if IdPersonal="BUSCADOR" then
      filtroListado = strAdd( filtroListado, " AND ", bobj.condicion("origenParticipante.IdParticipante") )
      sq = "SELECT COUNT(*) AS cuantos FROM co_Participante origenParticipante WHERE " & strAdd( "Status=" & COBO_APLICADA, " AND ", filtroListado )
      strPersonal = getBDnum("cuantos",sq) & ": " & lblKHOR_UsandoCondicionesBusqueda
    elseif IdPersonal="*" then
      sq = "SELECT COUNT(*) AS cuantos FROM co_Participante origenParticipante WHERE " & filtroListado
      strPersonal = getBDnum("cuantos",sq) & ": " & lblCobo_TodasLasQueRespondieron
    else
      strPersonal = strLang( lblFRS_SeHa_nSeleccionado_N_X, "0||" & lblKHOR_Persona_s )
    end if
  ELSE
    '-- Filtro auxiliar de participantes con encuesta aplicada
    filtroAuxAplicados = "EXISTS(SELECT * FROM co_Participante origenParticipante WHERE " & strAdd( filtroAuxParticipantes, " AND ", "Status=" & COBO_APLICADA & " AND origenParticipante.IdPersonal=Personal.IdPersonal)" )
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
    filtroListado = strAdd( filtroAuxAplicados, " AND ", filtroListado )
  END IF
  filtroPeriodo = "EXISTS(SELECT * FROM co_Participante WHERE Status=" & COBO_APLICADA & " AND co_Participante.IdPeriodo=ed_periodo.IdPeriodo)"
  '=== Incializa menu auxiliar de Factores-Grupos para la seleccion. OJO: Los grupos se manejan con Id negativo
  conIM = false
  currGrupo = ""
  menuFijoFactores = ""
  sq = "SELECT DISTINCT Escalas, IdGrupo, IdFactor, Grupo, Factor, (CASE WHEN Factor=Grupo THEN Factor ELSE (Grupo" & db_concat &  "':'" & db_concat &  "Factor) END) AS Descripcion" & _
      " FROM vco_encuesta"
  if IdEncuesta <> 0 then
    sq = sq & " WHERE IdFactor IN (SELECT DISTINCT(IdFactor) FROM co_Reactivo WHERE IdReactivo IN (SELECT IdReactivo FROM co_EncuestaReactivo WHERE IdEncuesta = " & IdEncuesta & "))"
  end if
  sq = sq & " ORDER BY grupo, factor"
  set rs = getrs(conn,sq)
  while not rs.eof
    Grupo = rs("Grupo")
    if currGrupo <> Grupo AND Grupo <> rs("Factor") then
      menuFijoFactores = menuFijoFactores & (-1 * rsNum(rs,"IdGrupo")) & ":" & rs("Grupo") & "|"
      currGrupo = rs("Grupo")
    end if
    menuFijoFactores = menuFijoFactores & rsNum(rs,"IdFactor") & ":" & rs("Descripcion") & "|"
    if rsNum(rs,"Escalas")=COBO_Ambas then conIM = true
    rs.movenext
  wend
  rs.close
  set rs = nothing
  '=== Datos del reporte
  IdFactor1 = reqn("IdFactor1")
  IdFactor2 = reqn("IdFactor2")
  if conIM then
    Variable1 = request("Variable1")
    Variable2 = request("Variable2")
  else
    Variable1 = "SA"  'No traducir
    Variable2 = "SA"  'No traducir
  end if
  infoCompleta = (IdPeriodo1<>0) AND (IdPeriodo2<>0) AND _
                 (IdFactor1<>0) AND (IdFactor2<>0) AND _
                 (Variable1<>"") AND (Variable2<>"") AND _
                 IdPersonal<>""
  if infoCompleta AND IdPeriodo1=IdPeriodo2 AND IdFactor1=IdFactor2 AND Variable1=Variable2 then
    errmsg = lblCOBO_LosEjesSonIguales
    infoCompleta = false
  end if
  set regMapa = new coboItemMapa
  IF infoCompleta THEN
    regMapa.inicializa IdPeriodo1, IdPeriodo2, IdFactor1, IdFactor2, Variable1, Variable2, IdEncuesta, filtroListado, true
    fuenteExport = "./coboRepCompExp.asp?" & regMapa.qryString
    fuenteXML = "ECCOJSON.asp?" & regMapa.qryString
  END IF
  '--- Titulos
  titulo=lblCOBO_ClimaOrganizacional
  tit1=lblCOBO_RepComparativo
  tit2=regMapa.Titulo
  curtab=reqn("curtab")
  printerfriendly=(request("printerfriendly")<>"")
  childwin=(reqplain("childwin")=thispageid)

'================================================================================'
%>
<% 'Mostrar o no la version flash de mapa de talento
   'useSWFVersion = true 
%>
<HTML>
<HEAD>
<TITLE><%=khorAppName()%> - <%=titulo%> - <%=tit1%></TITLE>

<% if useSWFVersion then %>
<script type="text/javascript" src="./swfobject.js"></script>
<% else %>

<% end if %>
<% includeJS %>

<script src="./js/Chart.bundle.js"></script>
<script src="./js/utils.js"></script>
<script src="./js/chartjs-plugin-annotation.min.js"></script>
<script language="JavaScript">
<!--
<%IF infoCompleta THEN %>
var tabcount=2;
var tab;
function tabSel(t) {
  if (t!=tab) {
    for (var i=0; i<tabcount; i++) {
      setVisible('tabC'+i,i==t);
      ot=MM_findObj('tab'+i);
      ot.className=(i==t)?'tabOpen':'tabClose';
    }
    tab=t;
    op=MM_findObj('curtab');
    op.value=tab;
  }
}
<%if not printerfriendly then%>
function myPrintPage() {
  var whdl=window.open("","coboDetalle","toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=700,height=520");
  if (whdl) {
    document.TrueForm.target = 'coboDetalle';
    document.TrueForm.action="<%=thispage%>?printerfriendly=1&childwin=<%=thispageid%>";
    document.TrueForm.submit();
  }
  return false;
}
<%end if%>
function exportar() {
  location.href = '<%=replace(fuenteExport,"'","\'")%>';
}
<%END IF %>
function cambiaLista() {
  document.TrueForm.target = '';
  document.TrueForm.action="<%=thispage%>";<%
  if graficaDinamica THEN%>
  var ubi1 = document.getElementById("Ubicacion1").value;
  var ubi2 = document.getElementById("Ubicacion2").value;
  var limite = <%=coboEscalaLen()%>;
  if (ubi1 < <%=iif(coboRepComp_valorInicialGraficaEnUno,"1","0")%> | ubi1 > limite | ubi2 < <%=iif(coboRepComp_valorInicialGraficaEnUno,"1","0")%> | ubi2 > limite | isNaN(ubi1) | isNaN(ubi2))
  {
    alert("<%=mensajeLimites%>"+limite);
    return false;
  }
  else
    sendval('','mov','');<%
  ELSE%>
  sendval('','mov','');<%
  END IF%>
  
}
function cambiaPeriodo(idx) {
  var curr = MM_findObj('IdPeriodo'+idx);
  var otro = MM_findObj('IdPeriodo'+((idx==1)?2:1));
  if (curr.value != 0 && otro.value == 0) otro.value = curr.value;
  setFlashObj('btnRefresh');
}
<%
  coboBuscadorJS
  IF NOT coAnonima THEN %>
function seleccionaPersona() {
  var presel = getValor('IdPersonal');
  if (presel == '*' || presel == 'BUSCADOR') presel = '';
  abreSeleccion('PERSONA',true,presel,null,null,null,null,'<%=setSelectionFilter("PERSONA",filtroPersonas)%>');
} <%
  END IF %>
function setSeleccion(tipo,lista) {
  document.TrueForm.target = '';
  document.TrueForm.action="<%=thispage%>";<%
  if graficaDinamica THEN%>
  var ubi1 = document.getElementById("Ubicacion1").value;
  var ubi2 = document.getElementById("Ubicacion2").value;
  var limite = <%=coboEscalaLen()%>;
  if (ubi1 < 0 | ubi1 > limite | ubi2 < 0 | ubi2 > limite | isNaN(ubi1) | isNaN(ubi2))
  {
    alert("<%=mensajeLimites%>"+limite);
  }
  else
    sendval('','IdPersonal',lista);<%
  ELSE%>
  sendval('','IdPersonal',lista);<%
  END IF%>
}
function mapaDetalle (ID,Nombre,ValorX,ValorY) {
  stopLoad();
  alert(ID + "\n <%=titX%>:" + ValorX + ", <%=titY%>:" + ValorY);
}<%
  if graficaDinamica then %>
function cambiaHabilita(index)
{
  var check=document.getElementById("Promedio"+index).checked;
  var text = document.getElementById("Ubicacion"+index);
  var text = document.getElementById("Ubicacion"+index);
  text.disabled = check;
  text.value = <%=iif(coboRepComp_valorInicialGraficaEnUno,"1","0")%>;
}<%
  end if%>

//-->
</script>
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

    <form name="TrueForm" action="<%=thispage%>" method="POST" class="plana" onSubmit="return false;">
      <input type="hidden" name="childwin" value="<%=reqplain("childwin")%>">
      <INPUT type="hidden" name="curtab" value="<%=curtab%>">
      <input type="hidden" name="mov" value="">

      <table border=0 cellspacing=1 cellpadding=1 align="center">
        <%
          coboSeleccionSucursal idf, "cambiaLista()"" style=""font-size:90%;width:auto;""", "<tr><td align=""right"">" & lblKHOR_Sucursal & ":</td><td>", "</td></tr>"
        %>
        <%IF coAnonima THEN%>
        <tr>
          <td align="right"><%=lblKHOR_Periodo%>:</td>
          <td>
            <select name="IdPeriodo1" onChange="cambiaLista()" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
            <option value="0">- <%=lblFRS_Seleccione_ & lblKHOR_Periodo%> -</option>
            <% optionFromCatObj CAT_PERIODO, idf, IdPeriodo1, filtroPeriodo %>
            </select>
          </td>
        </tr>
        <%END IF%>
        <tr>
          <td align="right"><%=lblCOBO_Encuesta%>:</td>
          <td>
            <select name="IdEncuesta" onChange="cambiaLista()" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
            <option value="0">- <%=lblFRS_Cualquiera%> -</option>
            <% optionFromCat "co_Encuesta","IdEncuesta","Encuesta",false,filtroEncuestas,IdEncuesta %>
            </select>
          </td>
        </tr>
        <TR>
          <TD align="right"><%=iif(coAnonima,lblKHOR_Personas,lblKHOR_Persona_s)%>:</TD>
          <TD nowrap>
		  
            <INPUT  class="quitabordes" style="WIDTH:300px;" readOnly="true" name=strPersonal value="<%=strPersonal%>" > <%
            if not coAnonima then %>
            <INPUT type=button value="<%=lblFRS_Seleccionar%>" onclick="seleccionaPersona();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);> <%
            end if %>	
            <INPUT type=button value="<%=lblFRS_Buscar%>" onclick="callBuscador();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
            <INPUT type=button value="<%=lblFRS_Todas%>" onclick="setSeleccion('PERSONA','*');" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
		   <INPUT type=hidden name="IdPersonal" value="<%=IdPersonal%>">
          </TD>
        </TR> <%
        IF IdPersonal="BUSCADOR" THEN %>
        <TR>
          <TD><%=lblCOBO_Condiciones%>:</TD>
          <TD>
            <div class="bordeGris" style="margin:1px;">
              <%=bobj.explicacion("<BR>")%>
            </div>
          </TD>
        </TR> <%
        END IF %>
      </table>
      <table border=0 cellspacing=1 cellpadding=1 align="center">
        <TR class="celdaTit">
          <TD><%=lblFRS_Eje%></TD> <%
          if not coAnonima then %>
          <TD><%=lblKHOR_Periodo%></TD> <%
          end if %>
          <TD><%=lblKHOR_Grupo%>/<%=lblCOBO_Factor%></TD> <%
          if conIM then %>
          <TD><%=lblCobo_Variable%></TD> <%
          end if
          if graficaDinamica then %>
          <TD colspan="2"><%=lblCOBO_DivisionDeCuadrantes%></TD> <%
          end if %>
        </TR>
        <TR>
          <TD><%=lblFRS_Horizontal%></TD> <%
          if not coAnonima then %>
          <TD>
            <select name="IdPeriodo1" onChange="cambiaPeriodo(1)" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <option value="0">- <%=lblFRS_Seleccione_%> -</option>
              <% optionFromCatObj CAT_PERIODO, idf, IdPeriodo1, filtroPeriodo %>
            </select>
          </TD> <%
          end if %>
          <TD>
            <select name="IdFactor1" style="font-size:90%;width:auto;" onChange="setFlashObj('btnRefresh')" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <option value="0">- <%=lblFRS_Seleccione_%> -</option>
              <% optionFromMenuFijo menuFijoFactores, IdFactor1 %>
            </select>
          </TD> <%
          if conIM then %>
          <TD>
            <select name="Variable1" style="font-size:90%;width:auto;" onChange="setFlashObj('btnRefresh')" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <% optionFromMenuFijo menuFijoVariablesCO, Variable1 %>
            </select>
          </TD> <%
          else %>
          <input type="hidden" name="Variable1" value="<%=Variable1%>"> <%
          end if
          if graficaDinamica then %>
          <TD>
            <input type="checkbox" <%=iif(prom1=1,"checked","")%> onchange="setFlashObj('btnRefresh');cambiaHabilita('1');" name="Promedio1" value="1" id="Promedio1"><%=strJS(lblCOBO_UsarPromedio)%></input>
          </TD> 
          <TD><%=" &nbsp; &nbsp;"&lblFRS_Ubicacion&" &nbsp;"%>
            <input <%=iif (prom1=0,"","disabled")%> size="2" onChange="setFlashObj('btnRefresh')" type="textbox" name="Ubicacion1" value="<%=ubic1%>" id="Ubicacion1"></input>
          </TD><%
          end if %>
        </TR>
        <TR>
          <TD><%=lblFRS_Vertical%></TD> <%
          if not coAnonima then %>
          <TD>
            <select name="IdPeriodo2" onChange="cambiaPeriodo(2)" style="font-size:90%;width:auto;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <option value="0">- <%=lblFRS_Seleccione_%> -</option>
              <% optionFromCatObj CAT_PERIODO, idf, IdPeriodo2, filtroPeriodo %>
            </select>
          </TD> <%
          end if %>
          <TD>
            <select name="IdFactor2" style="font-size:90%;width:auto;" onChange="setFlashObj('btnRefresh')" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <option value="0">- <%=lblFRS_Seleccione_%> -</option>
              <% optionFromMenuFijo menuFijoFactores, IdFactor2 %>
            </select>
          </TD> <%
          if conIM then %>
          <TD>
            <select name="Variable2" style="font-size:90%;width:auto;" onChange="setFlashObj('btnRefresh')" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
              <% optionFromMenuFijo menuFijoVariablesCO, Variable2 %>
            </select>
          </TD> <%
          else %>
          <input type="hidden" name="Variable2" value="<%=Variable2%>"> <%
          end if
          if graficaDinamica then %>
          <TD>
            <input type="checkbox" <%=iif(prom2=1,"checked","")%> onchange="cambiaHabilita('2');setFlashObj('btnRefresh');" name="Promedio2" value="1" id="Promedio2"><%=strJS(lblCOBO_UsarPromedio)%></input>
          </TD> 
          <TD><%=" &nbsp; &nbsp;"&lblFRS_Ubicacion&" &nbsp;"%>
            <input <%=iif (prom2=0,"","disabled")%> size="2" onChange="setFlashObj('btnRefresh')" type="textbox" name="Ubicacion2" value="<%=ubic2%>" id="Ubicacion2"></input>
          </TD><%
          end if %>
        </TR>
        <TR>
          <TD colspan="<%=iif(coAnonima,2,3)+iif(conIM,1,0)+iif(graficaDinamica,2,0)%>" align="center">
            <INPUT type=button name="btnRefresh" value="<%=lblFRS_Procesar%>" onclick="cambiaLista();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
		</TD>
        </TR>
      </table> <%
      IF infoCompleta THEN
        IF NOT printerfriendly THEN %>
      <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
        <tr>
        <td>
          <table border="0" cellspacing="0" cellpadding="5" width="100%">
            <tr>
              <td id="tab0" nowrap onclick="tabSel(0);oculta(1);"><b><%=lblFRS_Grafica%></b></td>
              <td id="tab1" nowrap onclick="tabSel(1);oculta(2);"><b><%=lblFRS_Listado%></b></td>
              <td width="100%" class="tabFill">&nbsp;</td>
            </tr>
          </table>
        </td>
        </tr><tr valign="top">
        <td class="tabC"> <%
        ELSE %>
        <BR> <%
        END IF %>
          <!-- ==================== GRAFICA ==================== -->
		   <center><div id="container" >
				<canvas id="canvas" width="700px" height="650px"></canvas>
		   </div>  </center>  
              <%
                valorInicial = iif(coboRepComp_valorInicialGraficaEnUno,1,0)
                configStr = "&labelX=" & frsEncode(regMapa.Titulo1)
                configStr = configStr & "&labelY=" & frsEncode(regMapa.Titulo2)
                configStr = configStr & "&labelData=" & lblKHOR_Persona
                configStr = configStr & "&linearReg=0&quadrantBtn=1&quadrantFill=1"
                configStr = configStr & "&xStart=" & valorInicial & "&yStart=" & valorInicial
                configStr = configStr & "&xEnd=" & coboEscalaLen()
                configStr = configStr & "&yEnd=" & coboEscalaLen()
                configStr = configStr & "&xDiv=1&yDiv=1"
                configStr = configStr & "&xQuadDiv=" & iif(graficaDinamica and prom1=0,ubic1,coboFormatNum(regMapa.Promedio1))
                configStr = configStr & "&yQuadDiv=" & iif(graficaDinamica and prom2=0,ubic2,coboFormatNum(regMapa.Promedio2))
                fuenteXML = fuenteXML & configStr
              %>
			    <script>
       function oculta(paso){  
   if(paso!=1){
		$("#container").hide("slow");
		
   }else{
		$("#container").show("slow");
		
   }
}
        var bubbleChartData;
        
			
            
        window.onload = function() {
		     <%if mov = "PDF" then%>
				tabSel(1);
				oculta(2);
				$("#container").show("slow");
			<%end if%>
		     var jsonData= $.ajax({url:"<%=pageURL(1) & "/" & fuenteXML%>",
		                       dataType:'html',
							   }).done(function (results){
								  bubbleChartData=JSON.parse(results.toString());
								  console.debug(bubbleChartData);
								  var ctx = document.getElementById("canvas").getContext("2d");
								  window.myChart = new Chart(ctx, {
								  type: 'bubble',
                                                                 									
								  data: bubbleChartData,
								  
                                  options: {
                                      responsive: false,
                                      title:{
                                          display:true,
                                          text:'<%=lblECCO_TituloGraficaComprobacion%> '
                                      },
									  subtitle:{
										  displey:true,
										  text:'<%=lblECCO_TituloGraficaComprobacion%>'
									  },
										annotation: {
											drawTime: 'beforeDatasetsDraw',
											
											annotations: [{
												type: 'line',
												id:'line1',
												mode:'horizontal',
												scaleID:'y-axis-0',
												value:<%=coboFormatNum(regMapa.Promedio2)%>,
												
												borderColor: 'black',
												borderWidth: 1
											},{
												type: 'line',
												id:'line2',
												mode:'vertical',
												scaleID:'x-axis-0',
												value:<%=coboFormatNum(regMapa.Promedio1)%>,
												
												borderColor: 'black',
												borderWidth: 1
											}
											
											]
										},
                                      tooltips: {
                                          mode: 'nearest'
										},scales: {
										xAxes: [{										
											ticks: {
												min: 0,
												max: 5,
												stepSize: 1
											}
										}],
										yAxes: [{
											ticks: {
												min: 0,
												max: 5,
												stepSize: 1
											}
										}]
								        }
                                  }
                                  });
							   });
            
        };

       

       
    </script><!--
              <% if useSWFVersion then %>
              var so = new SWFObject("./swf/Scatter.swf", "Scatter", "690", "340", "10", "#FFFFFF");
              so.addParam("wmode", "transparent");
              so.addVariable("xmlUrl", escape("<%=pageURL(1) & "/" & fuenteXML%>"));
              so.write("flashCOBO");
              <% else %>
              $(function(){ startUp("<%=pageURL(1) & "/" & fuenteXML%>");});			    // ]]>      
              <% end if %>
              // ]]>
            </script>
            <% if not useSWFVersion then %>
            <table style="width:750px; position:relative;"><tr><td><div id="DivMapaTalento" style="height: 500px; width: 980px;"></div></td></tr></table>  
            <% end if %>
          </div>-->
          <!-- ==================== DATOS ==================== -->
          <div id="tabC1" style="display:<%=displayStyle(printerfriendly)%>;">
            <table class="tsmall" border="0" cellspacing="1" cellpadding="1" align="center">
              <tr class="celdaTit">
                <td colspan="4" align="center"><%=regMapa.Titulo%></td>
              </tr>
              <tr class="celdaTit">
                <td align="left"><%=lblKHOR_Persona%></td>
                <td align="center"><%=regMapa.Titulo1%></td>
                <td align="center"><%=regMapa.Titulo2%></td>
                <td align="center"><%=lblFRS_Cuadrante%></td>
              </tr>
              <tr class="celdaTit">
                <td align="right"><%=lblFRS_Promedios%>:</td>
                <td align="center"><%=coboFormatNum(regMapa.Promedio1)%></td>
                <td align="center"><%=coboFormatNum(regMapa.Promedio2)%></td>
                <td>&nbsp;</td>
              </tr> <%
              set rs = getrs( conn, regMapa.getQuery(0) )
              while not rs.EOF
                regMapa.getFromRS rs
                bgcuadrante = regMapa.cuadrante(true)
                cuadrante = regMapa.cuadrante(false)
                estilo = switchEstilo(estilo) %>
              <tr class="<%=estilo%>">
                <td><%=regMapa.Nombre%></td>
                <td align="center"><%=iif(regMapa.Dato1=0,"-",coboFormatNum(regMapa.Dato1))%></td>
                <td align="center"><%=iif(regMapa.Dato2=0,"-",coboFormatNum(regMapa.Dato2))%></td>
                <td align="center" bgcolor="<%=bgcuadrante%>"><%=cuadrante%></td>
              </tr> <%
                rs.movenext
              wend
              rs.close
              set rs = nothing %>
            </table>
          </div> <%
        IF NOT printerfriendly THEN %>
        </td>
        </tr>
      </table>
      <script language="JavaScript">
        tabSel(<%=curtab%>);
      </script> <%
        END IF
      END IF %>
      <%ponLigaTop true%>
      <div align="center" nowrap> <%
        IF infoCompleta THEN %>
        <INPUT type="button" value="<%=lblFRS_Exportar%>" onclick="exportar();" class="whitebtn" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this);" onmouseout="inOut(this);"> <%
        END IF %>
        <%ponLigaRegreso(thispageid)%>
      </div>
      <input type="hidden" name="<%=reqbuscador%>" value="<%=serverHTMLencode(buscador)%>">
    </form>

    </td>
  </tr>
</table>
<!--#include file="khorFooter.asp"-->
</BODY>
</HTML>
<%
  coboClean
  set regMapa = nothing
  set bobj=nothing
'================================================================================'
  conn.close
  set conn=nothing
'================================================================================'
%>