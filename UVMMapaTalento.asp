<!--#include file="./khorClass.asp"-->
<!--#include file="./ED_Class.asp"-->
<!--#include file="./e360class.asp"-->
<!--#include file="./UVMMapaTalentoClass.asp"-->
<%
  server.ScriptTimeout = 999999
  thispageid="UVMMapaTalento"
  thispage="UVMMapaTalento.asp"
  
  export = reqn("export") '0: No, 1: PDF, 2: XLS
  if export <> 0 then
    sesionFromRequest
    childwin = (1=1)
    printerfriendly = 1
    ocultarPersonasSinPlaza = (reqn("ocultarPersonasSinPlaza")=1)
  else
    childwin=(reqplain("childwin")=thispageid)
  end if
  checaSesion ses_super&","&ses_adminid, "", ""
  
  urlExpediente=khorExpedienteURL()
  conExpediente=khorPermisoExpediente(khorModulosActivos())
  curtab=reqn("curtab")
  '=== Filtros iniciales
  periodosPorEmpresa = khorCatalogoPorEmpresa(CAT_PERIODO)
  gruposPorEmpresa = khorEDporEmpresa() OR khor360porEmpresa() OR mciPorEmpresa()
  idf = reqn("idempresafiltro")
  if periodosPorEmpresa OR gruposPorEmpresa then
    if idf=0 then idf=khorEmpresaUsuario(0)
  else
    idf=0
  end if
  filtroPeriodos = khorSQLcatalogoEmpresa(CAT_PERIODO,idf)
  filtroPersonas = khorCondicionUsuario("Personal.IdPersonal")
  if (periodosPorEmpresa OR gruposPorEmpresa) AND idf<>0 then
    filtroPersonas = strAdd("Personal.IdSucursal="&idf, " AND ", filtroPersonas)
  end if
  '=== Obtiene periodo seleccionado y lo valida
  chgPeriodo = (ucase(reqplain("mov")) = "CHGPERIODO")
  sq = strAdd( "SELECT IdPeriodo FROM UVM_CatCiclos WHERE IdPeriodo=" & reqn("IdPeriodo"), " AND ", filtroPeriodos )
  IdPeriodo = getBDnum("IdPeriodo", sq )
  if IdPeriodo=0 then
    sq = strAdd( "SELECT IdPeriodo FROM UVM_CatCiclos", " WHERE ", filtroPeriodos ) & " ORDER BY IdPeriodo DESC"
    IdPeriodo = getBDnum("IdPeriodo", sq )
    chgPeriodo = true
  end if
  '=== Inicializa objeto de configuracion
  set mapa = new mapaTalentoClass
  mapa.Initalize khorModulosActivos(), NOT (khorIntegraModulosEnPeriodo() AND chgPeriodo), IdPeriodo, export = 1
  if export = 0 then
    validaEntrada khorPermisoModulo(Modulo_MapaTalento, khorModulosActivos()) AND (mapa.numItems(MT_PERFIL,false)>0) AND (mapa.numItems(MT_EJECUCION,false)>0), "", ""
  end if

  '=== Titulos
  tit = lblKHOR_MapaDeTalento
  tit1 = ""
  tit2 = ""
  titX = lblKHOR_MapaPerfil
  titY = lblKHOR_MapaEjecucion
 
  '-- Si se requieren cambiar estos valores, hacerlo en mapaTalentoEspecial.asp
  '-- division_perfil y division_ejecucion son los "tradicionales"
  '-- division_perfil2 y division_ejecucion2 son divisiones un nivel antes, para hacer un grid de 9, 
  '-- los valores de estas variables siempre deberán ser menores que su contraparte "tradicional".
  division_perfil = 70
  division_ejecucion = 70
  division_perfil2 = 0
  division_ejecucion2 = 0
  
  redim cuadranteTxt(9)
  cuadranteTxt(1) = "I"
  cuadranteTxt(2) = "II"
  cuadranteTxt(3) = "III"
  cuadranteTxt(4) = "IV"
  
  redim cuadranteTxtFull(9)
  
  redim cuadranteColor(9)
  cuadranteColor(1) = "#CC0000"
  cuadranteColor(2) = "#FFCC00"
  cuadranteColor(3) = "#00CC77"
  cuadranteColor(4) = "#0077CC"
  
  withInput = (khorConfigValue(468, true) = 1)
  maxValueX = khorConfigValueWithDefault(496, true, 100)
  maxValueY = maxValueX
  
  mapaTalento_NumColAdListado = 0
%>
<!--#include file="./mapaTalentoEspecial.asp"-->
<%
  '-- Limites (capturables)
  dim limitX1 : limitX1 = reqn("limitX1")
  dim limitX2 : limitX2 = reqn("limitX2")
  dim limitY1 : limitY1 = reqn("limitY1")
  dim limitY2 : limitY2 = reqn("limitY2")
  mapaVerificaLimites '-- Asigna defaults si los capturables vienen en ceros
  
  'mapaTalento_NumColAdListado = 1  '-- Considera este numero de columnas adicionales en el listado
  '-- Para ello, deben existir en mapaTalentoEspecial.asp dos funciones:
  '-- mapaTalento_ColAdTitle(i) que regrese el titulo de la columna i (i va desde 1 a mapaTalento_NumColAdListado)
  '-- mapaTalento_ColAdValue(i,rs) que regrese el valor de la columna i para el renglon actual del recordset del listado
  '-- mapaTalento_ColAdTotal(i,useTotal2) que regrese el valor de la columna i para el renglon de promedios
  
  'Para mostrar el campo que se usa para el Login como una columna más en la tabla del listado.
  'mapaTalento_MostrarCampoLoginEnListado = true
  if mapaTalento_MostrarCampoLoginEnListado then
    tipoLogin = khorTipoLogin()
    lblUsr = iif( tipoLogin=1, lblKHOR_ClaveDePersona, khorLoginLabel(tipoLogin, khorDefaultPais()))
  end if
  'Mostrar o no la version flash de mapa de talento
  'useSWFVersion = true 

  '=== Obtiene personas seleccionadas y valida que hayan sido evaluadas en el periodo y que sean de la empresa
  lstFromDB = reqs("lstFromDB")
  if lstFromDB <> "" then
    IdPersonal = getBD( "Value", "SELECT Value FROM mapaTalentoTemp WHERE IdTemp = '" & sqsf(lstFromDB) & "'" )
  else
    IdPersonal = request("IdPersonal")
  end if
  IF IdPeriodo<>0 THEN
    '=== Valida que las personas hayan sido evaluadas en el periodo y que sean de la empresa
    'filtroPersonas = strAdd(filtroPersonas, " AND ", mapa.condicionPersona(idperiodo))
    filtroPersonas = strAdd(filtroPersonas, " AND ", "(EXISTS(SELECT * FROM UVM_EvaluacionDimension e INNER JOIN UVM_TablaCRNS c ON c.IdCRN = e.IdCRN "&_
		"INNER JOIN Personal p ON CONVERT ( INT , REPLACE(p.Clave ,p.Prefijo ,'')) =c.Docente AND p.Clave = Personal.Clave "&_
	   " WHERE c.IdPeriodos="&idperiodo&"))")
    if IdPersonal="*" then
      strSel = lblKHOR_TodasLasEvaluadasEnElPeriodo
    else
      if IdPersonal<>"" then
        sq = "SELECT IdPersonal FROM Personal WHERE IdPersonal IN (" & IdPersonal & ") AND " & filtroPersonas
        IdPersonal = getBDlist("IdPersonal",sq,false)
      end if
      strSel = strLang( lblFRS_SeHa_nSeleccionado_N_X, lenCSV(IdPersonal) & "||" & lblKHOR_Persona_s)
    end if
    '=== Condicion general del listado
    if IdPersonal="*" then
      filtroListado = filtroPersonas
    else
      filtroListado = "Personal.IdPersonal IN (" & IdPersonal & ")"
    end if
  ELSE
    IdPersonal = ""
  END IF

  '=== Obtiene perfil para compatibilidad y recalcula
  IdPerfil = reqn("IdPerfil")
  groupBy = reqn("groupBy")
  hideDetail = reqn("hideDetail")
  if IdPerfil <> 0 then groupBy = 0
  if groupBy = 0 then hideDetail = 0
  
  '=== FUNCIONES AUXILIARES
  
  sub recalcula()
    response.flush
    '=== Calcula compatibilidad
    khorCalculaCompatibilidadMasivo IdPerfil, filtroPersonas
    '=== Verifica resultados 360
    if mapa.peso(MT_360)>0 then
      sq = "SELECT DISTINCT IdGrupo FROM rep360 WHERE IdPeriodo=" & IdPeriodo & " AND (Resultado_360 IS NULL OR Resultado_360=0) AND IdPersonal IN ("
      if IdPersonal="*" then
        sq = sq & strAdd("SELECT IdPersonal FROM Personal"," WHERE ",khorCondicionUsuario("Personal.IdPersonal"))
      else
        sq = sq & IdPersonal
      end if
      sq = sq & ")"
      set rs = getrs(conn,sq)
      while not rs.eof
        calculaResultados360 rsNum(rs,"IdGrupo")
        rs.movenext
      wend
      rs.close
      set rs = nothing
    end if
  end sub
  
  sub printPromedios(m, numper, colspan, useTotal2, asField)
    dim eje, i, strOut
    strOut = "<tr class='" & iif(asField, "celdaLight", "celdaTit") & "'><td align='right' colspan='" & colspan & "'>" & lblFRS_Promedio & ":</td>"
    for eje=1 to 2
      for i=1 to m.mtNum
        if m.mtArr(i).tipo = eje and m.mtArr(i).peso>0 then
          if m.mtArr(i).itemCols > 1 then
            strOut = strOut & "<td align='center'>" & khorFormatPorcentaje(iif(useTotal2, m.mtArr(i).auxtotal2, m.mtArr(i).auxtotal) / numper) & "</td>"
            m.mtArr(i).auxtotal = 0
          end if
          strOut = strOut & "<td align='center'>" & khorFormatPorcentaje(iif(useTotal2, m.mtArr(i).total2, m.mtArr(i).total) / numper) & "</td>"
          m.mtArr(i).total = 0
        end if
      next
      strOut = strOut & "<td align='center'>" & khorFormatPorcentaje((iif(useTotal2, sumavalor2(eje), sumavalor(eje)) / numper) / 100) & "</td>"
      sumavalor(eje) = 0
    next
    strOut = strOut & "<td>&nbsp;</td>"
    for i=1 to mapaTalento_NumColAdListado
      strOut = strOut & "<td>" & mapaTalento_ColAdTotal(i,useTotal2) & "</td>"
    next
    strOut = strOut & "</tr>"
    response.write strOut
  end sub
  
  function mapaVerificaLimites()
    '-- default si los capturables vienen en cero
    if limitX1 = 0 then
      limitX1 = division_perfil
    end if
    if limitY1 = 0 then
      limitY1 = division_ejecucion
    end if
    if limitX2 = 0 and division_perfil2 > 0 then
      limitX2 = division_perfil2
    end if
    if limitY2 = 0 and division_ejecucion2 > 0 then
      limitY2 = division_ejecucion2
    end if
  end function
  
  function mapaIndexCuadrante(vper,veje)
    dim idxCuadrante
    '-- Orden de los cuadrantes en Scatter.swf (de izquierda a derecha y de abajo hacia arriba)
    if division_ejecucion2 = 0 then
      '-- Mapa original: un corte en cada eje = 4 cuadrantes: I, II, III, IV / rojo, amarillo, verde, azul
      if veje>=limitY1 then
        idxCuadrante = iif( vper >= limitX1, 4, 3 )
      else
        idxCuadrante = iif( vper >= limitX1, 2, 1 )
      end if
    else
      '-- Mapa 9-block: dos cortes en cada eje = 9 bloques
      if veje>=limitY1 then
        idxCuadrante = iif( vper >= limitX1, 9, iif( vper >= limitX2, 8, 7 ) )
      elseif veje>=limitY2 then
        idxCuadrante = iif( vper >= limitX1, 6, iif( vper >= limitX2, 5, 4 ) )
      else
        idxCuadrante = iif( vper >= limitX1, 3, iif( vper >= limitX2, 2, 1 ) )
      end if
    end if
    mapaIndexCuadrante = idxCuadrante
  end function

'========================================
if export = 2 then 'XLS
  response.ContentType = "application/vnd.ms-excel"
  response.AddHeader "Content-Disposition", "attachment;filename=mapaTalento" & Year(Now()) & Month(Now()) & Day(Now()) & ".xls"
end if
layoutHeadStart khorAppName()& " - " & tit & " - " & tit1
'========================================
IF IdPeriodo<>0 AND IdPersonal<>"" THEN
  recalcula
END IF
IF export = 0 THEN
  includeJS
  if useSWFVersion then %>
<script type="text/javascript" src="./swfobject.js"></script> <% 
  else %>
<script type="text/javascript" src="./js/raphael.js"></script>
<script type="text/javascript" src="./js/raphaeljs-infobox.js"></script>
<script type="text/javascript" src="./Scatter.js.asp"></script> <% 
  end if %>
<SCRIPT LANGUAGE="javascript">
<!--
var tabcount=3;
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
    $("#descDiv").html("");
  }
  if(t == 0) {
    $("#DivMapaTalento").show();
  }else{
    $("#DivMapaTalento").hide();
  }
} <%
  IF IdPeriodo<>0 AND IdPersonal<>"" THEN
    mapa.declareJSdetalle %>
  function mapaDetalle(ID,Nombre,ValorX,ValorY) {
    stopLoad();
    alert(Nombre + "\n <%=titX%>:" + ValorX + ", <%=titY%>:" + ValorY);
  }
  function mapaEjePeso(eje) {
    var total=0; <%
    for eje=1 to 2
      auxs = ""
      for i=1 to mapa.mtNum
        if mapa.mtArr(i).tipo = eje then
          auxs = strAdd( auxs, " + ", "getValor('peso_" & mapa.mtArr(i).id & "','int')" )
        end if
      next %>
    if (eje == <%=eje%>) {
      total = <%=auxs%>;
    } <%
    next %>
    return total;
  }
  function mapaSetTotal(eje) {
    var total=mapaEjePeso(eje);
    MM_setTextOfLayer('totEje'+eje,0,total+'%');
    var obj = MM_findObj('totEje'+eje);
    obj.style.color = ((total==100)?document.body.style.color:'#FF0000');
  }
  function myPrintPage(format) {
    <% 
      url = thispage & "?mov=PDF&IdPeriodo=" & IdPeriodo & "&IdPerfil=" & IdPerfil & "&idempresafiltro=" & idf &_
            mapa.qryString() & "&groupBy=" & groupBy & "&hideDetail=" & hideDetail & "&limitX1=" & limitX1 & "&limitY1=" & limitY1 & "&limitX2=" &_
            limitX2 & "&limitY2=" & limitY2 & "&ocultarPersonasSinPlaza=" & iif(ocultarPersonasSinPlaza,1,0)
      lstFromDB = "mt" & adminSesion() & "_" & formatDateAMD(Now) & getIntLen(Hour(Now),2) & getIntLen(Minute(Now),2) & getIntLen(Second(Now),2)
      conn.execute "DELETE FROM mapaTalentoTemp WHERE IdTemp LIKE 'mt" & adminSesion() & "_%' AND IdTemp < '" & lstFromDB & "'"
      if len(IdPersonal) > 4096 then
        conn.execute "INSERT INTO mapaTalentoTemp(IdTemp, Value) VALUES('" & sqsf(lstFromDB) & "', '" & sqsf(IdPersonal) & "')"
        url = url & "&lstFromDB=" & lstFromDB
      else
        url = url & "&IdPersonal=" & IdPersonal
        lstFromDB = ""
      end if
    %>
    if(format == undefined) format = "PDF";
    if(format == "PDF") { <%
      pdfkey = initPDFurl( thispageid, pdf_URL() & url ) %>
      openPDFjob('<%=pdfkey%>', 'export', 1);
    } else if(format == "XLS") {
      noPageBlocker = true;
      sendval('', 'export', 2);
    }
  } <%
  END IF %>
function cambiaLista(chgPeriodo){
  if (chgPeriodo) setValor('mov','CHGPERIODO');
  sendval('', '_target','', 'export','', '_action','<%=thispage%>');
}
function setGroupBy(gb) {
  sendval('', 'groupBy', gb);
}
function seleccionaPersona() {
  var presel = getValor('IdPersonal');
  if (presel == '*') presel = '';
  abreSeleccion('PERSONA',true,presel,'<%=lblKHOR_PersonaEvaluadaEnElPeriodo%>',null,null,null,'<%=setSelectionFilter("PERSONA",filtroPersonas)%>');
}
function setSeleccion(tipo,lista) {
  setValor('IdPersonal',lista);
  cambiaLista();
}
function toggleDetail(detail) {
  sendval('', 'hideDetail', detail == 1 ? 0 : 1);
}
function printDivTxt(q) {
  $("#descDiv").html(cuadranteData[q - 1].txt + ". " + cuadranteData[q - 1].txtFull);
  $("#descDiv").css("background-color", cuadranteData[q - 1].color);
  if(cuadranteData[q - 1].color.match(/F/g).length > 2)
  {
    $("#descDiv").css("color", "#000000");
  }
  else
  {
    $("#descDiv").css("color", "#FFFFFF");
  }
}
function setLimit(kind) {
  if($("#limit" + kind + "2").val() > $("#limit" + kind + "1").val())
  {
    $("#limit" + kind + "2").val(1);
    if($("#limit" + kind + "2").val() == $("#limit" + kind + "1").val()) {
      $("#limit" + kind + "1").val(2);
    }
  }
  sendval('');
}
  var cuadranteData = new Array(); <%
  for i = 1 to UBound(cuadranteTxtFull)
    response.write vbCRLF & "cuadranteData.push({txtFull: '" & cuadranteTxtFull(i) & "', color: '" & cuadranteColor(i) & "', txt: '" & cuadranteTxt(i) & "'});"
  next %>
//-->
</SCRIPT> <% 
END IF
'========================================
layoutHeadEnd
if export = 2 then
  response.write vbCRLF & "<body>"
else
  layoutStart tit, tit1, tit2, errmsg, khorWinWidth(), ""
  defaultFormStart thispage, "", true %>
      <INPUT type="hidden" name="IdGrupo" value="">
      <INPUT type="hidden" name="curtab" value="<%=curtab%>">
      <input type="hidden" name="childwin" value="<%=reqplain("childwin")%>">
      <INPUT type="hidden" name="mov" value="">
      <input type="hidden" name="groupBy" value="<%=groupBy%>">
      <input type="hidden" name="hideDetail" value="<%=hideDetail%>">
      <input type="hidden" name="export" value="0"> <% 
end if

    '==================== FILTROS / PARAMETROS ====================
    if export = 0 then %>
      <TABLE cellSpacing=1 cellPadding=1 border=0 align="center">
        <%
          khorSeleccionSucursal idf, (periodosPorEmpresa OR gruposPorEmpresa), true, "cambiaLista()", "<TR><TD>" & lblKHOR_Sucursal & ":</TD><TD>", "</TD></TR>"
        %>
        <TR>
          <TD><%=lblKHOR_MapaPeriodo%>:</TD>
          <TD colspan="2">
            <select name="IdPeriodo" id="IdPeriodo" onChange="cambiaLista(true)"  class="sinbordes" style="width:auto;"><%
              set rsPeriodo =  getrs(conn, "SELECT * FROM UVM_CatCiclos ORDER BY IdPeriodo DESC")
              if rsPeriodo.EOF then%>
              <option value='0' SELECTED>No hay periodos registrados.</option><%
              else%>
              <option value="0">- Seleccione Periodo -</option><%
                while not rsPeriodo.EOF%>
              <option <%=iif(IdPeriodo=rsNum(rsPeriodo,"IdPeriodo"),"selected=''","")%> value='<%=rsNum(rsPeriodo,"IdPeriodo")%>'><%=rsNum(rsPeriodo,"IdPeriodo")%></option><%
                  rsPeriodo.movenext
                wend
              end if%>
            </select>
          </td>
        </TR> <% 
        IF IdPeriodo<>0 THEN %>
        <TR>
          <TD><%=lblKHOR_Persona_s%>:</TD>
          <TD>
            <INPUT  class="quitabordes" style="WIDTH:250px" readOnly="true" name=strSel value="<%=strSel%>" >
          </TD>
          <TD>
            <INPUT type=hidden name="IdPersonal" value="<%=IdPersonal%>">
            <INPUT type=button value="<%=lblFRS_Seleccionar%>" onclick="seleccionaPersona();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
            <INPUT type=button value="<%=lblFRS_Todas%>" onclick="setSeleccion('','*');" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
          </TD>
        </TR>
        <TR>
          <TD><%=lblKHOR_PerfilDeCompatibilidad%>:</TD>
          <TD colspan="2">
            <select name="IdPerfil" id="IdPerfil" onChange="cambiaLista()"  class="sinbordes" style="width:100%">
              <option value="0">- <%=lblKHOR_PerfilDeCadaPersona%> -</option>
              <% optionFromCat "vPerfil", "IdPerfil", "Perfil", false, khorSQLperfilesEmpresa(idf,-1), IdPerfil %>
            </select>
          </TD>
        </TR>
        <TR>
          <TD><%=lblLMS_GroupBy%>:</TD>
          <TD>
            <input name="groupByRadio" onChange="setGroupBy(0);" type="radio" value="0" <%=checkedif(groupBy = 0)%>><%=lblKHOR_Persona%></input>
            <input name="groupByRadio" onChange="setGroupBy(1);" type="radio" value="1" <%=disabledif(IdPerfil <> 0)%><%=checkedif(groupBy = 1)%>><%=lblKHOR_Perfil%></input>
            &nbsp;&nbsp;&nbsp;
            <input onChange="toggleDetail(<%=hideDetail%>);" type="checkbox" <%=disabledif(groupBy = 0)%><%=checkedif(hideDetail = 1)%>><%=lblKHOR_OcultarDatosDePersonas%></input>
          </TD>
        </TR> <% 
          if withInput then %>
        <TR>
          <TD><%=lblKHOR_LimitesCuadrantes%>:</TD>
          <TD>
            <table width="100%">
              <tr>
                <td width="60%"><%=titX%>:</td>
                <% if division_perfil2 > 0 then %>
                <td><input type="text" maxlength="3" onBlur="valida(this, 'int', 1, 99); setLimit('X');" name="limitX2" id="limitX2" style="width: 30px;" value="<%=limitX2%>" /></td>
                <% else %>
                <td><input type="hidden" value="0" name="limitX2" id="limitX2" /></td>
                <% end if %>
                <td><input type="text" maxlength="3" onBlur="valida(this, 'int', 1, 100); setLimit('X');" name="limitX1" id="limitX1" style="width: 30px;" value="<%=limitX1%>" /></td>
              </tr>
              <tr>
                <td width="60%"><%=titY%>:</td>
                <% if division_ejecucion2 > 0 then %>
                <td><input type="text" maxlength="3" onBlur="valida(this, 'int', 1, 99); setLimit('Y');" name="limitY2" id="limitY2" style="width: 30px;" value="<%=limitY2%>" /></td>
                <% else %>
                <td><input type="hidden" value="0" name="limitY2" id="limitY2" /></td>
                <% end if %>
                <td><input type="text" maxlength="3" onBlur="valida(this, 'int', 1, 100); setLimit('Y');" name="limitY1" id="limitY1" style="width: 30px;" value="<%=limitY1%>" /></td>
              </tr>
            </table>
          </TD>
        </TR> <% 
          end if
        END IF %>
      </TABLE> <% 
    end if %>
      <%
    IF IdPeriodo<>0 AND IdPersonal<>"" THEN
      '==================== TABS ====================
      IF NOT printerfriendly THEN %>
      <br>
      <table border=0 cellspacing=0 cellpadding=0 align=center width=100%> <% 
        if export = 0 then %>
        <tr>
          <td>
            <table border=0 cellspacing=0 cellpadding=5 width=100%>
              <tr>
                <td id=tab0 nowrap onclick="tabSel(0);"><b><%=lblFRS_Grafica%></b></td>
                <td id=tab1 nowrap onclick="tabSel(1);"><b><%=lblFRS_Listado%></b></td>
                <td id=tab2 nowrap onclick="tabSel(2);"><b><%=lblFRS_Pesos%></b></td>
                <td width="100%" class="tabFill" align="right">
                  <%=lblFRS_Exportar%>: <a href="javascript:void(0);" onClick="myPrintPage('PDF')"><%=lblLMS_PDF%>&nbsp;<img src="./khorImg/lms/icons/1.gif" border="0" /></a>
                  <a href="javascript:void(0);" onClick="myPrintPage('XLS')"><%=lblLMS_Excel%>&nbsp;<img src="./khorImg/lms/icons/5.gif" border="0" /></a>
                </td>
              </tr>
            </table>
          </td>
        </tr> <% 
        end if %>
        <tr valign="top">
          <td class="tabC"> <%
      END IF
      '==================== GRAFICA ====================
      fuenteXML = pageURL(1) & "/UVMMapaTalentoXML.asp"
      paramsXML = "?IdEmpresaFiltro=" & idf & "&IdPeriodo=" & IdPeriodo & "&IdPerfil=" & IdPerfil & mapa.qryString()
      quadDivX = limitX1
      if limitX2 <> 0 then
        quadDivX = limitX2 & "," & quadDivX
      end if
      quadDivY = limitY1
      if limitY2 <> 0 then
        quadDivY = limitY2 & "," & quadDivY
      end if
      paramsXML = paramsXML & "&fieldLbl=Nombre&labelX=" & FrsEncode(titX) & "&labelY=" & FrsEncode(titY) & "&labelData=" & lblKHOR_Persona & "&linearReg=0&quadrantBtn=1&quadrantFill=1&withStats=1&xStart=0&yStart=0&xEnd=" & maxValueX & "&yEnd=" & maxValueY & "&xDiv=10&yDiv=10&xQuadDiv=" & quadDivX & "&yQuadDiv=" & quadDivY & "&overrideColors=" & replace(getCSVFromArray(cuadranteColor), "#", "") & "&overrideText=" & getCSVFromArray(cuadranteTxt)
      if lstFromDB <> "" then
        paramsXML = paramsXML & "&lstFromDB=" & lstFromDB
      else
        paramsXML = paramsXML & "&IdPersonal=" & IdPersonal
      end if
      if export = 0 then %>
          <div id=tabC0 style="display:<%=displayStyle(printerfriendly)%>;" align="center"> <% 
            if useSWFVersion then %>
            <div id="flashmapatalento">
              <strong><%=lblFRS_FlashNecesitaActualizarse%>
            </div>
            <div id="descDiv" style="width: 600px;">&nbsp;</div>
            <script type="text/javascript">
              // <![CDATA[
              var so = new SWFObject('./swf/Scatter.swf', 'Scatter', '690', '340', '10', '#ffffff');
              so.addParam('wmode', 'transparent');
              so.addVariable("lblLoading", "<%=lblFRS_Espere%>");
              so.addVariable("lblError", "<%=lblFRS_Error%>");
              so.addVariable("xmlUrl", "<%=escape(fuenteXML & paramsXML)%>");
              so.write("flashmapatalento");
              // ]]>
            </script> <% 
            end if %>
          </div> <% 
            if useSWFVersion = 0 then %>
          <script type="text/javascript">
            $(function(){ startUp("<%="./UVMMapaTalentoXML.asp" & paramsXML%>");});			    // ]]>
          </script>
          <table style="width:750px; position:relative;"><tr><td><div id="DivMapaTalento" style="height: 500px; width: 980px;" ></div></td></tr></table> <% 
            end if
      elseif export = 1 then
        paramsASHX = getValueString("ashx", IdPeriodo, IdPersonal, IdPerfil, 0, idf)
        lstFromDB = "mtx" & adminSesion() & "_" & formatDateAMD(Now) & getIntLen(Hour(Now),2) & getIntLen(Minute(Now),2) & getIntLen(Second(Now),2)
        conn.execute "DELETE FROM mapaTalentoTemp WHERE IdTemp LIKE 'mtx" & adminSesion() & "_%' AND IdTemp < '" & lstFromDB & "'"
        if len(paramsASHX) > 4096 then
          conn.execute "INSERT INTO mapaTalentoTemp(IdTemp, Value) VALUES('" & sqsf(lstFromDB) & "', '" & sqsf(paramsASHX) & "')"
          paramsASHX = "&lstFromDB=" & lstFromDB
        end if %>
          <div id="grafashx" align="center">
            <table width="100%">
              <tr>
                <td width="100%" align="center">
                  <img src="<%=graficaURL()%>charts/Scatter.ashx?xStart=0&xEnd=<%=maxValueX%>&xDivision=10&yStart=0&yEnd=<%=maxValueY%>&yDivision=10&xQuadrantDiv=<%=quadDivX%>&yQuadrantDiv=<%=quadDivY%>&labelX=<%=FrsEncode(titX)%>&labelY=<%=FrsEncode(titY)%>&quadrantFill=1&linearReg=1&values=<%=paramsASHX%>&colorOverride=<%=replace(getCSVFromArray(cuadranteColor), "#", "")%>">
                </td>
              </tr> <%
                if (cuadranteTxtFull(1) <> "") AND (division_ejecucion2 > 0) AND (division_perfil2 > 0) then
                  strOut = "<tr><td width='100%' align='center'><table width='100%'>"
                  for i = 1 to 9
                    tmpColor = iif((len(cuadranteColor(i)) - len(replace(cuadranteColor(i), "F", ""))) > 1, "#000000", "#FFFFFF")
                    tmpColor = "#000000"
                    if cuadranteColor(i) = "#FF0000" then tmpColor = "#FFFFFF"
                    tmpVar = "<td align='center' style='filter:;background: " & cuadranteColor(i) & "; color: " & tmpColor & "'>" & cuadranteTxt(i) & ": " & cuadranteTxtFull(i) & "</td>"
                    if i >= 7 then
                      strOut79 = strOut79 & tmpVar
                    elseif i >= 4 then
                      strOut46 = strOut46 & tmpVar
                    else
                      strOut13 = strOut13 & tmpVar
                    end if
                  next
                  strOut = strOut & "<tr>" & strOut79 & "</tr>"
                  strOut = strOut & "<tr>" & strOut46 & "</tr>"
                  strOut = strOut & "<tr>" & strOut13 & "</tr>"
                  strOut = strOut & "</table></td></tr>"
                  response.write strOut
                end if %>
            </table>
          </div> <% 
      end if
      '==================== LISTADO ====================
      numColsIniciales = bool2num(mapaTalento_MostrarCampoLoginEnListado) + 1 + bool2num(IdPerfil=0 and groupBy=0)
      dim ejenumcol(2)
      for eje=1 to 2
        ejenumcol(eje) = mapa.numItems(eje,true)
        if ejenumcol(eje)>1 then ejenumcol(eje) = ejenumcol(eje)+1
      next
      numColsEstandar =  numColsIniciales + ejenumcol(1) + ejenumcol(2) %>
          <div id=tabC1 style="display:<%=displayStyle(printerfriendly)%>;">
            <table cellSpacing="1" cellPadding="1" border="0" align="center" class="tsmall">
              <tr class="celdaTit"><%
                if mapaTalento_MostrarCampoLoginEnListado then%>
                <td rowspan="2" align="center"><%=lblUsr%></td> <%
                end if%>
                <td rowspan="2" align="center"><%=iif(hideDetail = 0, lblKHOR_Persona, "")%></td> <%
                if IdPerfil=0 and groupBy=0 then %>
                <td rowspan="2" align="center"><%=lblKHOR_Perfil%></td><%
                end if
                for eje=1 to 2 %>
                <td colspan="<%=ejenumcol(eje)%>" align="center"><%=mapaAxisName(eje)%></td> <%
                next %>
                <td rowspan="3" align="center"><%=lblFRS_Cuadrante%></td> <%
                for i=1 to mapaTalento_NumColAdListado
                  response.write vbCRLF & "<td rowspan=""3"">" & mapaTalento_ColAdTitle(i) & "</td>"
                next %>
              </tr>
              <tr class="celdaTit"> <%
                for eje=1 to 2
                  for i=1 to mapa.mtNum
                    if mapa.mtArr(i).tipo=eje AND mapa.mtArr(i).peso>0 then %>
                <td align="center" <%=iif(mapa.mtArr(i).itemCols>1," colspan=""" & mapa.mtArr(i).itemCols & """","")%>><%=mapa.mtArr(i).desc%></td> <%
                    end if
                  next
                  if ejenumcol(eje)>1 then %>
                <td align="center"><%=mapaAxisName(eje)%></td> <%
                  end if
                next %>
              </tr>
              <tr class="celdaTit">
                <td colspan="<%=numColsIniciales%>" align="right"><%=lblFRS_Pesos%>:</td> <%
                for eje=1 to 2
                  for i=1 to mapa.mtNum
                    if mapa.mtArr(i).tipo=eje AND mapa.mtArr(i).peso>0 then %>
                <td align="center" <%=iif(mapa.mtArr(i).itemCols>1," colspan=""" & mapa.mtArr(i).itemCols & """","")%>><%=mapa.mtArr(i).peso%>%</td> <%
                    end if
                  next
                  if ejenumcol(eje)>1 then %>
                <td align="center"><%=mapa.SumaPesos(eje)%>%</td> <%
                  end if
                next %>
              </tr> <%
                dim ejevalor(2)
                dim sumavalor(2)
                dim sumavalor2(2)
                dim strOut
                dim numper : numper = 0
                dim numper2 : numper2 = 0
                dim lastIdPuesto : lastIdPuesto = 0
                dim printTitle
                sq = mapa.query(idperfil,idperiodo,filtroListado,groupBy)
                set rsMapa = getrs(conn,sq)
                while not rsMapa.EOF
                  if IdPerfil=0 then
                    idp = rsNum(rsMapa,"IdPuesto")
                    if lastIdPuesto <> idp then
                      lastIdPuesto = idp
                      printTitle = true
                    else
                      printTitle = false
                    end if
                    puesto = rsStr(rsMapa,"Puesto")
                  else
                    idp = IdPerfil
                  end if
                  valorPerfil = mapa.getValorRS(rsMapa,MT_PERFIL)
                  valorEjecucion = mapa.getValorRS(rsMapa,MT_EJECUCION)
                  if groupBy = 1 and printTitle then
                    if numper <> 0 then
                      printPromedios mapa, numper, numColsIniciales, false, hideDetail
                      numper = 0
                    end if
                    output = "<tr><th colspan='20'>" & puesto & "</th></tr>"
                    response.write output
                  end if
                  estilo=switchEstilo(estilo) 
                  strOut = "<tr class='" & estilo & "' onMouseOut='inOut(this);' onMouseOver='inOver(this);' onBlur='inBlur(this);' onFocus='inFocus(this);'><td>"
                  if mapaTalento_MostrarCampoLoginEnListado then
                    strOut = strOut & getDescripcion("Personal", "IdPersonal", khorLoginField(tipoLogin), rsNum(rsMapa, "IdPersonal"))&"</td><td>"
                  end if
                  if conExpediente and export = 0 then
                    strOut = strOut & "<a href='#' onClick='return abreExpediente(""" & urlExpediente & """," & rsNum(rsMapa,"IdPersonal") & ");'>"
                  end if
                  strOut = strOut & rsMapa("Nombre")
                  if conExpediente then
                    strOut = strOut & "</a>"
                  end if
                  strOut = strOut & "</td>"
                  if idperfil=0 and groupBy=0 then
                    strOut = strOut & "<td>" & puesto & "</td>"
                  end if
                  for eje=1 to 2
                    ejevalor(eje) = mapa.getValorRS(rsMapa,eje)
                    for i=1 to mapa.mtNum
                      if mapa.mtArr(i).tipo=eje AND mapa.mtArr(i).peso>0 then
                        strOut = strOut & mapa.mtArr(i).itemXtraCols(rsMapa,export)
                        tmpVal = mapa.mtArr(i).getCeldaRS(rsMapa,idp,3)                        
                        if tmpVal = "-" then tmpVal = 0
                        mapa.mtArr(i).total = mapa.mtArr(i).total + cdbl(tmpVal)
                        mapa.mtArr(i).total2 = mapa.mtArr(i).total2 + cdbl(tmpVal)
                        strOut = strOut & "<td align='center'>" & mapa.mtArr(i).getCeldaRS(rsMapa,idp,export) & "</td>"
                      end if
                    next
                    if ejenumcol(eje)>1 then
                      strOut = strOut & "<td align='center'>" & khorFormatPorcentaje(ejevalor(eje) / 100) & "</td>"
                      sumavalor(eje) = sumavalor(eje) + ejevalor(eje)
                      sumavalor2(eje) = sumavalor2(eje) + ejevalor(eje)
                    end if
                  next
                  idxCuadrante = mapaIndexCuadrante( ejevalor(MT_PERFIL), ejevalor(MT_EJECUCION) )
                  bgcuadrante = cuadranteColor(idxCuadrante)
                  cuadrante = cuadranteTxt(idxCuadrante)
                  strOut = strOut & "<td align='center' style='filter:;background:" & bgcuadrante & "'>" & cuadrante & "</td>"
                  for i=1 to mapaTalento_NumColAdListado
                    strOut = strOut & "<td>" & mapaTalento_ColAdValue(i,rsMapa) & "</td>"
                  next 
                  strOut = strOut & "</tr>"
                  if hideDetail = 0 then
                    response.write strOut
                  end if
                  '===== Continua el ciclo de personas
                  numper = numper + 1
                  numper2 = numper2 + 1
                  
                  rsMapa.movenext
                wend
                on error resume next
                printPromedios mapa, numper, numColsIniciales, false, hideDetail
                if groupBy = 1 then
                  printPromedios mapa, numper2, numColsIniciales, true, false
                end if
                rsMapa.close
                set rsMapa = nothing 
                %>
            </table>
            <% khorLeyendaModulos %>
          </div> <% 
      '==================== CONFIGURACION ====================
      if export = 0 AND NOT printerfriendly then %>
          <div id=tabC2 style="display:<%=displayStyle(printerfriendly)%>;">
            <table cellSpacing="10" cellPadding="0" border="0" align="center">
              <tr valign="top"> <%
              for eje=1 to 2 %>
                <td>
                  <table cellSpacing="1" cellPadding="1" border="0" align="center">
                    <tr class="celdaTit">
                      <td colspan="2" align="center"><%=mapaAxisName(eje)%></td>
                    </tr> <%
                allowInput = (mapa.numItems(eje,false)>1)
                for i=1 to mapa.mtNum
                  if mapa.mtArr(i).tipo = eje then
                    estilo = switchEstilo(estilo) %>
                    <tr class="<%=estilo%>">
                      <td><%=mapa.mtArr(i).desc%></td>
                      <td align="right"><%
                  if allowInput then %>
                        <INPUT type="text" name="peso_<%=mapa.mtArr(i).id%>" value="<%=mapa.mtArr(i).peso%>" maxlength="3" onblur="inBlur(this);valida(this,'int',0,100);mapaSetTotal(<%=eje%>);" style="WIDTH:30px;text-align:right;" class="whiteblur" onmouseover="inOver(this);" onfocus="inFocus(this);" onmouseout="inOut(this);"> <%
                  else %><%=mapa.mtArr(i).peso%>
                        <INPUT type="hidden" name="peso_<%=mapa.mtArr(i).id%>" value="<%=mapa.mtArr(i).peso%>"<%
                  end if %>
                      </td>
                    </tr> <%
                  end if
                next %>
                    <tr class="celdaTit">
                      <td align="right"><%=lblFRS_Total%>:</td>
                      <td>
                        <div id="totEje<%=eje%>" align="right"><%=mapa.SumaPesos(eje)%>%</div>
                      </td>
                    </tr>
                  </table>
                </td><%
              next %>
              </tr>
              <tr>
                <td colspan="2" align="center">
                  <INPUT type=button value="<%=lblFRS_Actualizar%>" onclick="tabSel(0);cambiaLista();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);>
                </td>
              </tr>
            </table>
            <div align="center">
              <%=lblKHOR_SiAsiganPesoCeroSeOmiteResultado%>
            </div> <%
            khorLeyendaModulos%>
          </div> <%
      end if
      '==================== TABS CLOSE ====================
      IF NOT printerfriendly THEN %>
        </td>
        </tr>
      </table> <% 
        if export = 0 then %>
      <script language="JavaScript">
        tabSel(<%=curtab%>);
      </script> <% 
        end if
      END IF
    ELSE
      for i=1 to mapa.mtNum %>
      <input type="hidden" name="peso_<%=mapa.mtArr(i).id%>" value="<%=mapa.mtArr(i).peso%>"> <%
      next
    END IF

set mapa = nothing
'========================================
if export = 2 then
  response.write vbCRLF & "</body>" & vbCRLF & "</html>"
else
  defaultFormEnd "", "", not printerfriendly
  layoutEnd
end if
'========================================
%>