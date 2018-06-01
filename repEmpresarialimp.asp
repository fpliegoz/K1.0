<!--#include file="./khorClass.asp"-->
<!--#include file="./personalClass.asp"-->
<!--#include file="./repEmpresarialClass.asp"-->
<%
  thispage = "repEmpresarial.asp"
  thispageid = "FRMrepEmpresarial"

  mov=ucase(reqplain("mov"))
  if mov<>"PDF" then
    childwin=(reqplain("childwin")=thispageid)
    selCond=reqplain("selCond")
    selFiltro = getSelectionFilter(CURRENT_PAGE_NAME,selCond)
  else
    sesionFromRequest
    childwin=true
  end if
  checaSesion ses_super&","&ses_adminid, "", ""
  IdEmpresa=khorEmpresaUsuario(0)

  errmsg=""
  modulosActivos = khorModulosActivos()
  'khorPermisoUsuarioSucursal(persona.idempresa)

  '-- Request parametros
  IdPersona= reqn("IdPersona")
  IdPerfil=reqn("IdPerfil")
  if IdPersona<>0 and IdPerfil=0 then
    IdPerfil=khorGetPerfilPersona(IdPersona)
  end if
  dim persona : set persona=new clsPersona
  dim perfil : set perfil=new clsPerfil
  persona.getfromDB conn, IdPersona
  perfil.getfromDB conn, IdPerfil
  
  '-- Recalculo compatibilidad e inicializacion
  if persona.idpersona>0 and perfil.idperfil>0 then
    call khorCalculaCompatibilidadPersonaPerfil(conn,persona.idpersona,perfil.idperfil)
    remp_inicializa
  end if

  '-- Titulos
  tit = rempLBL_Titulo
  tit1=lblKHOR_Persona&": "&persona.nombre
  tit2=lblKHOR_Perfil&": "&perfil.perfil

  '-- Datos de la persona
  dim perData : set perData = new frsCollection
  perData.addKeyDesc lblKHOR_Persona, persona.nombre, iif(khorPermisoExpediente(modulosActivos), "[<a href=""#"" onClick=""abreExpediente('" & khorExpedienteURL() & "'," & persona.idpersona & ");return false;"">" & lblFRS_Detalle & "</a>]", "")
  perData.addKeyDesc lblKHOR_Perfil, perfil.perfil, iif(khorPermisoPerfiles(modulosActivos), "[<a href=""#"" onClick=""abrePerfil('" & khorPerfilURL() & "'," & perfil.IdPerfil & ");return false;"">" & lblFRS_Detalle & "</a>]", "")
  
'================================================================================'

  sub remp_graficaCompatibilidad(id,valor,width)
  %>
    <center>
      <div id="canvas-holder" style="width:<%=width%>">
        <canvas id="chart-area-<%=id%>" />
      </div>
      <div><%=khorFormatResultado(valor*100)%></div>
    </center>
    <script language="JavaScript">
    <!--
      var opt<%=id%> = {
          type: 'doughnut',
          data: {
            datasets: [{ 
              data: [ <%=khorFormatResultado(valor*100)%>, <%=khorFormatResultado(100-(valor*100))%> ],
              backgroundColor: [ "#2980B9", "#C0392B" ]
            }],
            labels: [ "<%=strJS(lblKHOR_Compatibilidad)%>", "" ]
          },
          options: {
            cutoutPercentage: 75,
            rotation: -Math.PI,
            responsive: true,
            circumference: Math.PI,
            legend: {
              display: false,
            },
            animation: {
              animateScale: true,
              animateRotate: true
            }
          }
      };
      $( document ).ready(function() {
        var ctx = document.getElementById("chart-area-<%=id%>").getContext("2d");
        window.myDoughnut = new Chart(ctx, opt<%=id%>);
      });
    //-->
    </script>
  <%
  end sub

  '----------------------------------------
  
  sub remp_InterpretacionAlterna(i,titulo)
    dim auxs : auxs = remp_InterpretacionFactoresAlt(i)
    if auxs <> "" then
    %>
      <div style="page-break-inside:avoid;">
        <div class="tit2"><%=titulo%></div>
        <br/>
      </div>
      <ul>
        <li><%=replace(auxs,vbCRLF,"</li><li>")%></li>
      </ul>
    <%
    end if
  end sub
  
'================================================================================'

'================================================================================'
layoutHeadStart khorAppName() & " - " & tit
IF mov<>"PDF" THEN
  includeJS
'----------------------------------------
%>
<script src="js/Chart.min.js"></script>
<script src="js/utils.js"></script>
<script language="JavaScript">
<!--
function seleccionaPersona() {
  abreSeleccion('PERSONA',false,'<%=IdPersona%>',null,null,null,null,'<%=setSelectionFilter("PERSONA",selFiltro)%>');
}
function setSeleccion(tipo,lista) {
  if (tipo=='PERSONA')
    setValor('IdPersona',lista);
  else if (tipo=='PERFIL')
    setValor('IdPerfil',lista);
  sendval('','mov','');
}
<% IF pdf_enabled() AND persona.idpersona>0 AND perfil.idperfil>0 THEN %>
function myPrintPage() { <%
  pdfkey = initPDFurl( thispageid & "_" & idpersona & "_" & idperfil, _
                        pdf_URL() & thispage & "?mov=pdf&idpersona=" & idpersona & "&idperfil=" & idperfil & "&pctipo=" & pctipo ) %>
  openPDFjob( '<%=pdfkey%>' );
}
<% END IF %>
//-->
</script> <% 
END IF
'----------------------------------------
layoutHeadEnd
layoutStart tit, tit1, tit2, errmsg, khorWinWidth(), ""
defaultFormStart thispage, " onSubmit=""return false;""", true
'================================================================================'

  IF persona.idpersona=0 THEN %>
  <div class="alerta"><%=lblFRS_NoSeHaSeleccionado & " " & lblKHOR_Persona%></div>  <%
  END IF
  IF perfil.idperfil=0 THEN %>
  <div class="alerta"><%=lblFRS_NoSeHaSeleccionado & " " & lblKHOR_Perfil%></div> <%
  END IF
  IF persona.idpersona<>0 and perfil.idperfil<>0  THEN
  
      '==============================  DATOS GENERALES ============================== %>
      <TABLE BORDER="0" CELLSPACING="5" CELLPADDING="0" width="100%">
        <TR VALIGN="top">
          <TD>
            <div class="tit2"><%=lblFRS_DatosGenerales%></div>
            <table border="0" cellspacing="1" cellpadding="1" style="margin-left:30px;"> <%
              for i=1 to perData.count()
                set obj = perData.obj(i) %>
              <tr>
                <td><b><%=obj.key%>:</b></td>
                <td><%=obj.desc%></td>
                <td class="noshowimp"><%=obj.aux%></td>
              </tr> <%
              next %>
            </table>
          </TD>
          <TD>
            <div class="tit2"><%=lblFRS_Recomendacion%></div>
            <table border="0" cellspacing="1" cellpadding="1" style="margin-left:30px;">
              <tr>
                <td rowspan="2" valign="center"><img src="khorImg/spacer_<%=compatibilidadColor(compatibilidad)%>.gif" border="1" width="30" height="30"></td>
                <td nowrap><b><%=lblKHOR_Compatibilidad%>:</b> <%=khorFormatCompatibilidadDisp(compatibilidad,false)%></td>
              </tr>
              <tr>
                <td class="tit3"><%=compatibilidadLeyenda(compatibilidad)%></td>
              </tr>
            </table> <%
            selectTipoCompatibilidad pctipo, "pctipo", (mov<>"PDF"), "sendval('','mov','');" %>
          </TD>
        </TR>
      </TABLE> <br/>
	  <%
    
	  '==============================  Introducccion  ==============================
    if true OR mov = "PDF" then '-- true para pruebas?
	  %>
	  <DIV style="page-break-inside:avoid;">
      <div class="tit2"><%=rempLBL_Introduccion%></div>
      <br/>
      <div id="introduccion"><%=rempLBL_MensajeIntroduccion%></div>
    </DIV>
	  <%
    end if
    
	  '==============================  Compatibilidad general  ==============================
    %>
    <DIV style="page-break-inside:avoid;">
      <div class="tit2"><%=rempLBL_CompatibilidadGeneral%></div>
      <br/>
      <%
        remp_graficaCompatibilidad "gral", compatibilidad, "40%"
      if mov<>"PDF" then %>
      <div align="center">[<a href="#" onClick="return abreCompatibilidad(<%=idpersona%>,<%=IdPerfil%>,<%=pctipo%>)"><%=lblFRS_Detalle%></a>]</div> <%
      end if %>
    </DIV>
	  <%
    
    '==============================  Compatibilidad Especifica (psicometria-pruebas)  ==============================
    %>
    <DIV style="page-break-inside:avoid;">
      <div class="tit2"><%=rempLBL_CompatibilidadEspecifica%></div>
      <br/>
      <table border=0 cellspacing=1 cellpadding=2 class="tsmall" align="center">
        <tr class="celdaTit">
          <td><%=rempLBL_Prueba%></td>
          <td><%=lblKHOR_Compatibilidad_Abrev2%></td>
          <td><%=lblFRS_Peso%></td>
          <td><%=lblKHOR_Compatibilidad_Abrev2%><br><%=lblFRS_Pond%></td>
        </tr> <%
        '-- Logica extraida de la funcion pcTablePsi de clsCompatibilidadReporte en khorCompatibilidadClass.asp
        for i=1 to rep.psi.pNum
          estilo = switchEstilo(estilo)
          if rep.psi.pArr(i).bPondera then
            strPeso = khorFormatPorcentajeInt(rep.psi.pArr(i).peso)
            strCom = khorFormatPorcentaje(rep.psi.pArr(i).compatibilidad)
          else
            strPeso = lblFRS_N_P
            strCom = lblFRS_N_P
          end if
          if not rep.psi.pArr(i).aplicada then
            strCom = lblFRS_N_R
          end if %>
        <tr class="<%=subestilo%>">
          <td><%=rep.psi.pArr(i).Prueba%></td>
          <td align="right"><%=rep.psi.displayResultado(i,IdPersona,Escenario)%></td>
          <td align="right"><%=strPeso%></td>
          <td align="right"><%=strCom%></td>
        </tr> <%
        next %>
        <tr class="celdaTit">
          <td colspan="3"><%=lblKHOR_Psicometria%>:</td>
          <td align="right"><%=khorFormatPorcentaje(rep.psi.Compatibilidad)%></td>
        </tr>
      </table>  <%
      if mov<>"PDF" then %>
      <div align="center">[<a href="#" onClick="abrePsicometria(<%=IdPersona%>,<%=IdPerfil%>,<%=pctipo%>);return false;"><%=lblFRS_Detalle%></a>]</div> <%
      end if %>
    </DIV>
    <%
    
    '==============================  Meta-factores, en Resumen (Pruebas) ==============================
    %>
    <DIV style="page-break-inside:avoid;">
      <div class="tit2"><%=rempLBL_MetaFactoresEnResumen%></div> <%
        c = 0
        for i=1 to rep.psi.pNum
          if rep.psi.pArr(i).aplicada then
            listaPruebasAplicadas = strAdd(listaPruebasAplicadas,",",rep.psi.pArr(i).idprueba)
          end if
          if rep.psi.pArr(i).aplicada AND rep.psi.pArr(i).bPondera then
            valor=rep.psi.pArr(i).valor  '-- resultado crudo %>
        <div style="page-break-inside:avoid;">
          <table border="0" cellspacing="10" cellpadding="10" width="100%">
            <tr>
              <td style="width:40%;vertical-align:middle;">
                <%
                  remp_graficaCompatibilidad i, valor, "60%"
                %>
              </td>
              <td style="width:60%;vertical-align:middle;">
                <div class="tit3"><%=rep.psi.pArr(i).Prueba%></div> <%
                response.write rep.psi.displayResultado(i,persona.IdPersona,escenario)
                interpretacion = remp_interpretacion( rep.psi.pArr(i).idprueba )
                if mov<>"PDF" then
                  response.write vbCRLF & "<div>[<a href=""#"" onClick=""return abreResultadosPrueba(" & idpersona & "," & rep.psi.pArr(i).idprueba & "," & IdPerfil & ");"">" & lblFRS_Detalle & "</a>]</div>"
                end if
                if interpretacion<>"" then
                  response.write vbCRLF & "<P style=""text-align:justify;"">" & interpretacion & "</P>"
                end if %>
              </td>
            </tr>
          </table>
        </div> <%
            if c = 0 then %>
    </DIV> <%
            end if
            c = c + 1
          end if
        next
      
    '==============================  Palabras descriptivas ==============================
    lpreg = remp_palabrasdescriptivas()
    if lpreg <> "" then %>
      <DIV style="page-break-inside:avoid;">
        <div class="tit2"><%=rempLBL_PalabrasDescriptivas%></div>
        <ul id="palabras">
        <%
          arrpreg = split( getListaAleatoria(lpreg,"",remp_num_palabras), "," )
          for i=lbound(arrpreg) to ubound(arrpreg)
            auxs = p7_reactivo(arrpreg(i))
            auxs = mid(auxs,instr(auxs," ")+1) & " (" & p7_tooltip(arrpreg(i)) & ")" %>
          <li><%=auxs%></li> <%
          next
        %>
        </ul>
      </DIV> <%
    end if
      
    '==============================  Desccripcion detallada del Perfil ==============================
    set colAux = new frsCollection
    remp_InterpretacionFactores colAux, 0
    if colAux.count > 0 then
    %>
      <div style="page-break-inside:avoid;">
        <div class="tit2"><%=rempLBL_DescripcionDetalladaPerfil%></div>
        <br/>
      </div> <%
      for i=1 to colAux.count
        set auxo = colAux.obj(i)
        if not isnull(auxo.ResultadoPersona) then %>
        <div style="margin-top:5px;page-break-inside:avoid;">
          <div class="tit3"><%=auxo.NombreFactor%></div>
          <div align="center">Aqui va la grafica del resultado (<%=auxo.ResultadoPersona%> en una escala con un maximo de <%=auxo.ResultadoMaximo%>)</div>
          <p><%=auxo.Interpretacion%></p>
        </div> <%
        end if
      next
      set auxo = nothing
    end if
    colAux.clean
    set colAux = nothing
        
    '==============================  Que lo motiva / Necesita / Limitaciones  ==============================
    remp_InterpretacionAlterna 1, rempLBL_QueLoMotiva
    remp_InterpretacionAlterna 2, rempLBL_Necesita
    remp_InterpretacionAlterna 3, rempLBL_Limitaciones
    
    '==============================  Competencias  ==============================
    if persona.conCompetencias then
      dim gfxTipo : gfxTipo = 0 '-- radar
      dim colCom : set colCom = new frsCollection
      dim colTipo : set colTipo = new frsCollection
      dim auxperfil
      if khorPerfilTieneCompetencia(perfil.IdPerfil,0) then
        set auxperfil = perfil
      else
        set auxperfil = new clsPerfil 
      end if
      cargaCompetencias colCom, persona, auxperfil, gfxTipo, colTipo, Escenario, "", 0, true
      if colCom.count>0 then %>
      <div style="page-break-inside:avoid;">
        <div class="tit2"><%=rempLBL_GraficoRadarDeCompetencias%></div> <%
        creaGraficaCompetencias colCom, persona, auxperfil, gfxTipo, colTipo %>
      </div> <%
        creaInterpretacion colCom, persona, auxperfil, gfxTipo, colTipo, nivelperfil
      end if
      set auxperfil = nothing
      colCom.clean
      set colCom = nothing
      colTipo.clean
      set colTipo = nothing
    end if

    '==============================  Tendencia de comportamiento ==============================
    %>
      <div style="page-break-inside:avoid;">
        <div class="tit2"><%=rempLBL_TendenciaDeComportamiento%></div>
        <br/>
      </div>
    <%
      '-- Obtiene datos
      dim colFactores : set colFactores = new frsCollection '--  of clsPruebaFactor
      remp_FactoresOpuestos colFactores
      '-- Grafica
      dim auxper : auxper = ""  '-- CSV de resultados de la persona
      dim auxpue : auxpue = ""  '-- CSV de valores del perfil
      dim lblPares : lblPares = ""  '-- CSV de etiquetaIzquierda:etiquetaDerecha
      dim showPuesto : showPuesto = false '-- Para incluir perfil en la gráfica
      'dim edri_PruebasResumenLimite : edri_PruebasResumenLimite = 60 '-- Para la interpretcion binaria (bajo: izquierda | alto: derecha)
      dim f, objF
      for f=1 to colFactores.count
        set objF = colfactores.obj(f)
        auxper = strAdd( auxper, ",", objF.valPersona )
        if showPuesto then
          auxpue = strAdd( auxpue, ",", objF.valPerfil )
        end if
        lblAbreviaturas = strAdd( lblAbreviaturas, ",", f )
        lblPares = strAdd( lblPares, ",", objF.poloBajo & ":" & objF.poloAlto )
        '-- interpretacionBinaria = f & ". " & objF.PruebaFactor & ": " & iif( objF.valPersona < edri_PruebasResumenLimite, objF.definicionBajo, objF.definicionAlto )
      next
      set objF = nothing
      colFactores.clean
      set colFactores = nothing
      dim url : url = graficaURL() & "charts/Pairs.ashx?values=" & auxper & "&hasProfile=" & iif(auxpue<>"","true","false") & "&profileValues=" & auxpue & "&labels=" & FrsEncode(lblPares) & "&labelAbbrs=" & lblAbreviaturas & "&topValue=100&bottomValue=0&divisionEach=10&shadowArea=-1,-1&hideScale=1"
      %>
      <div align="center"><img id="graficaPsi" border="0" alt="" style="-ms-interpolation-mode:bicubic;" src="<%=url%>"></div>
    <%
      
    '==============================  FIN ==============================
      
    remp_finaliza
    
  END IF %>
      
  <input type="hidden" name="mov" value="">
  <input type="hidden" name="IdPersona" value="<%=IdPersona%>">
  <input type="hidden" name="IdPerfil" value="<%=IdPerfil%>">
  <input type="hidden" name="childwin" value="<%=server.htmlencode(reqplain("childwin"))%>">
  <input type="hidden" name="selCond" value="<%=server.htmlencode(selCond)%>">
<%        
'================================================================================'

extraBtns = ""
extraHtml = ""
IF mov<>"PDF" THEN
  IF NOT childwin THEN
    extraBtns = lblFRS_Seleccionar & " " & lblKHOR_Persona & iif(selFiltro<>""," *","") & "||seleccionaPersona()" & _
        "@@" & lblFRS_Seleccionar & " " & lblKHOR_Perfil & "||abreSeleccion('PERFIL',false,'" & IdPerfil & "')"
  END IF
  if selFiltro<>"" then
    extraHtml = "<div class=""tsmall"">(*) " & lblKHOR_UsandoCondicionesBusqueda & "</span>"
  end if
  if persona.idpersona=0 and perfil.idperfil=0 then
    extraHtml = strAdd( extraHtml, vbCRLF, "<script languaje=""javascript"">seleccionaPersona();</script>" )
  end if
END IF

  perData.clean
  set perData = nothing
  set persona=nothing
  set perfil=nothing

'================================================================================'
defaultFormEnd extraBtns, extraHtml, true
layoutEnd
'================================================================================'
%>
