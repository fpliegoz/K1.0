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
  
  'childwin=true
  'mov = "PDF"

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
  fechaReporte = fechaReporteIntegral(idpersona,idperfil)
  
  '-- Recalculo compatibilidad e inicializacion
  if persona.idpersona>0 and perfil.idperfil>0 then
    call khorCalculaCompatibilidadPersonaPerfil(conn,persona.idpersona,perfil.idperfil)
    remp_inicializa
  end if

  '-- Titulos
  tit = rempLBL_Titulo
  tit1=lblKHOR_Persona&": "&persona.nombre
  tit2=lblKHOR_Perfil&": "&perfil.perfil
  
'================================================================================'

  sub remp_graficaCompatibilidad(id,valor,width)
    IF mov="PDF" THEN %>
    <div class="contenedor__grafica">
      <canvas id="chart-area-<%=id%>"></canvas>
      <h1><%=khorFormatResultado(valor*100)%>%</h1>
    </div>
    <script language="JavaScript">
    <!--
      var opt<%=id%> = {
          type: 'doughnut',
          data: {
            datasets: [{ 
              data: [ <%=khorFormatResultado(valor*100)%>, <%=khorFormatResultado(100-(valor*100))%> ],
              backgroundColor: [ "#357DD5", "#D1EAFF" ]
            }],
            labels: [ "<%=strJS(lblKHOR_Compatibilidad)%>", "" ]
          },
          options: { <%
            IF id = "gral" THEN %>
            cutoutPercentage: 75,
            rotation: -Math.PI,
            circumference: Math.PI, <%
            ELSE %>
            cutoutPercentage: 85, <%
            END IF %>
            responsive: true,
            legend: { display: false },
            tooltips: { enabled: false }
          },
          animation: {
            animateScale: true,
            animateRotate: true
          }
      };
      $( document ).ready(function() {
        var ctx = document.getElementById("chart-area-<%=id%>").getContext("2d");
        window.myDoughnut = new Chart(ctx, opt<%=id%>);
      });
    //-->
    </script> <%
    ELSE %>
      <h1><%=khorFormatResultado(valor*100)%>%</h1> <%
    END IF
  end sub

  '----------------------------------------
  
  sub remp_InterpretacionAlterna(i,titulo)
    dim auxs : auxs = remp_InterpretacionFactoresAlt(i)
    if auxs <> "" then
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=titulo%></h1>
          <hr/>
          <ul>
            <li><%=replace(auxs,vbCRLF,"</li><li>")%></li>
          </ul>
        </article>
      </div>
    </section>
    <%
    end if
  end sub
  
'================================================================================'

'================================================================================'
response.write "<!DOCTYPE html>"
layoutHeadStart khorAppName() & " - " & tit
%>
<meta http-equiv="X-UA-Compatible" content="IE=edge"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<link href="./styles/style.min.css" rel="stylesheet"/>
<link href="./styles/bootstrap.min.css" rel="stylesheet"/>
<link rel="stylesheet" href="./styles/font-awesome.min.css"/>
<script src="js/jquery-3.2.1.min.js" ></script>
<script src="js/Chart2.js"></script>
<%
includeJS
'----------------------------------------
IF mov<>"PDF" THEN
%>
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
if mov = "PDF" then
  response.write vbCRLF & "<BODY>"
else
  layoutStart tit, tit1, tit2, errmsg, khorWinWidth(), ""
  defaultFormStart thispage, " onSubmit=""return false;""", true
end if
'================================================================================'

  IF persona.idpersona=0 THEN %>
  <div class="alerta"><%=lblFRS_NoSeHaSeleccionado & " " & lblKHOR_Persona%></div>  <%
  END IF
  IF perfil.idperfil=0 THEN %>
  <div class="alerta"><%=lblFRS_NoSeHaSeleccionado & " " & lblKHOR_Perfil%></div> <%
  END IF
  IF persona.idpersona<>0 and perfil.idperfil<>0  THEN
  
    '==============================  DATOS GENERALES ==============================
    IF mov = "PDF" THEN
       colorResultado = compatibilidadColor(compatibilidad)
    %>
    <header>
	
      <section>
        <div class="top">
          <div class="logo__cont">
            <figure><img class="img-responsive" src="khorImg/LogoKhor.png" alt="logo"></figure>
          </div>
        </div>
      </section>
      <div class="container">
        <div class="portada">
        <article>
          <div class="header__container">
            <div class="header__portada">
              <h1><%=tit%></h1>
              <p><%=strLang(rempLBL_SubTitFecha,formatDateDisp(fechaReporte,true))%></p>
            </div>
          </div>
          <div class="body__container">
            <div class="body__portada">
              <div class="row">
                <div class="col-md-6">
                  <ul id="datos">
                    <li><p><%=lblKHOR_Persona%></p><img src="styles/images/personia.PNG"></img> <%=persona.nombre%></li>
                    <li><p><%=lblKHOR_Perfil%></p><img src="styles/images/portfolio.PNG"></img><%=perfil.perfil%></li>
                  </ul>
				  
                </div>
                <div class="col-md-6">
                  <div class="resultado__com">
                    <p><%=lblKHOR_CompatibilidadPersonaPerfil%></p>
                    <h3 style="color:<%=colorResultado%> !important"><%=khorFormatCompatibilidadDisp(compatibilidad,false)%></h3>
                    <p><%=descFromMenuFijo(rempMF_TipoCompatibilidad,pctipo)%></p>
                  </div>
                </div>
              </div>
            </div>
          </div>
		  
          <div class="footer__container">
            <div class="footer__portada">
              <p style="color:<%=colorResultado%> !important">
              <!-- <img src="styles/images/advertenia.PNG"></img> -->
              <img src="khorImg/spacer_<%=colorResultado%>.gif" border="1" width="22" height="22">
              <%=compatibilidadLeyenda(compatibilidad)%></p>
            </div>
          </div>
        </article>
        </div>
      </div>
    </header>
    <%
    ELSE
    %>
    <section>
      <div class="container">
        <article class="basica">
          <div class="row">
            <div class="col-md-6">
              <div class="tit2"><%=lblFRS_DatosGenerales%></div>
              <table border="0" cellspacing="1" cellpadding="1" style="margin-left:30px;"> <%
                dim perData : set perData = new frsCollection
                perData.addKeyDesc lblKHOR_Persona, persona.nombre, iif(khorPermisoExpediente(modulosActivos), "[<a href=""#"" onClick=""abreExpediente('" & khorExpedienteURL() & "'," & persona.idpersona & ");return false;"">" & lblFRS_Detalle & "</a>]", "")
                perData.addKeyDesc lblKHOR_Perfil, perfil.perfil, iif(khorPermisoPerfiles(modulosActivos), "[<a href=""#"" onClick=""abrePerfil('" & khorPerfilURL() & "'," & perfil.IdPerfil & ");return false;"">" & lblFRS_Detalle & "</a>]", "")
                for i=1 to perData.count()
                  set obj = perData.obj(i) %>
                <tr>
                  <td><b><%=obj.key%>:</b></td>
                  <td><%=obj.desc%></td>
                  <td class="noshowimp"><%=obj.aux%></td>
                </tr> <%
                next
                perData.clean
                set perData = nothing %>
              </table>
			  <br/>
			  <a target="_blank" href="./repEmpresarialIndigo.asp?idPersona=<%=idpersona%>&idPerfil=<%=idperfil%>">[Reporte Individual]</a>
            </div>
            <div class="col-md-6">
              <div class="tit2"><%=rempLBL_CompatibilidadGeneral%></div>
              <table border="0" cellspacing="5" cellpadding="5" style="margin-left:30px;">
                <tr>
                  <td style="padding:5px;vertical-align:center;"><img src="khorImg/spacer_<%=compatibilidadColor(compatibilidad)%>.gif" border="1" width="30" height="30"></td>
                  <td nowrap style="padding:5px;">
                    <div><b><%=lblKHOR_CompatibilidadPersonaPerfil%>:</b> <%=khorFormatCompatibilidadDisp(compatibilidad,false)%>
                      [<a href="#" onClick="return abreCompatibilidad(<%=idpersona%>,<%=IdPerfil%>,<%=pctipo%>)"><%=lblFRS_Detalle%></a>]
                    </div>
                    <div><%=compatibilidadLeyenda(compatibilidad)%></div>
                  </td>
                </tr>
              </table>
              <div style="text-align:center;padding:5px;"><% selectTipoCompatibilidad pctipo, "pctipo", (mov<>"PDF"), "sendval('','mov','');" %></div>
            </div>
          </div>
        </article>
      </div>
    </section>        
      <%
    END IF
    
    response.write vbCRLF & "<main>"
    
	  '==============================  Introducccion  ==============================
    if mov = "PDF" then '-- true para pruebas?
	  %>
    <section>
      <div class="container">
        <article class="intro">
          <h1><%=rempLBL_Introduccion%></h1>
          <%=rempLBL_MensajeIntroduccion%>
        </article>
      </div>
    </section>        
	  <%
    
	  '==============================  Compatibilidad general  ==============================
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_CompatibilidadGeneral%></h1>
          <hr/>
          <div class="row">
            <div class="col-md-6">
            <%
              remp_graficaCompatibilidad "gral", compatibilidad, "40%"
            %>
            </div>
            <div class="col-md-6">
              <p style="text-align:center;color:<%=colorResultado%> !important">
                <%=compatibilidadLeyenda(compatibilidad)%>
              </p>
              <%
                if mov<>"PDF" then %>
                <div align="center">[<a href="#" onClick="return abreCompatibilidad(<%=idpersona%>,<%=IdPerfil%>,<%=pctipo%>)"><%=lblFRS_Detalle%></a>]</div> <%
                end if
              %>
            </div>
          </div>
        </article>
      </div>
    </section>        
	  <%
    end if
    
    '==============================  Compatibilidad Especifica (psicometria-pruebas)  ==============================
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_CompatibilidadEspecifica%></h1>
          <hr/>
          <%
          if mov<>"PDF" then %>
          <div align="center">[<a href="#" onClick="abrePsicometria(<%=IdPersona%>,<%=IdPerfil%>,<%=pctipo%>);return false;"><%=lblFRS_Detalle%></a>]</div> <%
          end if
          %>
          <div class="table__container">
            <div class="table-responsive">
              <table class="table">
                <thead>
                  <tr class="celdaTit">
                    <th><%=rempLBL_Prueba%></th>
                    <th style="text-align:right"><%=lblKHOR_Compatibilidad%></th>
                    <th style="text-align:right"><%=lblFRS_Peso%></th>
                    <th style="text-align:right"><%=lblKHOR_Compatibilidad%><br/><%=lblFRS_Ponderada%></th>
                  </tr>
                </thead>
                <tbody>	<%
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
                    <td style="text-align:right"><%=rep.psi.displayResultado(i,IdPersona,Escenario)%></td>
                    <td style="text-align:right"><%=strPeso%></td>
                    <td style="text-align:right"><%=strCom%></td>
                  </tr> <%
                  next %>
                  <tr class="celdaTit">
                    <th style="text-align:right" colspan="3"><%=lblKHOR_Psicometria%>:</th>
                    <th style="text-align:right"><%=khorFormatPorcentaje(rep.psi.Compatibilidad)%></th>
                  </tr>
                </tbody>	
              </table>
            </div>
          </div>
        </article>
      </div>
    </section>        
    <%
    
    '==============================  Meta-factores, en Resumen (Pruebas) ==============================
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_MetaFactoresEnResumen%></h1>
          <hr/>
          <%
            c = 0
            for i=1 to rep.psi.pNum
              if rep.psi.pArr(i).aplicada then
                listaPruebasAplicadas = strAdd(listaPruebasAplicadas,",",rep.psi.pArr(i).idprueba)
              end if
              if rep.psi.pArr(i).aplicada AND rep.psi.pArr(i).bPondera then
                valor=rep.psi.pArr(i).valor  '-- resultado crudo %>
          <div class="row">
            <div class="col-md-6">
            <%
              remp_graficaCompatibilidad i, valor, "60%"
            %>
            </div>
            <div class="col-md-6">
              <div class="h3"><%=rep.psi.pArr(i).Prueba%></div> <%
              'response.write rep.psi.displayResultado(i,persona.IdPersona,escenario)
              interpretacion = remp_interpretacion( rep.psi.pArr(i).idprueba )
              if mov<>"PDF" then
                response.write vbCRLF & "<div>[<a href=""#"" onClick=""return abreResultadosPrueba(" & idpersona & "," & rep.psi.pArr(i).idprueba & "," & IdPerfil & ");"">" & lblFRS_Detalle & "</a>]</div>"
              end if
              if interpretacion<>"" then
                response.write vbCRLF & "<P style=""text-align:justify;"">" & interpretacion & "</P>"
              end if %>
            </div>
          </div>
          <hr/> <%
              end if
            next
          %>
        </article>
      </div>
    </section>    
    <%
    
    '==============================  Palabras descriptivas ==============================
    lpreg = remp_palabrasdescriptivas()
    if lpreg <> "" then
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_PalabrasDescriptivas%></h1>
          <hr/>
          <div class="tags">
            <ul id="my-list">
            <%
              arrpreg = split( getListaAleatoria(lpreg,"",remp_num_palabras), "," )
              for i=lbound(arrpreg) to ubound(arrpreg)
                auxs = p7_reactivo(arrpreg(i))
                auxs = mid(auxs,instr(auxs," ")+1) & " (" & p7_tooltip(arrpreg(i)) & ")" %>
              <li><%=auxs%></li> <%
              next
            %>
            </ul>
          </div>
        </article>
      </div>
    </section>
    <%
    end if
      
    '==============================  Desccripcion detallada del Perfil ==============================
    set colAux = new frsCollection
    remp_InterpretacionFactores colAux, 0
    if colAux.count > 0 then
    %>
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_DescripcionDetalladaPerfil%></h1>
          <hr/>
          <%
          for i=1 to colAux.count
            set auxo = colAux.obj(i)
            id = i
            if not isnull(auxo.ResultadoPersona) then %>
            <div class="row">
              <div class="col-md-6">
                <h3><%=auxo.NombreFactor%></h3>
                <div align="center">Aqui va la grafica SLIDER del resultado (<%=auxo.ResultadoPersona%> en una escala con un maximo de <%=auxo.ResultadoMaximo%>)</div>
              </div>
              <div class="col-md-6">
                <!--div class="contenedor__grafica"-->
                <div style="background:white;box-shadow:0 2px 20px #346ba459;padding:5px;width:100%;height:150px">
                  <canvas id="chart-area-F<%=id%>"></canvas>
                </div>
                <script language="JavaScript">
                <!--
                  var optF<%=id%> = {
                      type: 'horizontalBar',
                      data: {
                        datasets: [{ 
                          data: [ <%=auxo.ResultadoPersona%> ],
                          backgroundColor: [ "#357DD5" ]
                        }]
                      },
                      options: {
                        responsive: true,
                        legend: { display: false },
                        tooltips: { enabled: true },
                        scales: {
                          xAxes: [{
                            ticks: { 
                              beginAtZero: true, 
                              max: <%=auxo.ResultadoMaximo%>
                            } 
                          }]
                        }
                      }
                  };
                  
                  $( document ).ready(function() {
                    var ctx = document.getElementById("chart-area-F<%=id%>").getContext("2d");
                    window.myBar = new Chart(ctx, optF<%=id%>);
                  });
                //-->
                </script>
              </div>
              <div><p><%=auxo.Interpretacion%></p></div>
            </div>
            <hr/> <%
            end if
          next
          set auxo = nothing
          %>
        </article>
      </div>
    </section>
    <%
    end if
    colAux.clean
    set colAux = nothing
        
    '==============================  Que lo motiva / Necesita / Limitaciones  ==============================
    remp_InterpretacionAlterna 1, rempLBL_QueLoMotiva
    remp_InterpretacionAlterna 2, rempLBL_Necesita
    remp_InterpretacionAlterna 3, rempLBL_Limitaciones
    
    '==============================  Competencias  ==============================
    
    '-- Con codigo copiado de creaGraficaCompetencias (CompetenciaResultadosClass.asp), GraficaFlashEspecial (graficasFlash.asp)
    sub remp_graficacompetencias(colCom)
      dim Resultados : Resultados = ""
      dim ResPerfil : ResPerfil = ""
      dim numc, auxC
      for numc=1 to colCom.count
        set auxC = colCom.obj(numc)
        if idpersona>0 then
          Resultados = strAdd( Resultados, ",", auxC.valorPersona )
        end if
        if idperfil>0 then
          ResPerfil = strAdd( ResPerfil, ",", auxC.valorPerfil )
        end if
      next
      '-- Grafica
      dim Ancho : Ancho = 400
      dim Alto : Alto = cint(400 * Ancho / 550)
      Dim nivelmaximo : nivelmaximo = khorMaxNivelCompetencia()
      dim url : url = graficaURL() & "charts/Circular.ashx?divisionEach=1&showScale=true&labels=&labelAbbrs=&lblScore=&lblProfile=" & _
                "&topValue=" & nivelmaximo & "&hasProfile=" & iif(resPerfil="","false","true") & "&values=" & Resultados & "&profileValues=" & ResPerfil
      %>
      <div style="text-align:center;page-break-inside:avoid;"><img id="gfxCompetencias" border="0" alt="" width="<%=Ancho%>" height="<%=Alto%>" src="<%=url%>" style="-ms-interpolation-mode:bicubic;"></div>
      <%
      '-- Leyenda
      if idpersona>0 and idperfil>0 then %>
      <div align="center"><font color="#0000FF"><%=lblKHOR_Persona%></font> - <font color="#FF0000"><%=lblKHOR_Perfil%></font></div> <%
      end if
    end sub

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
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_GraficoRadarDeCompetencias%></h1>
          <hr/>
          <%
            remp_graficacompetencias colCom
            creaInterpretacion colCom, persona, auxperfil, gfxTipo, colTipo, nivelperfil
          %>
        </article>
      </div>
    </section>
    <%
      end if
      set auxperfil = nothing
      colCom.clean
      set colCom = nothing
      colTipo.clean
      set colTipo = nothing
    end if

    '==============================  Tendencia de comportamiento ==============================
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
    <section>
      <div class="container">
        <article class="basica">
          <figure><img src="styles/images/circulo.PNG"/></figure>
          <h1><%=rempLBL_TendenciaDeComportamiento%></h1>
          <hr/>
          <div align="center"><img id="graficaPsi" border="0" alt="" style="-ms-interpolation-mode:bicubic;" src="<%=url%>"></div>
        </article>
      </div>
    </section>
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
  set persona=nothing
  set perfil=nothing

extraBtns = ""
extraHtml = ""
IF mov="PDF" THEN
  response.write vbCRLF & "</main>"
  response.write vbCRLF & "<footer></footer></body></html>"
ELSE
  IF NOT childwin THEN
    extraBtns = lblFRS_Seleccionar & " " & lblKHOR_Persona & iif(selFiltro<>""," *","") & "||seleccionaPersona()" & _
        "@@" & lblFRS_Seleccionar & " " & lblKHOR_Perfil & "||abreSeleccion('PERFIL',false,'" & IdPerfil & "')"
  END IF
  if selFiltro<>"" then
    extraHtml = "<div class=""tsmall"">(*) " & lblKHOR_UsandoCondicionesBusqueda & "</span>"
  end if
  if idpersona=0 and idperfil=0 then
    extraHtml = strAdd( extraHtml, vbCRLF, "<script languaje=""javascript"">seleccionaPersona();</script>" )
  end if
  '================================================================================'
  defaultFormEnd extraBtns, extraHtml, true
  response.write vbCRLF & "</main>"
  layoutEnd
  '================================================================================'
END IF
'================================================================================'
%>
