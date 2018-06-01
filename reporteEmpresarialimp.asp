<!--#include file="./khorClass.asp"-->
<!--#include file="./personalClass.asp"-->
<!--#include file="./repEmpresarialClass.asp"-->
<%
  thispage = "repEmpresarialimp.asp"
  thispageid = "FRMrepEmpresarial"
  'Estos valores estan fijos me los imagino haciendo que desde el reporte de LP salten a esta pagina con estos valores en post o get cuando se haga quitar el comentarios abajo
  IdPersona= 32859
  IdPerfil=0
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

  '-- Request parametrosQuitar esto jamas recibiremos sin estas variables
  'IdPersona= reqn("IdPersona")
  'IdPerfil=reqn("IdPerfil")
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
<article class="interior"><div class="row"><div class="col-md-6">

<div class="contenedor__grafica">
<canvas id="chart-area-<%=id%>"></canvas>
</div></div>
 
      <!--<h1><%=khorFormatResultado(valor*100)%></h1>-->
    <script language="JavaScript">
    <!--
      var opt<%=id%> = {
          type: 'doughnut',
          data: {
            datasets: [{ 
              data: [ <%=khorFormatResultado(valor*100)%>, <%=khorFormatResultado(100-(valor*100))%> ],
              backgroundColor: ["#357DD5", "#D1EAFF"]
            }],
            labels: [ "<%=strJS(lblKHOR_Compatibilidad)%>", "" ]
          },
         options: {
      cutoutPercentage: 85,
      legend: { display: false },
      tooltips: {
        enabled: false
      }
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
    </script>
  <%
  end sub
  
  
   sub remp_graficaCompatibilidad2(id,valor,titulo2, interpre)
  
%>  
<article class="interior"><div class="row"><div class="col-md-6">
 <div class="contenedor__grafica"><canvas id="<%=id&id&id%>"></canvas><h1><%=valor%>%</h1></div></div><div class="col-md-6">
 <h3><%=titulo2%></h3>
 <p class="inter">Interpretación</p><p><%=interpre%>
 </p>
</div></div><hr/></article>
   <script type="text/javascript">$(document).ready(function () {
  var ctx_<%=id&id&id%> = document.getElementById("<%=id&id&id%>");
  var donut_<%=id&id&id%> = new Chart(ctx_<%=id&id&id%>, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [<%=valor%>,<%=100-valor%> ],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      rotation: 1 * Math.PI,
      circumference: 1 * Math.PI,
      tooltips: {
        enabled: false
      }
    }
  });
});
</script>
  <%
  end sub

  '----------------------------------------
  
  sub remp_InterpretacionAlterna(i,titulo)
    dim auxs : auxs = remp_InterpretacionFactoresAlt(i)
    if auxs <> "" then
    %>
      
        <li><%=replace(auxs,vbCRLF,"</li><li>")%></li>
      
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
%>
<!DOCTYPE html><html>
<head>
<title>Reporte Integral</title>

<meta http-equiv="X-UA-Compatible" content="IE=edge"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<link href="./styles/style.min.css" rel="stylesheet"/>
<link href="./styles/bootstrap.min.css" rel="stylesheet"/>
<link rel="stylesheet" href="./styles/font-awesome.min.css"/>
<script src="js/jquery-3.2.1.min.js" ></script>
<script src="js/Chart2.js"></script>
</head>
<body>
<header>
   <section><div class="top">
   <div class="logo__cont">
   <figure>
	<img class="img-responsive" src="khorImg/LogoKhor.png" alt="logo">
   </figure>
   </div>
   </div></section>
  
   <div class="container">
	<div class="portada"><article>
			<div class="header__container"><div class="header__portada"><h1>Reporte Integral</h1>
			<%'<!--Falta sacarlo<p>De la evaluación presentada el 28/Feb/2018 17:55</p>-->
			%></div></div>
			<div class="body__container">
			<div class="body__portada">
			<div class="row">
			<div class="col-md-6">
			 <%
				set obj1 = perData.obj(1)
                set obj2 = perData.obj(2)
			%>
			<ul id="datos"><li><p>Persona</p><img src="styles/images/personia.PNG"></img> <%=obj1.desc%></li>
			                <li><p>Perfil</p><img src="styles/images/portfolio.PNG"></img><%=obj2.desc%></li>
			</ul>
			</div>
			<div class="col-md-6">
			<div class="resultado__com">
			         <p>Compatibilidad persona/puesto</p><h3 class="nocompatible"><%=khorFormatCompatibilidadDisp(compatibilidad,false)%></h3>
					 <p>Reclutamiento y selección</p>
			</div>
			</div>
			</div>
			</div>
			</div>
	<div class="footer__container">
	<div class="footer__portada">
	    <p class="nocompatible"><img src="styles/images/advertenia.PNG"></img>
		                         </i> <%=compatibilidadLeyenda(compatibilidad)%></p>
     </div>
	 </div>
	 </article>
	</div>
   </div>
   </header>
   
   <main>
   <section>
   <div class="container">
   <article class="intro">
     <h1>Introducción</h1>
	 <p> KHOR es una herramienta de evaluación robusta que permite conocer sus áreas de oportunidad y fortalezas, 
	    en aspectos relacionados con el trabajo y su vida personal. Los instrumentos contenidos han pasado por 
		rigurosos exámenes para asegurar su validez, confiabilidad y estandarización en población mexicana.
	</p>
	<p>Si usted ha seguido las instrucciones durante la fase de evaluación, usted notará que este 
	 reporte lo describe a profundidad, en distintos aspectos como competencias, habilidades, 
	 estilos de trabajo, y preferencias personales.</p><p>Hay algunas situaciones que pueden incidir 
	 en una evaluación poco válida. Es decir, que no lo describa con exactitud. Si usted presentó la 
	 evaluación en un estado emocional alterado, o bajo condiciones de estrés, o con distracciones frecuentes, 
	 es probable que su rendimiento no haya sido el esperado. De la misma manera, si usted respondió de manera 
	 poco sincera a las preguntas que se le hicieron, notará que sus resultados pueden ser contradictorios.</p>
	 <p>Recomendamos que lea este reporte con calma, y que identifique situaciones de su vida en las que usted 
	 manifiesta lo aquí descrito. Más importante aún, léalo con mente abierta, dispuesto a identificar áreas de 
	 ejora y a trazar un plan individual de desarrollo.</p><p class="text-right"> 
	 <strong>Algunas preguntas que usted debe tener en mente al leer este reporte, son:</strong>
	 </p>
	 <ul>
		<li>¿Cómo puedo usar esta información a mi favor?</li>
		<li>¿Cómo puedo aprovechar estas habilidades?</li>
		<li>¿Qué cosas debo corregir para potencializar mi desarrollo?</li>
		<li>¿Qué pasara si cambiara este rasgo o atributo indeseable? ¿Qué ganaría?</li>
	</ul>
	<p class="text-right">Atentamente,</p>
	<p><h2 class="text-right">KHOR</h2></p>
	</article>
	</div>
	</section>
  <section>
  <div class="container">
  <article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure>
  <h1>Compatibilidad general</h1>
  <hr/><div class="row">
  <div class="col-md-6">
  <div class="contenedor__grafica">
  <canvas id="compatibilidad_general"></canvas>
  <h1><%=khorFormatResultado(compatibilidad*100)%></h1>
  </div></div>
  <div class="col-md-6"><p>Interpretación</p>
  <p><%=compatibilidadLeyenda(compatibilidad)%>.</p>
  </div></div></article></div></section>
  <section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure>
  <h1>Compatibilidad Específica</h1><hr/>
  <div class="table__container">
	<div class="table-responsive">
		<table class="table"><thead><tr><th >Meta-factores</th>
										<th >Compatibilidad</th>
										<th >Peso</th>
										<th >Compatibilidad ponderada</th>
									</tr>
							</thead>
							<tbody>	
	 <script type="text/javascript">$(document).ready(function () {
  var ctx_compatibilidad_general = document.getElementById("compatibilidad_general");
  var donut_compatibilidad_general = new Chart(ctx_compatibilidad_general, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [<%=khorFormatResultado(compatibilidad*100)%>,<%=khorFormatResultado(100-(compatibilidad*100))%> ],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      rotation: 1 * Math.PI,
      circumference: 1 * Math.PI,
      tooltips: {
        enabled: false
      }
    }
  });
});
</script>
								<%
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
          <td><%=rep.psi.displayResultado(i,IdPersona,Escenario)%></td>
          <td><%=strPeso%></td>
          <td><%=strCom%></td>
        </tr> <%
        next %>
							</tbody>
			</table></div></div></article></div></section>
			<section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Meta-factores, en Resumen</h1><hr/></article>
		<%
		c = 0
        for i=1 to rep.psi.pNum
          if rep.psi.pArr(i).aplicada then
            listaPruebasAplicadas = strAdd(listaPruebasAplicadas,",",rep.psi.pArr(i).idprueba)
          end if
          if rep.psi.pArr(i).aplicada AND rep.psi.pArr(i).bPondera then
            valor=rep.psi.pArr(i).valor  '-- resultado crudo %>
        
                <%
                  remp_graficaCompatibilidad i, valor, "60%"
                %>
             
                <div class="tit3"><%=rep.psi.pArr(i).Prueba%></div> <%
                response.write rep.psi.displayResultado(i,persona.IdPersona,escenario)
                interpretacion = remp_interpretacion( rep.psi.pArr(i).idprueba )
                if mov<>"PDF" then
                  response.write vbCRLF & "<div>[<a href=""#"" onClick=""return abreResultadosPrueba(" & idpersona & "," & rep.psi.pArr(i).idprueba & "," & IdPerfil & ");"">" & lblFRS_Detalle & "</a>]</div>"
                end if
                if interpretacion<>"" then
                  response.write vbCRLF & "<P style=""text-align:justify;"">" & interpretacion & "</P>"&"</div><hr/></article>"
                end if %>
              <%
            if c = 0 then %>
     <%
            end if
            c = c + 1
          end if
        next
      
%>	


 <%
  '==============================  Palabras descriptivas ==============================
    lpreg = remp_palabrasdescriptivas()
    if lpreg <> "" then %>
       <section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/></figure>
 <h1>Palabras Descriptivas</h1></article><div class="tags">
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
      </DIV> <%
    end if
      
 
 %>
 
 </div></section>
 
 <%
  '==============================  Desccripcion detallada del Perfil ==============================
    set colAux = new frsCollection
    remp_InterpretacionFactores colAux, 0
    if colAux.count > 0 then
    %>
      <section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure>
 <h1>Descripción Detallada del Perfil</h1><hr/></article>
    <%
      for i=1 to colAux.count
        set auxo = colAux.obj(i)
        if not isnull(auxo.ResultadoPersona) then 
			remp_graficaCompatibilidad2  i,auxo.ResultadoPersona,auxo.NombreFactor, auxo.Interpretacion
        end if
      next
      set auxo = nothing
    end if
    colAux.clean
    set colAux = nothing
 %>
 
  
    
    
 <section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/>
 </figure><h1>¿Qué lo motiva?</h1><ul id="que-lo-motiva">
 <%remp_InterpretacionAlterna 1, rempLBL_QueLoMotiva%>
 
 </ul></article></div></section><section><div class="container">
 
 <article><figure><img src="styles/images/circulo.PNG"/></figure><h1>Necesita</h1><ul id="necesita">
  <%remp_InterpretacionAlterna 2, rempLBL_Necesita%>
 </ul></article></div></section><section><div class="container"><article>
 <figure><img src="styles/images/circulo.PNG"/></figure><h1>Limitaciones</h1>
 <ul id="limitaciones">
 <%remp_InterpretacionAlterna 3, rempLBL_Limitaciones%>
 </ul>
 </article></div>
 </section><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure>
 <h1>Gráfico Radar de Competencias</h1><hr/><div class="row">
 <div class="col-md-6"><!--<div class="contenedor__grafica"></div>--></div>
 <%if persona.conCompetencias then
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
        <div class="tit2">Gráfico</div> <%
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
	%>
 </div></article></div></section><script type="text/javascript">$(document).ready(function () {
  var ctx_tendencia_radar = document.getElementById("tendencia_radar");
  var donut_tendencia_radar = new Chart(ctx_tendencia_radar, {
    type: 'radar',
    data: {
      labels: [1,2,3,4,5,6,7],
      datasets: [
        {
          label: "Perfil",
          data: [10,20,5,50,30,20,100],
          backgroundColor: "rgba(53, 125, 213, 0.7)"
        },
        {
          label: "Persona",
          data: [20,20,20,20,20,20,20,100],
          backgroundColor: "rgba(209, 234, 255, 0.7)"
        }
      ]
    },
    options: {
      tooltips: { enabled: false }
    }
  });
});</script>
<%
'==============================  Tendencia de comportamiento ==============================
    %>
     <section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure>
<h1>Tendencia de Comportamiento</h1><hr/><div class="row"><div class="col-md-12"><div class="contenedor__grafica__range">
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
	  
	  
	  </div></div></div></div></div></article></div></section></main><footer></footer></body></html>