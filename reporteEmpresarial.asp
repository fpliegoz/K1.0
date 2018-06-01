<!--#include file="./khorClass.asp"-->
<!--#include file="./personalClass.asp"-->
<!--#include file="./repEmpresarialClass.asp"-->

<!DOCTYPE html><html>
<head>
<title>Reporte Integral</title>
<meta charset="utf-8"/>
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
			<div class="header__container"><div class="header__portada"><h1>Reporte Integral</h1
			<p>De la evaluación presentada el 28/Feb/2018 17:55</p></div></div>
			<div class="body__container">
			<div class="body__portada">
			<div class="row">
			<div class="col-md-6">
			<ul id="datos"><li><p>Persona</p><img src="styles/images/personia.PNG"></img> Candidato 1</li>
			                <li><p>Perfil</p><img src="styles/images/portfolio.PNG"></img>Psicólogo</li>
			</ul>
			</div>
			<div class="col-md-6">
			<div class="resultado__com">
			         <p>Compatibilidad persona/puesto</p><h3 class="nocompatible">44.4%</h3>
					 <p>Reclutamiento y selección</p>
			</div>
			</div>
			</div>
			</div>
			</div>
	<div class="footer__container">
	<div class="footer__portada">
	    <p class="nocompatible"><img src="styles/images/advertenia.PNG"></img>
		                         </i> No compatible con el Perfil</p>
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
  <h1>83.8%</h1>
  </div></div>
  <div class="col-md-6"><p>Interpretación</p>
  <p>Se considera que la persona es compatible con el puesto evaluado.</p>
  </div></div></article>
  <script type="text/javascript">$(document).ready(function () {
  var ctx_compatibilidad_general = document.getElementById("compatibilidad_general");
  var donut_compatibilidad_general = new Chart(ctx_compatibilidad_general, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [83.8, 16.200000000000003],
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
</script></div></section><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Compatibilidad Específica</h1><hr/><div class="table__container"><div class="table-responsive"><table class="table"><thead><tr><th>Meta-factores</th><th>Compatibilidad</th><th>Peso</th><th>Compatibilidad ponderada</th></tr></thead><tbody><tr><td>Estilo social</td><td>100%</td><td>20%</td><td>20%</td></tr><tr><td>Dominancia cerebral</td><td>100%</td><td>20%</td><td>20%</td></tr><tr><td>Motivación</td><td>88.7%</td><td>20%</td><td>17.7%</td></tr><tr><td>Valores</td><td>100%</td><td>20%</td><td>20%</td></tr></tbody></table></div></div></article></div></section><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Meta-factores, en Resumen</h1><hr/></article><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="estilo_social"></canvas><h1>100%</h1></div></div><div class="col-md-6"><h3>Estilo Social</h3><p class="inter">Interpretación</p><p>Es entusiasta y optimista. Se relaciona de manera equilibrada con un amplio número de personas. Se gana facilmente la confianza de la gente. Sabe planear su trabajo. Responde con rapidez ante las urgencias. Su empuje esta dirigido hacia el trabajo especializado y tecnico. Tiene facilidad para integrar al grupo de trabajo y participa con el para lograr las metas que se hayan propuesto. Reconoce y respeta normas establecidas.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_estilo_social = document.getElementById("estilo_social");
  var donut_estilo_social = new Chart(ctx_estilo_social, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [100, 0],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      tooltips: {
        enabled: false
      }
    }
  });
});
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="dominancia_cerebral"></canvas><h1>100%</h1></div></div><div class="col-md-6"><h3>Dominancia Cerebral</h3><p class="inter">Interpretación</p><p>Es analitico, logico, ordenado y sistematico. Es objetivo y racional. Trabaja de acuerdo a normas, procesos y metodos que aseguran la exactitud de los resultados. Utiliza conceptos teoricos, ideas y metodos practicos en los que presentan formulas de trabajo probadas.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_dominancia_cerebral = document.getElementById("dominancia_cerebral");
  var donut_dominancia_cerebral = new Chart(ctx_dominancia_cerebral, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [100, 0],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      tooltips: {
        enabled: false
      }
    }
  });
});
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="motivacion"></canvas><h1>88.7%</h1></div></div><div class="col-md-6"><h3>Motivación</h3><p class="inter">Interpretación</p><p>Le interesa buscar la verdad y el conocimiento. Es curioso por naturaleza. Desea aprender, actualizarse en los conocimientos y desarrollarse en su area de competencia. Es un intelectual que lleva sus proyectos al plano cientifico, a la investigacion y al estudio profundo de los problemas. Se orienta hacia la gente interesandose en sus cualidades, sentimientos y problemas. La ayuda desinteresada hacia la gente constituye su principio basico. Aprecia lo útil o lo practico, pues es pragmatico. Las cosas valen por su utilidad, no por su belleza. Prefiere el trabajo.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_motivacion = document.getElementById("motivacion");
  var donut_motivacion = new Chart(ctx_motivacion, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [88.7, 11.299999999999997],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      tooltips: {
        enabled: false
      }
    }
  });
});
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="valores"></canvas><h1>100%</h1></div></div><div class="col-md-6"><h3>Valores</h3><p class="inter">Interpretación</p><p>Le interesa buscar la verdad y el conocimiento. Es curioso por naturaleza. Desea aprender, actualizarse en los conocimientos y desarrollarse en su area de competencia. Es un intelectual que lleva sus proyectos al plano cientifico, a la investigacion y al estudio profundo de los problemas. Se orienta hacia la gente interesandose en sus cualidades, sentimientos y problemas. La ayuda desinteresada hacia la gente constituye su principio basico. Aprecia lo útil o lo practico, pues es pragmatico. Las cosas valen por su utilidad, no por su belleza. Prefiere el trabajo subordinado a la direccion y el mando de un superior. Gusta del trabajo participativo.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_valores = document.getElementById("valores");
  var donut_valores = new Chart(ctx_valores, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [100, 0],
        backgroundColor: ["#357DD5", "#D1EAFF"]
      }]
    },
    options: {
      cutoutPercentage: 85,
      legend: { display: false },
      tooltips: {
        enabled: false
      }
    }
  });
});
</script></div></section><section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/></figure><h1>Palabras Descriptivas</h1></article><div class="tags"><ul id="my-list"><li>Temeroso</li><li>Inquieto</li><li>Atractivo</li><li>Humilde</li><li>Rebelde</li><li>Moderado</li><li>Odediente</li><li>Complaciente</li></ul></div></div></section><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Descripción Detallada del Perfil</h1><hr/></article><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="empuje"></canvas><h1>69%</h1></div></div><div class="col-md-6"><h3>Empuje</h3><p class="inter">Interpretación</p><p>Las personas con empuje ALTO, se caracterizan por las siguientes conductas: Da resultados,Apresura la accion,Acepta retos,Se aventura en lo desconocido,Toma decisiones,Cuestiona el status,Toma autoridad,Da soluciones,Reduce costos,Resuelve problemas.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_empuje = document.getElementById("empuje");
  var donut_empuje = new Chart(ctx_empuje, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [69, 31],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="influencia"></canvas><h1>93%</h1></div></div><div class="col-md-6"><h3>Influencia</h3><p class="inter">Interpretación</p><p>Las personas con influencia ALTA, se caracterizan por las siguientes conductas:Estableciendo contratos,Impresionando favorablemente,Hablando con soltura</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_influencia = document.getElementById("influencia");
  var donut_influencia = new Chart(ctx_influencia, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [93, 7],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="estabilidad"></canvas><h1>90%</h1></div></div><div class="col-md-6"><h3>Estabilidad</h3><p class="inter">Interpretación</p><p>Tiene preferencia por el trabajo estructurado, que impliquen bajo riesgo, y que este basado en politicas, procedimientos y procesos claramente definidos. Es habil en el uso de informacion concreta, que usa para su solucion de problemas. Tiende a ser poco emocional, prefiere mantenerse objetivo ante las situaciones que se presenten. Tiene habilidad para el trabajo que implica la ejecucion de tareas secuenciales y ordenadas de manera logica.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_estabilidad = document.getElementById("estabilidad");
  var donut_estabilidad = new Chart(ctx_estabilidad, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [90, 10],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="apego"></canvas><h1>80%</h1></div></div><div class="col-md-6"><h3>Apego a normas</h3><p class="inter">Interpretación</p><p>Tiene preferencia por el uso de informacion concreta. Tiene la capacidad de relacionarse con la gente, pero mantendra un tono mas bien superficial en la relacion. Evitara relacionarse a un nivel emocional profundo.  Es posible que se sienta incomoda ante las demostraciones de efusividad.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_apego = document.getElementById("apego");
  var donut_apego = new Chart(ctx_apego, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [80, 20],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="social"></canvas><h1>85%</h1></div></div><div class="col-md-6"><h3>Social</h3><p class="inter">Interpretación</p><p>Este valor implica un sentimiento altruista por la gente, tanto por extraños como conocidos. Estos individuos buscan desinteresadamente mejorar el bienestar de otros sirviendoles. Ellos procuran ayudar a toda clase de personas: quienes estan aventajados y/o se sientan maltratados. Sus simpatizantes son impulsados a la accion por un sentido de justicia social. Sus juicios son objetivos, matizados de emociones e idealismo. La indignacion social frecuentemente causa conflictos con el individuo de valores economicos.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_social = document.getElementById("social");
  var donut_social = new Chart(ctx_social, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [85, 15],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="justicia"></canvas><h1>89%</h1></div></div><div class="col-md-6"><h3>Justicia</h3><p class="inter">Interpretación</p><p>Estos individuos no tienen claro cuales son los derechos de los demas, ni hasta donde esta permitido llegar con sus acciones. Estas personas tienen otros tipos de intereses que opacan el valor de justicia, por este motivo es que pueden llegar a mostrar anti-valores de manera inconsciente. Tienden a parecer como injustos porque buscan lograr los objetivos sin importar los medios, esta situacion los puede llevar a ser vistos como personas frias y calculadoras, tienden a mostrar un liderazgo autocratico.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_justicia = document.getElementById("justicia");
  var donut_justicia = new Chart(ctx_justicia, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [89, 11],
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
</script><article class="interior"><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="regulatorio"></canvas><h1>79%</h1></div></div><div class="col-md-6"><h3>Regulatorio</h3><p class="inter">Interpretación</p><p>Estos individuos buscan identificarse con una fuerza reconocida  por el bien o gobernar sus vidas por un codigo de conducta que prometera aprobacion o aceptacion por una alta autoridad. Buscan unidad en su propio cosmos y una relacion con esa totalidad. Lo “correcto” o “incorrecto” es importante para ellos y tiende a hacer juicios morales de conformidad con ellos. Quieren estar en lo “correcto”. Generalmente ellos tienden a ser cooperativos, controlados y a observar estandares establecidos.</p></div></div><hr/></article><script type="text/javascript">$(document).ready(function () {
  var ctx_regulatorio = document.getElementById("regulatorio");
  var donut_regulatorio = new Chart(ctx_regulatorio, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [79, 21],
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
</script></div></section><section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/></figure><h1>¿Qué lo motiva?</h1><ul id="que-lo-motiva"><li>Poder,autoridad</li><li>Posicion y prestigio</li><li>Dinero y cosas materiales</li><li>Retos</li><li>Oportunidad de avance</li><li>El saber</li><li>Amplio margen para operar</li><li>Oportunidades de avance</li><li>Respuestas directas</li><li>Libertad de controles, supervision y detalle</li><li>Eficiencia en la operacion</li><li>Actividades variadas</li><li>Popularidad y reconocimiento social</li><li>Recompensas monetarias para mantener su ritmo de vida</li><li>Reconocimiento público que indique su habilidad</li><li>Libertad de palabra y personas con las cuales hablar</li><li>Condiciones favorables de trabajo</li><li>Actividades con gente fuera del trabajo</li><li>Relaciones democráticas</li><li>Libertad de control y detalles</li><li>Ingreso psicológico</li><li>Identificación con la compañía</li></ul></article></div></section><section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/></figure><h1>Necesita</h1><ul id="necesita"><li>Compromisos negociados de igual a igual</li><li>Identificion con la compañia</li><li>Tareas dificiles</li><li>Desarrollar valores intrínsecos</li><li>Aprender a tomar su paso y relajarse</li><li>Saber los resultados esperados</li><li>Entender a las personas</li><li>Enfoque lógico</li><li>Empatía</li><li>Técnicas basadas en experiencias prácticas</li><li>Conciencia de que las sanciones existen</li><li>Sacudidas ocasionales</li><li>Control de su tiempo</li><li>Objetividad</li><li>Énfasis en la utilidad de la empresa</li><li>Un supervisor democrático  con quien pueda asociarse</li><li>Presentarlo con gente influyente</li><li>Control emocional</li><li>Sentido de urgencia</li><li>Control de su desempeño por proyectos</li><li>Confianza en el producto</li><li>Datos analizados</li><li>Administración financiera personal</li><li>Supervisión más estricta</li></ul></article></div></section><section><div class="container"><article><figure><img src="styles/images/circulo.PNG"/></figure><h1>Limitaciones</h1><ul id="limitaciones"><li>Excederse en sus prerrogativas.</li><li> Actuar intrepidamente.</li><li>Ser cortante sarcastico con los demas.</li><li>Mostrarse impaciente y descontento con el trabajo de rutina.</li><li>Inspirar temor en los demás.</li><li>Imponerse a la gente.</li><li>Malhumorarse cuando no tiene el primer lugar.</li><li>Ser crítico y buscar errores.</li><li>Descuidar los detalles.</li><li>Resistirse a participar  como parte de un grupo.</li><li>Preocuparse más de su popularidad que de los resultados tangibles</li><li>Ser exageradamente persuasivo</li><li>Actuar impulsivamente, siguiendo su corazón en lugar de su inteligencia</li><li>Tomar decisiones basado en análisis superficiales</li><li>Ser poco realista al evaluar a las personas</li><li>Ser descuidado con los detalles</li><li>Confiar en las personas indiscriminadamente</li><li>Ser superficial</li></ul></article></div></section><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Gráfico Radar de Competencias</h1><hr/><div class="row"><div class="col-md-6"><div class="contenedor__grafica"><canvas id="tendencia_radar"></canvas></div></div></div></article></div></section><script type="text/javascript">$(document).ready(function () {
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
});</script><section><div class="container"><article class="basica"><figure><img src="styles/images/circulo.PNG"/></figure><h1>Tendencia de Comportamiento</h1><hr/><div class="row"><div class="col-md-12"><div class="contenedor__grafica__range"><!--div.valorp Resultado
div.resultado
p Deseado
div.deseado--><div class="range__contenedor"><div class="range__row"><p>Pacifico</p><input id="resultado" type="range" value="#{ tendencia.I }"/><p>Energico</p><!--input(id="deseado" type="range" value="80")--></div><div class="range__row"><p>Introvertido</p><input id="resultado" type="range" value="#{ tendencia.II }"/><p>Extrovertido</p><!--input(id="deseado" type="range" value="90")--></div><div class="range__row"><p>Flexible</p><input id="resultado" type="range" value="#{ tendencia.III }"/><p>Metódico</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Visceral</p><input id="resultado" type="range" value="#{ tendencia.V }"/><p>Racional</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Original</p><input id="resultado" type="range" value="#{ tendencia.VI }"/><p>Ordenado</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Analítico</p><input id="resultado" type="range" value="#{ tendencia.VII }"/><p>Emocional</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Concreto</p><input id="resultado" type="range" value="#{ tendencia.VIII }"/><p>Conceptual</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Empírico</p><input id="resultado" type="range" value="#{ tendencia.IX }"/><p>Estudioso</p><!--input(id="deseado" type="range" value="80")--></div><div class="range__row"><p>Práctico</p><input id="resultado" type="range" value="#{ tendencia.XI }"/><p>Creativo</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Exigente</p><input id="resultado" type="range" value="#{ tendencia.XII }"/><p>Comprensivo</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Parcial</p><input id="resultado" type="range" value="#{ tendencia.XIII }"/><p>Imparcial</p><!--input(id="deseado" type="range" value="50")--></div><div class="range__row"><p>Liberal</p><input id="resultado" type="range" value="#{ tendencia.XV }"/><p>Conservador</p><!--input(id="deseado" type="range" value="50")--></div></div></div></div></div></article></div></section></main><footer></footer></body></html>