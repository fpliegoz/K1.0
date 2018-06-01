<%
  lblCOBO_BienvenidaEncuesta = ""'en coboAplica.asp se muestra una bienvenida si el cliente lo requiere (se modifica en khorLabelEspecial.asp)
  lblCOBO_Encuesta = "Encuesta"
  
  lblCobo_ExcluirSeleccionados = "Excluir seleccionados"
  lblCobo_ConfirmacionExcluirPersonasExcluirRespuestas = "Al excluir " & lblKHOR_Persona_s & ", sus respuestas se eliminar�n tambi�n. �Desea continuar?"
  lblCobo_TodasLasQueRespondieron = "Todas las que respondieron"

  lblCOBO_DatoDemografico = "Dato demogr&aacute;fico"
  lblCOBO_DatosDemograficos = "Datos demogr&aacute;ficos"
  lblCOBO_NoHayDatosDem = "La encuesta seleccionada no tiene asociados " & lblCOBO_DatosDemograficos & "."
  
  lblCOBO_Poblacion = lblFRS_Poblacion  'Poblaci&oacute;n
  lblCOBO_PoblacionInsuficiente = "No hay suficente poblaci&oacute;n para calcular."

  lblCOBO_Factor = "Factor"
  lblCOBO_Factores = "Factores"
  lblCOBO_Factor_es = "Factor(es)"
  lblCOBO_FactoresSeleccionados = "Factores Seleccionados"
  
  '--- Menu principal
  lblCOBO_Factor = "Elemento"
  lblCOBO_Factores = "Elementos"
  lblCOBO_Factor_es = "Elemento(s)"
  lblCOBO_FactoresSeleccionados = "Elementos Seleccionados"
  lblCOBO_Aplicar = "Aplicar Cuestionario"
  lblCOBO_AplicarAdmin = "Captura de Respuestas"
  lblCOBO_Asignacion = "Registro de Participantes"
  lblCOBO_RepGeneral = "Reporte de Resultados"
  lblCOBO_RepComparativo = "Comprobaci�n de Hip�tesis"
  lblCOBO_RepComentarios = "Preguntas de Retroalimentaci�n"
  lblCOBO_Exportacion = lblFRS_Exportacion & " de datos"
  lblCOBO_Ingresar = "Aplicar Cuestionario de " & lblCOBO_ClimaOrganizacional & " de forma AN�NIMA"
  lblCOBO_OldRepGeneral = "(*) Reporte por Factor/Pregunta"
  lblCOBO_OldRepComparativo = "(*) Reporte de Comparaci�n"
  lblCOBO_RepDatoDem = "Reporte de Datos Demogr�ficos"
  lblCOBO_RepEstadisticoDeFactor = "Reporte Estad�stico"
  lblCOBO_RepResXDatoDem = "Reporte de Resultados por Dato Demogr�fico"
  
  '--- Variables de evaluaci�n
  lblCobo_Variable = "Variable"
  lblCOBO_SituacionActual = "Situaci�n Actual"
  lblCOBO_SituacionActual_Abrev = "S.A."
  lblCOBO_Importancia = "Importancia"
  lblCOBO_Importancia_Abrev = "Imp."

  '--- Aplicacion anonima
  lblCOBO_ClaveDeAccesoEncuesta = "Clave de Acceso a " & lblCOBO_Encuesta
  lblCOBO_NoHayPeriodosConClave = "La Clave ingresada no tiene periodos de evaluaci�n activos."
  lblCOBO_NoHayEncuesta = "No se encontr� " & lblCOBO_Encuesta & " disponible"

  lblCOBO_ClaveUsada = "La " & lblCOBO_ClaveDeAccesoEncuesta & " indicada ya esta en uso. Cambiela para continuar."
  
  lblCOBO_ClavesGeneradas = "Claves Generadas"
  lblCOBO_ClavesUtilizadas = "Claves Utilizadas"
  lblCOBO_NoClavesAGenerar = "N�mero de Claves a Generar"
  lblCOBO_ClaveYaUtilizada = "La clave ya ha sido utilizada"
  lblCOBO_ClaveInvalida = "La clave ingresada es inv�lida"

  '--- Aplicaci�n de la encuesta

  lblCOBO_InstruccionTitulo = "Bienvenido a la Evaluaci�n de " & lblCOBO_ClimaOrganizacional
  lblCOBO_InstruccionInicio = "<P class=""tjustify"">El objetivo de esta evaluaci�n, es encontrar las �reas de oportunidad que tenemos como empresa. Hacer esto, nos permitir� trabajar juntos en la mejora de nuestro ambiente de trabajo, as� como de aquellos elementos que contribuyen a que nuestro recurso m�s valioso, usted, se sienta satisfecho y orgulloso de trabajar con nosotros.</P>" & _
                              "<P class=""tjustify"">Se le presentar� a continuaci�n una serie de preguntas. Reflexione en ellas y seleccione las respuestas que mejor reflejen su punto de vista. Algunas preguntas se dirigen espec�ficamente a usted, la mayor�a se refieren a la empresa. Todas nos permitir�n conocer mejor la manera c�mo estamos funcionando.</P>" & _
                              "<P class=""tjustify"">En este cuestionario no hay respuestas correctas ni incorrectas. Por lo anterior, le pedimos que siga las instrucciones de cada secci�n del cuestionario y que reflexione en las preguntas que se le presentar�n y que las conteste con honestidad y tranquilidad. Sus respuestas ser�n manejadas con confidencialidad. Gracias de antemano por su cooperaci�n.</P>"
  lblCOBO_InstruccionParcial1 = "Si precisa interrumpir la evaluaci�n en alg�n momento, sus respuestas quedar�n guardadas y puede continuarla posteriormente."
  lblCOBO_InstruccionParcial2 = " Sin embargo, no podr� modificar las respuestas ingresadas en sesiones anteriores."
  lblCOBO_InstruccionFinal = ""
  lblCOBO_InstruccionFirma = lblFRS_GraciasParticipacion

  lblCOBO_InstruccionSAIM = "Ingrese las respuestas en """ & lblCOBO_SituacionActual & """ e """ & lblCOBO_Importancia & """ que mejor refleje su punto de vista."
  lblCOBO_InstruccionSA = "Ingrese la respuesta de """ & lblCOBO_SituacionActual & """ que mejor lo describa."
  lblCOBO_InstruccionComentarios = "Si tiene alg�n comentario acerca de esta evaluaci�n o en general, escribalo en este espacio."
  lblCOBO_InstruccionPregAbiertas = "Ingrese las respuestas a las siguientes preguntas."
  lblCOBO_InstruccionRankFactores = "Ordene los elementos seleccionando un n�mero en la lista de acuerdo a su grado de importancia." & _
                                    "<br/>Donde 1 es el m�s importante y %%0 es el menos importante. No debe repetir n�meros."
  lblCOBO_InstruccionDatosDemograficos = "Ingrese las respuestas a los siguientes datos demogr�ficos"
  lblCOBO_GuardadoParticipacion = lblKHOR_GraciasParticipacion & ", " & lblKHOR_SusRespuestasEstanSiendoProcesadas

  lblCOBO_NoHaRespondidoLasPreguntasAbiertasTerminar = "No ha respondido todas las preguntas abiertas opcionales. �Desea dar por terminada la evaluacion?"
  lblCOBO_NoHaRespondidoDatoDemograficoTerminar = "No ha respondido todas las preguntas demogr�ficas opcionales. �Desea dar por terminada la evaluacion?"

  lblCOBO_FolioParticipacion = "Folio de participaci�n"
  lblCOBO_ImprimirEncuesta = lblFRS_Imprimir & " " & lblCOBO_Encuesta

  '--- Escalas: deben tener el mismo numero de opciones
  menuFijoCOsa = "1:Completamente en desacuerdo|2:Parcialmente en desacuerdo|3:Neutro|4:Parcialmente de acuerdo|5:Completamente de acuerdo|"
  menuFijoCOim = "1:Nada importante|2:Poco importante|3:Neutro|4:Importante|5:Extremadamente importante|"

  '--- Recomendaci�n
  menuFijoCOBORecomendacionValor = "1:2.01|2:1.00|"
  menuFijoCOBORecomendacionBurnValor = ""
  menuFijoCOBORecomendacionLabel = "1:Se requiere atenci�n urgente|2:Se requiere atenci�n|"
  menuFijoCOBORecomendacionColor = "1:FF0000|2:FFFF00|"

  '--- Funcionamiento de las pantallas
  lblCobo_EstadoEncuesta = "Estado de " & lblCOBO_Encuesta
  lblCobo_NoHayEvalConCaracIndicadas = "No hay evaluaciones con las caracteristicas indicadas."
  lblCOBO_NoHayComentarios = "No hay comentarios registrados."
  lblCOBO_ParaAgregarPersonasSeleccionarEncuesta = "Para agregar personas debe seleccionar " & lblCOBO_Encuesta
  lblCOBO_Condiciones = "Condiciones"
  lblCOBO_LosEjesSonIguales = "Los dos ejes son iguales. Seleccione alguna diferencia entre ambos antes de Procesar"
  menuFijoExportacionCO = "0:Resultados por Grupos y " & lblCOBO_Factores & "|1:Respuestas a Reactivos (" & lblCOBO_SituacionActual & ")|2:Respuestas a Reactivos (" & lblCOBO_Importancia & ")|"
  
  '--- Reportes
  lblCOBO_UsarPromedio = "Usar " & lblFRS_Promedio
  lblCOBO_DivisionDeCuadrantes = "Divisi�n de cuadrantes"
  lblCOBO_ElValorDeLaUbicacionDebeSerEntre_X_y_ = "El valor de la ubicaci�n debe ser entre %%0 y "
  lblCOBO_MostrarMediaEnPorcentaje = "Mostrar media como porcentaje"
  lblECCO_GraficaGeneral = "Gr�fica General"
  lblECCO_DetalleXfactor = "Detalle por " & lblCOBO_Factor
  lblECCO_TituloGraficaComprobacion="Grafica de dispersi�n de personas. (Porcentaje de la poblacion%/Numero de Personas)"
  lblECCO_DatosSeleccionados="Datos Seleccionados"
  lblECCO_SatisfaccionTotal="Satisfacci�n Total"
  lblECCODetalleImp="Detalle de Importancia"
  
  
%>