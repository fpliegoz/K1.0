<!--#include file="./graficasFlash.asp"-->
<!--#include file="./ED_IntegralReportes.asp"-->
<!--#include file="./competenciaResultadosClass.asp"-->
<!--#include file="./khorPruebaFormat.asp"-->
<!--#include file="./prueba7str.asp"-->
<%
  '-- Etiquetas provisionalmente aqui hasta que se libere versión final
  '-- Deben terminar en un archivo que se incluya en khorLabel.asp o en una seccion de ese archivo, para que puedan personalizarse en khorLabelEspecial.asp
  rempLBL_Titulo = "Reporte Integral"
  rempLBL_SubTitFecha = "De la evaluación presentada el %%0"
  rempMF_TipoCompatibilidad = "1:Reclutamiento y Selección|2:Desaarrollo de Talento"
  
  rempLBL_Introduccion = "Introducci&oacute;n"
  rempLBL_MensajeIntroduccion = "<p>KHOR es una herramienta de evaluaci&oacute;n robusta que permite conocer sus &aacute;reas de oportunidad y fortalezas, en aspectos relacionados con el trabajo y su vida personal. Los instrumentos contenidos han pasado por rigurosos ex&aacute;menes para asegurar su validez, confiabilidad y estandarizaci&oacute;n en poblaci&oacute;n mexicana.</p>" & _
                            "<p>Si usted ha seguido las instrucciones durante la fase de evaluaci&oacute;n, usted notar&aacute; que este reporte lo describe a profundidad, en distintos aspectos como competencias, habilidades, estilos de trabajo, y preferencias personales.</p>" & _
                            "<p>Hay algunas situaciones que pueden incidir en una evaluaci&oacute;n poco v&aacute;lida. Es decir, que no lo describa con exactitud. Si usted present&oacute; la evaluaci&oacute;n en un estado emocional alterado, o bajo condiciones de estr&eacute;s, o con distracciones frecuentes, es probable que su rendimiento no haya sido el esperado. De la misma manera, si usted respondi&oacute; de manera poco sincera a las preguntas que se le hicieron, notar&aacute; que sus resultados pueden ser contradictorios.</p>" & _
                            "<p>Recomendamos que lea este reporte con calma, y que identifique situaciones de su vida en las que usted manifiesta lo aqu&iacute; descrito. M&aacute;s importante a&uacute;n, l&eacute;alo con mente abierta, dispuesto a identificar &aacute;reas de mejora y a trazar un plan individual de desarrollo.</p>" & _
                            "<p class=""text-right""><strong>Algunas preguntas que usted debe tener en mente al leer este reporte, son:</strong></p>" & _
                            "<ul>" & _
                            "<li>&iquest;C&oacute;mo puedo usar esta informaci&oacute;n a mi favor?</li>" & _
                            "<li>&iquest;C&oacute;mo puedo aprovechar estas habilidades?</li>" & _
                            "<li>&iquest;Qu&eacute; cosas debo corregir para potencializar mi desarrollo?</li>" & _
                            "<li>&iquest;Qu&eacute; pasara si cambiara este rasgo o atributo indeseable? &iquest;Qu&eacute; ganar&iacute;a?</li>" & _
                            "</ul>" & _
                            "<p class=""text-right"">Atentamente,</p>" & _
                            "<p><h2 class=""text-right"">KHOR</h2></p>"

  rempLBL_CompatibilidadGeneral = "Compatibilidad General"
  
  rempLBL_CompatibilidadEspecifica = "Compatibilidad Especifica"
  rempLBL_Prueba = "Meta-Factor"

  rempLBL_MetaFactoresEnResumen = "Meta-factores, en Resumen"

  rempLBL_PalabrasDescriptivas = "Palabras descriptivas"
  remp_num_palabras = 8

  rempLBL_DescripcionDetalladaPerfil = "Descripci&oacute;n Detallada del Perfil"

  rempLBL_QueLoMotiva = "¿Qu&eacute; lo motiva?"
  rempLBL_Necesita = "Necesita"
  rempLBL_Limitaciones = "Limitaciones"

  rempLBL_GraficoRadarDeCompetencias = "Gr&aacute;fico Radar de Competencias"
  
  rempLBL_TendenciaDeComportamiento = "Tendencia de comportamiento"

  '----------------------------------------
  '----------------------------------------

  pf_excepcionescampos = "'p48_S','p48_T','p49_I','p49_J','p68_ResultadoGeneral'" '-- Indices de V-AV de IVO y Resultado general de Liderazgo
  pf_p16_campo2index = "p16_SucesosDeLaVida:1|p16_PresionesDeTrabajo:2|p16_PresionesPersonales:3|p16_ConcienciaEmocionalDeMiMismo:4|p16_ExpresionEmocional:5|p16_ConcienciaEmocionalDeOtros:6|p16_Intencion:7|p16_Creatividad:8|p16_Elasticidad:9|p16_ConexionesInterpersonales:10|p16_DescontentoConstructivo:11|p16_Compasion:12|p16_Perspectiva:13|p16_Intuicion:14|p16_RadioDeConfianza:15|p16_PoderPersonal:16|p16_Integridad:17|p16_SaludGeneral:18|p16_CalidadDeVida:19|p16_CocienteDeRelaciones:20|p16_OptimoRendimiento:21|"

  dim rep : set rep = new clsCompatibilidadReporte  '-- Objeto general de reportes de compatibilidad y pruebas
  dim compatibilidad : compatibilidad = 0
  dim idperfil : idperfil = 0
  dim pctipo : pctipo = reqn("pctipo") : if pctipo=0 then pctipo=CR_RECYSEL
  dim escenario: escenario = khorEscenarioComportamiento(0)
  dim listaPruebasAplicadas : listaPruebasAplicadas = ""
  
'================================================================================'

  class clsReporteFactor
    public IdPrueba
    public IdPruebaFactor
    public PruebaFactor
    public CampoVistaPersona
    public ResultadoMaximo
    public ResultadoPersona
    public Interpretacion
    
    public property get NombreFactor()
      NombreFactor = PruebaFactor '-- Para conversiones si es preciso
    end property
    
    private sub class_initialize()
    end sub
    private sub class_terminate()
    end sub
  end class
  
  '----------------------------------------
  
  sub remp_inicializa()
    escenario = khorEscenarioComportamiento(IdPerfil)
    compatibilidad = rep.compatibilidadTotal(IdPersona,IdPerfil,pctipo)
    rep.initPsi IdPersona, IdPerfil
  end sub
  
  sub remp_finaliza()
    rep.clean
    set rep = nothing
  end sub
  
  '----------------------------------------
    
  function remp_interpretacion(idprueba)
    dim interpretacion : interpretacion = ""
    if inCSV("46,47,50,51,68,72,73,74,75,76,77,78,79,81", cstr(idprueba)) >= 0 then 'interpretacion por factores, se excluye
      interpretacion = lblFRS_NoHayDetalleDisponible  '-- Provisional
    else
      dim personaprueba : set personaprueba = new clsPersonaPrueba
      personaprueba.getfromDB conn, idpersona, idprueba
      interpretacion = personaprueba.interpretacion(1+escenario,false,true) '-- interpretacion "corta"
      set personaprueba = nothing
    end if
    if interpretacion<>"" then
      '-- Quita titulos de las interpretaciones, y sustituye <br> por espacio en las interpretaciones
      int_tit_regexp = "<span class=""khorInt_Titulo1"" style=""page-break-inside:avoid;"">[^>]+>"
      interpretacion = ereg_replace(interpretacion, int_tit_regexp, "", true)
      int_tit_regexp = "<span class=""khorInt_Titulo2"" style=""page-break-inside:avoid;"">[^>]+>"
      interpretacion = ereg_replace(interpretacion, int_tit_regexp, "", true)
      int_tit_regexp = "</br>"
      interpretacion = ereg_replace(interpretacion, int_tit_regexp, " ", true)
      int_tit_regexp = "<br>"
      interpretacion = ereg_replace(interpretacion, int_tit_regexp, " ", true)
    end if
    remp_interpretacion = interpretacion
  end function
  
  '----------------------------------------
  
  function remp_palabrasdescriptivas()
    dim lpreg : lpreg = "" '-- Lista de todas las palabras marcadas como "L" en la prueba
    dim sq, rs, page, grupo, lgrupo, opcion, nopcion, npreg
    for page=1 to 4
      sq = "SELECT * FROM resCleaver WHERE IdPersona = "& IdPersona & " AND pagina = "& page
      set rs=getrs(conn,sq)
      if not rs.EOF then
        FOR grupo = 1 to 3
          lgrupo = rsNum(rs,"L"&grupo)
          for opcion = 1 to 4
            nopcion = ((grupo - 1) * 4) + opcion
            npreg = ((page - 1) * 24) + nopcion
            if cstr(lgrupo) = cstr(nopcion) then
              lpreg = strAdd( lpreg, ",", npreg )
            end if
          next
        next
      end if
      rs.close
      set rs=nothing
    next
    remp_palabrasdescriptivas = lpreg
  end function
  
  '----------------------------------------
  
  sub remp_InterpretacionFactores(col,tipo)
    dim i, auxo
    col.clean
    dim tabla, tablacondicion
    if tipo = 0 then
      tabla = "pruebaInterpretacion"
      tablacondicion = ""
    else
      tabla = "pruebaInterpretacionAlt"
      tablacondicion = " AND (TipoInterpretacion = " & tipo & ")"
    end if
    '-- Obtiene factores
    dim sq : sq = "SELECT SecuenciaGrupo, Secuencia, p.IdPrueba, pf.IdPruebaFactor, PruebaFactor, CampoVistaPersona, MAX(ValorMaximo) AS maxRes" & _
                  " FROM vPrueba p INNER JOIN PruebasFactores pf ON p.idprueba = pf.IdPrueba" & _
                  " INNER JOIN " & tabla & " pi ON pf.IdPruebaFactor = pi.IdPruebaFactor" & tablacondicion & _
                  " WHERE p.Activa = 1 AND CampoVistaPersona<>'' AND p.IdPrueba IN (" & listaPruebasAplicadas & ")" & _
                  " AND CampoVistaPersona NOT IN (" & pf_excepcionescampos & ")" & _
                  " GROUP BY SecuenciaGrupo, Secuencia, p.IdPrueba, pf.IdPruebaFactor, PruebaFactor, CampoVistaPersona" & _
                  " ORDER BY SecuenciaGrupo, Secuencia, pf.IdPruebaFactor"
    dim rs : set rs = getrs(conn,sq)
    dim lstp : lstp = " FROM Personal"
    sq = "SELECT Personal.IdPersonal"
    while not rs.eof
      set auxo = new clsReporteFactor
      auxo.IdPrueba = rsNum(rs,"IdPrueba")
      auxo.IdPruebaFactor = rsNum(rs,"IdPruebaFactor")
      auxo.PruebaFactor = rsStr(rs,"PruebaFactor")
      auxo.CampoVistaPersona = rsStr(rs,"CampoVistaPersona")
      auxo.ResultadoMaximo = rsNum(rs,"maxRes")
      if auxo.ResultadoMaximo = 0 then  auxo.ResultadoMaximo = 100          
      auxo.ResultadoPersona = null
      auxo.Interpretacion = ""
      col.add auxo, auxo.IdPruebaFactor
      sq = sq &  ", " & auxo.CampoVistaPersona
      if instr(lstp,"repPrueba"&auxo.IdPrueba) <= 0 then
        lstp = lstp & " LEFT JOIN repPrueba" & auxo.IdPrueba & " ON Personal.IdPersonal = repPrueba" & auxo.IdPrueba & ".IdPersonal"
      end if
      rs.movenext
    wend
    rs.close
    set rs = nothing
    '-- Obtiene resultados e interpretaciones
    sq = sq & lstp & " WHERE Personal.IdPersonal = " & idpersona
    set rs = getrs(conn,sq)
    if not rs.eof then
      for i=1 to col.count
        set auxo = col.obj(i)
        if not isnull(rs(auxo.CampoVistaPersona)) then
          auxo.ResultadoPersona = rsNum(rs,auxo.CampoVistaPersona)
          if auxo.IdPrueba = 16 then auxo.ResultadoPersona = CalculaValorGraficaEQ( auxo.ResultadoPersona, descFromMenuFijo(pf_p16_campo2index, auxo.CampoVistaPersona) )
          sq = "SELECT Interpretacion FROM " & tabla & " WHERE (IdPruebaFactor = " & auxo.IdPruebaFactor & ")" & tablacondicion & _
              " AND (ValorMinimo <= " & auxo.ResultadoPersona & ") AND (" & auxo.ResultadoPersona & " <= ValorMaximo)"
          auxo.Interpretacion = getBD("Interpretacion",sq)
        end if
      next
    end if
    rs.close
    set rs = nothing
    set auxo = nothing
  end sub
  
  '----------------------------------------
  
  function remp_InterpretacionFactoresAlt(tipo)
    dim retval : retval = ""
    dim colAux : set colAux = new frsCollection
    remp_InterpretacionFactores colAux, tipo
    if (tipo = 1) AND (colAux.count > 0) then  '-- que lo motiva
      '-- Elije unicamente la del resultado mayor y la del resultado menor
      '-- Falta definir que hacer con empates
      dim rmenor : rmenor = 1
      dim rmayor : rmayor = 1
      for i=1 to colAux.count
        if colAux.obj(i).ResultadoPersona > colAux.obj(rmayor).ResultadoPersona then rmayor = i
        if colAux.obj(i).ResultadoPersona < colAux.obj(rmenor).ResultadoPersona then rmenor = i
      next
      retval = colAux.obj(rmayor).Interpretacion
      if rmayor <> rmenor then retval = strAdd( retval, vbCRLF, colAux.obj(rmenor).Interpretacion )
    else  '-- todos los que apliquen
      retval = ""
      for i=1 to colAux.count
        retval = strAdd( retval, vbCRLF, colAux.obj(i).Interpretacion )
      next
    end if
    colAux.clean
    set colAux = nothing
    remp_InterpretacionFactoresAlt = retval
  end function
  
  '----------------------------------------

  '-- Logica copiada de la funcion edri_PruebasResumen de repPersonaPruebasResumen.asp
  sub remp_FactoresOpuestos(colFactores)
    dim lista_pruebas_posibles : lista_pruebas_posibles = "7,5,4,48,6,12,66,46,2,25,27,30,9,8,11"
    dim IdPuesto : IdPuesto = IdPerfil
    '-- Escenario de comportamiento
    dim auxdisc : auxdisc = iif( escenario=0, "Cotidiana", iif( escenario=1, "Motivante", "BajoPresion" ) )  '--- NO TRADUCIR: nombre de campo en vista
    '-- Determina las pruebas a incluir en el reporte
    dim sq : sq = "SELECT p.IdPrueba, Prueba, PruebaGrupo FROM vPrueba p"
    if IdPuesto > 0 then  '-- Las del puesto
      sq = sq & " INNER JOIN PuestosPruebas pp ON p.idprueba = pp.idprueba AND pp.idpuesto = " & IdPuesto
    else  '-- Las que haya aplicado la persona
      sq = sq & " INNER JOIN vPersonaPruebaReciente pp ON p.idprueba = pp.idprueba AND pp.idpersonal = " & IdPersona
    end if
    sq = sq & " WHERE p.idprueba in (" & lista_pruebas_posibles & ")" & _
              " ORDER BY SecuenciaGrupo, Secuencia, Prueba"
    dim colPruebas : set colPruebas = new frsCollection
    colPruebas.keyDescFromQuery "IdPrueba", "Prueba", "PruebaGrupo", sq
    if colPruebas.count > 0 then
      dim rs, objp, objf
      '-- Lee los factores de cada prueba
      sq = "SELECT pf.* FROM vprueba p INNER JOIN pruebasfactores pf ON p.idprueba = pf.idprueba" & _
          " WHERE p.idprueba in (" & colPruebas.keyList() & ")" & _
          " AND ((p.idprueba<>66 AND posperfil>0) OR (p.idprueba=66 AND posperfil=0))" & _
          " ORDER BY SecuenciaGrupo, Secuencia, Prueba" & _
          ", (CASE WHEN p.idprueba=48 THEN CampoVistaPerfil ELSE RIGHT('000000'+IdPruebaFactor,6) END)"
      set rs = getrs(conn,sq)
      while not rs.eof
        set objf = new clsPruebaFactor
        objf.getFromRS rs
        '-- Excepciones
        if objf.idprueba = 7 then '-- Comportamiento: en persona usa campo correspondiente al escenario
          objf.CampoVistaPersona = objf.CampoVistaPerfil & "_" & auxdisc
        elseif objf.idprueba = 66 then '-- Test de roles y Necesidades: factores calculados no se perfilan
          objf.CampoVistaPerfil = objf.CampoVistaPersona
        elseif incsv("8,11,25,27,30",objf.idprueba) >= 0 then '-- Un solo factor: usa nombre de la prueba
          set objp = colPruebas.objByKey( objf.IdPrueba )
          objf.PruebaFactor = objp.Desc
        end if
        '-- Agrega a colección
        colFactores.add objf, objf.IdPruebaFactor
        rs.movenext
      wend
      rs.close
      set rs = nothing
      '-- Armado de querys para persona y perfil
      dim sqper : sqper = "SELECT per.IdPersonal"
      dim sqpue : sqpue = "SELECT pue.IdPuesto"
      dim auxper : auxper = " FROM Personal per"
      dim auxpue : auxpue = " FROM Puestos pue"
      dim lblAbreviaturas : lblAbreviaturas = ""
      dim p, f
      for f=1 to colFactores.count
        set objF = colfactores.obj(f)
        if p <> objF.IdPrueba then
          p = objF.IdPrueba
          auxper = auxper & " LEFT JOIN repPrueba" & p & " ON per.IdPersonal = repPrueba" & p & ".IdPersonal"
          auxpue = auxpue & " LEFT JOIN rpPrueba" & p & " ON pue.IdPuesto = rpPrueba" & p & ".IdPuesto"
        end if
        sqper = sqper & ", " & objF.CampoVistaPersona
        sqpue = sqpue & ", " & objF.CampoVistaPerfil
        objF.valPersona = ""
        objF.valPerfil = ""
        lblAbreviaturas = strAdd( lblAbreviaturas, ",", f )
      next
      sqper = sqper & auxper & " WHERE per.IdPersonal = " & IdPersona
      sqpue = sqpue & auxpue & " WHERE pue.IdPuesto = " & IdPuesto
      '-- Lectura de resultados: persona
      auxper = ""
      set rs = getrs(conn,sqper)
      if not rs.eof then
        for f=1 to colFactores.count
          set objF = colfactores.obj(f)
          if not isnull(rs(objF.CampoVistaPersona)) then  '-- Se usa asi en lugar de rsNum por el armado de la grafica
            objF.valPersona = cdbl(rs(objF.CampoVistaPersona))
            '-- Conversiones y estandarizacion a la escala de este reporte
            if objF.IdPrueba = 2 then '-- Cognicion
              objF.valPersona = khorNivelThermanSeccion( objF.valPersona, objF.IdPruebaFactor - (objF.IdPrueba * 100) )
              objF.valPersona = round( 100 * objF.valPersona / 7 )
            elseif objF.IdPrueba = 9 then '-- Habilidades operativas
              objF.valPersona = khorNivelHO( objF.valPersona, objF.IdPruebaFactor - (objF.IdPrueba * 100) )
              objF.valPersona = round( 100 * objF.valPersona / 7 )
            elseif objF.IdPrueba = 6 then '-- Estilo Gerencial
              objF.valPersona = round( 100 * objF.valPersona / 36 )
            elseif objF.IdPrueba = 25 then '-- Inteligencia basica I
              objF.valPersona = round( 100 * objF.valPersona / 4 )
            elseif objF.IdPrueba = 48 OR objF.IdPrueba = 66 then '-- IVO | Roles y Necesidades
              objF.valPersona = round( 100 * objF.valPersona / 9 )
            elseif objF.IdPrueba = 12 OR objF.IdPrueba = 46 then '-- PPV | EQ
              objF.valPersona = objF.valPersona * 10
            end if
          end if
          auxper = auxper & "," & objF.valPersona
        next
      end if
      if auxper<>"" then auxper = mid(auxper,2)
      rs.close
      set rs = nothing
      '-- Lectura de resultados: perfil
      if IdPuesto > 0 then
        set rs = getrs(conn,sqpue)
        if not rs.eof then
          for f=1 to colFactores.count
            set objF = colfactores.obj(f)
            if not isnull(rs(objF.CampoVistaPerfil)) then  '-- Se usa asi en lugar de rsNum por el armado de la grafica
              objF.valPerfil = cdbl(rs(objF.CampoVistaPerfil))
              '-- Conversiones y estandarizacion a la escala de este reporte
              if objF.IdPrueba = 2 OR objF.IdPrueba = 9 then '-- Cognicion | Habilidades operativas
                objF.valPerfil = round( 100 * objF.valPerfil / 7 )
              elseif objF.IdPrueba = 6 then '-- Estilo Gerencial
                objF.valPerfil = round( 100 * objF.valPerfil / 36 )
              elseif objF.IdPrueba = 25 then '-- Inteligencia basica I
                objF.valPerfil = round( 100 * objF.valPerfil / 4 )
              elseif objF.IdPrueba = 48 OR objF.IdPrueba = 66 then '-- IVO | Roles y Necesidades
                objF.valPerfil = round( 100 * objF.valPerfil / 9 )
              elseif objF.IdPrueba = 46 then '-- EQ
                objF.valPerfil = objF.valPerfil * 10
              elseif objF.IdPrueba = 8 OR objF.IdPrueba = 11 then '-- Ingles | Ortografia
                objF.valPerfil = ""
              end if
            end if
          next
        end if
        rs.close
        set rs = nothing
      end if
      set objp = nothing
      set objf = nothing
    end if
    colPruebas.clean
    set colPruebas = nothing
  end sub
  
'================================================================================'
%>