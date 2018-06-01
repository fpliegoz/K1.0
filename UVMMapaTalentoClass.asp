<%

'================================================================================

CONST MT_PERFIL = 1
CONST MT_EJECUCION = 2

CONST MT_PRUEBAS = "MTCP"
CONST MT_COMPETENCIAS = "MTCC"

CONST MT_MCIE = "MCIE"
CONST MT_MCIP = "MCIP"
CONST MT_MCIT = "MCIT"
CONST MT_MCID = "MCID"
CONST MT_360 = "360"
CONST MT_EDL = "EDL"
CONST MT_EDD = "EDD"
CONST MT_EDP = "EDP"
CONST MT_EDK = "EDK"
CONST MT_EDS = "EDS"
CONST MT_EDE = "EDE"
CONST MT_ADP = "ADP"
CONST MT_ADJ = "ADJ"
CONST MT_EDT = "EDT"
CONST MT_EDATC = "EDATC"
CONST MT_EDASI = "EDASI"
CONST MT_EDASIO = "EDASIO"
CONST MT_EDBACH = "EDBACH"
CONST MT_EDTSU = "EDTSU"
CONST MT_EDFAC = "EDFAC"

function mapaAxisName(tipo)
  mapaAxisName = iif(tipo=MT_PERFIL,lblKHOR_MapaEje_Perfil,lblKHOR_MapaEje_Ejecucion)
end function

function getValueString(Kind, IdPeriodo, IdPersonal, IdPerfil, groupBy, idempresafiltro)
  dim mapa, modulosActivos, salida, condicion, sq, rs, valorPerfil, valorEjecucion
  set mapa = new mapaTalentoClass
  modulosActivos = khorModulosActivos()
  mapa.Initalize modulosActivos, true, IdPeriodo, true
  if Kind = "xml" then
    salida = scatterConfig() & "<data>"
  end if
  IF IdPeriodo<>0 AND IdPersonal<>"" THEN
    if IdPersonal="*" then
      condicion = strAdd( khorCondicionUsuario("Personal.IdPersonal"), " AND ", "(EXISTS(SELECT * FROM UVM_EvaluacionDimension e INNER JOIN UVM_TablaCRNS c ON c.IdCRN = e.IdCRN INNER JOIN Personal p ON CONVERT ( INT , REPLACE(p.Clave ,p.Prefijo ,'')) =c.Docente AND p.Clave = Personal.Clave  WHERE c.IdPeriodos="&idperiodo&"))") 
      if idempresafiltro > 0 then condicion = strAdd( "Personal.IdSucursal="&idempresafiltro, " AND ", condicion )
    else
      condicion = "Personal.IdPersonal IN (" & IdPersonal & ")"
    end if
    sq = mapa.query(idperfil,idperiodo,condicion,groupBy)
    set rs = getrs(conn,sq)
    while not rs.EOF
      valorPerfil = mapa.getValorRS(rs,MT_PERFIL)
      valorEjecucion = mapa.getValorRS(rs,MT_EJECUCION)
      if Kind = "xml" then
        salida = salida & "<point extraData='IdPersonal=" & rsNum(rs,"IdPersonal") & "' data='" & rsStr(rs,"Nombre") & "' x='" & khorFormatResultado(valorPerfil) & "' y='" & khorFormatResultado(valorEjecucion) & "' />"
      else
        salida = strAdd(salida, ":", khorFormatResultado(valorPerfil) & "," & khorFormatResultado(valorEjecucion))
      end if
      rs.movenext
    wend
    rs.close()
    set rs = nothing
  END IF
  if Kind = "xml" then
    salida = salida & "</data></scatter>"
  end if

  set mapa = nothing
  getValueString = salida
end function

'================================================================================

class mapaTalentoItem
  public id
  public tipo
  public desc
  public peso
  public total
  public total2
  public auxtotal
  public auxtotal2

  function itemConPerfil()
    itemConPerfil = (id = MT_PRUEBAS) OR (id = MT_COMPETENCIAS)
  end function
  
  function itemCols()
    itemCols = iif(id=MT_EDT,2,1)
  end function
  
  function itemXtraCols(rs,printtype)
    dim retval : retval = ""
    dim valor : valor = 0
    if id = MT_EDT then
      if isnull(rs("Resultado_" & id)) then
        retval = "-"
      else
        valor = rsNum(rs, "Resultado_" & id ) / 100
        auxtotal = auxtotal + cdbl(valor)
        auxtotal2 = auxtotal2 + cdbl(valor)
        retval = iif(printtype=3, valor, khorFormatPorcentaje(valor))
      end if
      retval = "<td align='center'>" & retval & "</td>"
    end if
    itemXtraCols = retval
  end function
  
  function tipoCR()
    dim retval : retval = 0
    select case id
      case MT_PRUEBAS
        retval = tipoCR_PSI
      case MT_COMPETENCIAS
        retval = tipoCR_COM
      case MT_MCIE
        retval = tipoCR_MCIE
      case MT_MCIP
        retval = tipoCR_MCIP
      case MT_MCIT
        retval = tipoCR_MCIT
      case MT_MCID
        retval = tipoCR_MCID
      case MT_360
        retval = tipoCR_360
      case MT_EDL
        retval = tipoCR_EDL
      case MT_EDD
        retval = tipoCR_EDD
      case MT_EDP
        retval = tipoCR_EDP
      case MT_EDK
        retval = tipoCR_EDK
      case MT_EDS
        retval = tipoCR_EDS
      case MT_EDE
        retval = tipoCR_EDE
      case MT_ADP
        retval = tipoCR_ADP
      case MT_ADJ
        retval = tipoCR_ADJ
      case MT_EDT
        retval = tipoCR_EDT
      case MT_EDATC
        retval = tipoCR_EDATC
      case MT_EDASI
        retval = tipoCR_EDASI
      case MT_EDASIO
        retval = tipoCR_EDASIO
      case MT_EDBACH
        retval = tipoCR_EDBACH
      case MT_EDTSU
        retval = tipoCR_EDTSU
      case MT_EDFAC
        retval = tipoCR_EDFAC
    end select
    tipoCR = retval
  end function

  function getCeldaRS(rsCelda,idperfil,printtype) '0: full, 1 & 2: export, 3: number
    dim retval, valor, idpersona
    if isnull(rsCelda("Resultado_" & id)) then
      retval = "-"
    else
      if id = MT_EDT then
        if rsNum(rsCelda,"Status_EDT")<2 then
          retval = "-"
        else
          valor = rsNum(rsCelda, "ResultadoValidacion_" & id )
        end if
      else
        valor = rsNum(rsCelda, "Resultado_" & id )
      end if
      if retval<>"-" then
        idpersona = rsNum(rsCelda,"IdPersonal")
        valor = valor / 100
        if printtype = 3 then
          retval = valor
        elseif id = MT_COMPETENCIAS then
          retval = khorFormatCompatibilidadDisp(valor,(khorDiferenciaCompetencias(idpersona,idperfil,khorEscenarioComportamiento(idperfil))>0))
        else
          retval = khorFormatPorcentaje(valor)
        end if
        conLiga = true
        modulosActivos = khorModulosActivos()
        select case id
          case MT_PRUEBAS
            conLiga = khorPermisoModulo(Modulo_Psicometria,modulosActivos)
          case MT_COMPETENCIAS
            conLiga = khorPermisoModulo(Modulo_CompatibilidadCompetencias,modulosActivos)
        end select
        'if printtype = 0 and conLiga AND NOT khorPage_isPDF() then
         ' if itemConPerfil then
           ' retval = "<a href=""#"" onClick=""return detalle" & id & "('" & idpersona & "','" & idperfil & "');"">" & _
          '            retval & "</a>"
          'else
            'retval = "<a href=""#"" onClick=""return detalle" & id & "('" & idpersona & "','" & rsNum(rsCelda,"IdGrupo_" & id) & "');"">" & _
            '          retval & "</a>"
          'end if
        'end if
      end if
    end if
    getCeldaRS = retval
  end function

  function getValorRS(rs)
     if id=MT_EDT then
      getValorRS = rsNum(rs,"ResultadoValidacion_" & id ) * peso / 100
    else
      getValorRS = rsNum(rs,"Resultado_" & id ) * peso / 100
    end if
  end function

  function qrySelect()
    dim retval : retval = "rep" & id & ".Resultado_" & id
    if not itemConPerfil then
      'retval = retval& ", rep" & id & ".IdGrupo AS IdGrupo_" & id
    end if
    if id=MT_EDL then
      retval = retval & ", repEDL.Status_EDL"
    elseif id=MT_EDD then
      retval = retval & ", repEDD.numRevisiones_EDD"
    elseif id=MT_EDP then
      retval = retval & ", repEDP.numRevisiones_EDP"
    elseif id=MT_EDT then
      retval = retval & ", repEDT.resultadoValidacion_EDT, repEDT.Status_EDT"
    end if
    qrySelect = retval
  end function

  sub declareJSdetalle() %>
    function detalle<%=id%>(idper,idaux) {
      var loc = '';  <%
    SELECT CASE id
      CASE MT_PRUEBAS %>
      abrePsicometria(idper,idaux); <%
      CASE MT_COMPETENCIAS %>
      abreCompetencias(idper,idaux); <%
      CASE MT_MCIE, MT_MCIP, MT_MCIT, MT_MCID %>
      loc = 'mciRepIndividual.asp?childwin=FRMmciRepIndividual'; <%
      CASE MT_360 %>
      loc = '<%=urlReporteDefault360()%>';
      loc += ( (loc.indexOf('?') == -1) ? '?' : '&') +'childwin=<%=idReporteDefault360()%>&IdPersona='+idper; <%
      CASE MT_EDL %>
      loc = 'EDL_Rep_Individual.asp?childwin=FRMedlRepIndividual'; <%
      CASE MT_EDD %>
      loc = 'EDD_Rep_Individual.asp?childwin=FRMeddRepIndividual'; <%
      CASE MT_EDP %>
      loc = 'EDP_Rep_Individual.asp?childwin=FRMedpRepIndividual'; <%
      CASE MT_EDK %>
      loc = 'EDK_Rep_Individual.asp?childwin=FRMedkRepIndividual'; <%
      CASE MT_EDS %>
      loc = 'EDS_Evaluacion.asp?childwin=FRMedsEvaluacion'; <%
      CASE MT_EDE %>
      alert('<%=strJS(lblFRS_NoHayDetalleDisponible)%>'); <%
      CASE MT_ADP %>
      loc = 'ED_ADP.asp?childwin=FRMedADP'; <%
      CASE MT_ADJ %>
      loc = 'ED_ADJ.asp?childwin=FRMedADJ'; <%
      CASE MT_EDT %>
      loc = 'EDT_Rep_Individual.asp?childwin=FRMedtRepIndividual'; <%
    END SELECT
    IF id<>MT_PRUEBAS AND id<>MT_COMPETENCIAS AND id<>MT_EDE THEN %>
      if (loc != '') {
        var whdl=window.open("","mapaDetalle","toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=700,height=520");
        var cact=document.TrueForm.action;
        if (whdl) sendval('', '_target','mapaDetalle', '_action',loc, 'IdPersonal',idper, 'IdGrupo',idaux );
        sendval('', 'nosubmit', '_target','', '_action',cact);
      } <%
    END IF %>
      return false;
    } <%
  end sub

  function qryFactor()
    if id=MT_EDT then
      qryFactor = "(CASE WHEN rep" & id & ".ResultadoValidacion_" & id & " IS NULL THEN 0 ELSE rep" & id & ".ResultadoValidacion_" & id & " * " & peso & " END)"
    else
      qryFactor = "(CASE WHEN rep" & id & ".Resultado_" & id & " IS NULL THEN 0 ELSE rep" & id & ".Resultado_" & id & " * " & peso & " END)"
    end if
  end function

  function qryFrom(idperfil,idperiodo)
    dim retval
    retval = "LEFT JOIN rep" & id & " ON Personal.IdPersonal = rep" & id & ".IdPersonal"
    if not itemConPerfil then
      retval = retval & " AND rep" & id & ".IdPeriodo = " & idperiodo
    elseif idperfil=0 then
      retval = retval & " AND Personal.Puesto = rep" & id & ".IdPuesto"
    end if
    qryFrom = retval
  end function

  function qryWhere(idperfil)
    if itemConPerfil AND idperfil<>0 then
      qryWhere = "(rep" & id & ".IdPuesto=" & idperfil & ")"
    else
      qryWhere = ""
    end if
  end function

  function qryWherePartial()
    qryWherePartial = "(rep" & id & ".IdPersonal IS NOT NULL)"
  end function

  sub clean()
    id=""
    tipo=0
    desc=""
    peso=0
    total=0
    total2=0
    auxtotal=0
    auxtotal2=0
  end sub
  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'================================================================================

class mapaTalentoClass
  public mtArr
  public mtNum

  '---------- Funciones particulares de cada rubro

  function pesoEnPeriodo(idperiodo,rubro)
    pesoEnPeriodo = getBDnum("Peso","SELECT Peso FROM ED_PeriodoRubro WHERE IdPeriodo=" & idperiodo & " AND IdRubro=" & rubro)
  end function
  
  sub Initalize(modulosActivos,fromrequest,idperiodo,override)

    mtAdd MT_PERFIL, MT_COMPETENCIAS, lblKHOR_Competencias, 0, false
    mtAdd MT_PERFIL, MT_PRUEBAS, lblKHOR_Psicometria, 100, false
    
    'ajustaPesos MT_PERFIL
    
   
    noDefault = true
    if reqn("peso_EDATC") = 0 AND reqn("peso_EDASI") = 0 AND reqn("peso_EDASIO") = 0 AND reqn("peso_EDBACH") = 0 AND reqn("peso_EDTSU") = 0 AND reqn("peso_EDFAC") = 0 THEN
      noDefault = false
    END IF
    
    mtAdd MT_EJECUCION, MT_EDATC, "EDATC", iif(fromrequest and noDefault,reqn("peso_EDATC"),17), fromrequest
    mtAdd MT_EJECUCION, MT_EDASI, "EDASI", iif(fromrequest and noDefault,reqn("peso_EDASI"),17), fromrequest
    mtAdd MT_EJECUCION, MT_EDASIO, "EDASIO", iif(fromrequest and noDefault,reqn("peso_EDASIO"),17), fromrequest
    mtAdd MT_EJECUCION, MT_EDBACH, "EDBACH", iif(fromrequest and noDefault,reqn("peso_EDBACH"),17), fromrequest
    mtAdd MT_EJECUCION, MT_EDTSU, "EDTSU", iif(fromrequest and noDefault,reqn("peso_EDTSU"),16),fromrequest
    mtAdd MT_EJECUCION, MT_EDFAC, "EDFAC", iif(fromrequest and noDefault,reqn("peso_EDFAC"),16), fromrequest
    
    'ajustaPesos MT_EJECUCION
    mapaTalento_Initalize Me, modulosActivos, fromrequest, idperiodo, override
  end sub

  sub declareJSdetalle()
    dim i
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        mtArr(i).declareJSdetalle
      end if
    next
  end sub

  '---------- Funciones auxiliares

  function qryString()
    dim retval, i
    retval = ""
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        retval = retval & "&peso_" & mtArr(i).id & "=" & mtArr(i).peso
      end if
    next
    qryString = retval
  end function

  function condicionPersona(idperiodo)
    dim i, auxs
    dim retval : retval = ""
    dim hayed : hayed = false
    dim haymci : haymci = false
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        if mtArr(i).tipoCR = tipoCR_360 then
          auxs = "EXISTS(SELECT * FROM GrupoEntidad360 ge360 INNER JOIN Grupo360 g360 ON ge360.IdGrupo = g360.IdGrupo WHERE ge360.TipoEntidad <> 4 AND ge360.IdEntidad=Personal.IdPersonal AND g360.IdPeriodo=" & IdPeriodo & ")"
        elseif compatibilidadEsRubroMCI(mtArr(i).tipoCR) then
          auxs = "" : haymci = true
        elseif compatibilidadEsRubroED(mtArr(i).tipoCR) AND (mtArr(i).tipoCR <> tipoCR_EDE) then  '-- EDE no esta en grupo ED
          auxs = "" : hayed = true
        elseif not mtArr(i).itemConPerfil then
          auxs = "EXISTS(SELECT * FROM rep" & mtArr(i).id & " WHERE rep" & mtArr(i).id & ".IdPersonal=Personal.IdPersonal AND IdPeriodo=" & IdPeriodo & ")"
        end if
        retval = strAdd( retval, " OR ", auxs )
      end if
    next
    if hayed then
      auxs = "EXISTS(SELECT * FROM ED_GrupoPersona edgp INNER JOIN ED_Grupo edg ON edgp.IdGrupo = edg.IdGrupo WHERE edgp.IdEvaluado=Personal.IdPersonal AND edg.IdPeriodo=" & IdPeriodo & ")"
      retval = strAdd( retval, " OR ", auxs )
    end if
    if haymci then
      auxs = "EXISTS(SELECT * FROM repMCI mci WHERE mci.IdPersonal=Personal.IdPersonal AND mci.IdPeriodo=" & IdPeriodo & ")"
      retval = strAdd( retval, " OR ", auxs )
    end if
    if retval<>"" then retval = "(" & retval & ")"
    condicionPersona = retval
  end function

  '---------- Funciones para obtencion de datos

  function getValorRS(rs,tipo)
    dim valor, i
    valor = 0
    for i=1 to mtNum
      if mtArr(i).tipo=tipo AND mtArr(i).peso > 0 then
        valor = valor + mtArr(i).getValorRS(rs)
      end if
    next
    getValorRS = valor
  end function

  function queryWhereJoin(tipo)
    dim retval, i
    retval = ""
    for i=1 to mtNum
      if mtArr(i).tipo=tipo AND mtArr(i).peso > 0 then
        retval = strAdd(retval, " OR ", mtArr(i).qryWherePartial())
      end if
    next
    if retval<>"" then retval = "(" & retval & ")"
    queryWhereJoin = retval
  end function

  function query(idperfil,idperiodo,condicion,groupBy)
    dim retval, auxc, i
    retval = "SELECT Personal.IdPersonal, Personal.Nombre, Personal.IdSucursal, Personal.Puesto AS IdPuesto, Puestos.Puesto"
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        retval = strAdd( retval, ",", mtArr(i).qrySelect() )
      end if
    next
    retval = retval & " FROM Personal LEFT JOIN Puestos ON Personal.Puesto = Puestos.IdPuesto"
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        retval = strAdd( retval, " ", mtArr(i).qryFrom(idperfil,idperiodo) )
      end if
    next
    auxc = condicion
    for i=1 to mtNum
      if mtArr(i).peso > 0 then
        auxc = strAdd(auxc, " AND ", mtArr(i).qryWhere(IdPerfil))
      end if
    next
    retval = retval & " WHERE " & auxc
    
    
    retval = retval & " GROUP BY Personal.IdPersonal, Personal.Nombre, Personal.IdSucursal, Personal.Puesto, Puestos.Puesto, "&_
                      "repMTCP.Resultado_MTCP"
    for i=1 to mtNum
      if mtArr(i).peso > 0 and i <> 2  then
          retval = retval & ",rep"&mtArr(i).desc&".Resultado_"&mtArr(i).desc
      end if
    next
    
    auxc = "" 
    
    for i=1 to mtNum
      if mtArr(i).peso > 0  then
        auxc = strAdd(auxc, " + ", mtArr(i).qryFactor())
      end if
    next

    retval = retval & " ORDER BY " & iif(groupBy = 1, "IdPuesto, ", "") & "(" & auxc & ") DESC, Nombre"
    query = retval
  end function

  '---------- Funciones generales

  function getIdx(id)
    dim retval,i
    retval=-1
    for i=1 to mtNum
      if mtArr(i).id = id then
        retval = i
        exit for
      end if
    next
    getIdx = retval
  end function

  function peso(id)
    dim idx
    peso = 0
    idx = getIdx(id)
    if idx<>-1 then
      peso = mtArr(idx).peso
    end if
  end function

  function numItems(tipo,conpeso)
    dim retval,i
    retval=0
    for i=1 to mtNum
      if mtArr(i).tipo = tipo AND (mtArr(i).peso>0 OR NOT conPeso) then
        retval = retval + mtArr(i).itemCols
      end if
    next
    numItems = retval
  end function

  function sumaPesos(tipo)
    dim retval,i
    retval=0
    for i=1 to mtNum
      if mtArr(i).tipo = tipo then
        retval = retval + mtArr(i).peso
      end if
    next
    sumaPesos = retval
  end function

  function ajustaPesos(tipo)
    dim auxt, slice, i
    auxt = sumaPesos(tipo)
    if auxt<>100 and numItems(tipo,auxt<>0)>0 then
      slice = int(((100 - auxt)/numItems(tipo,auxt<>0)))
      for i=1 to mtNum
        if mtArr(i).tipo = tipo AND (mtArr(i).peso>0 OR auxt=0) then
          mtArr(i).peso = mtArr(i).peso + slice
        end if
      next
      slice = 100 - sumaPesos(tipo)
      while slice<>0
        for i=1 to mtNum
          if mtArr(i).peso>0 AND mtArr(i).tipo = tipo and slice<>0 then
            auxt = iif(slice>0,1,-1)
            mtArr(i).peso = mtArr(i).peso + auxt
            slice = slice - auxt
          end if
        next
      wend
    end if
  end function

  sub mtAdd(tipo,id,desc,pesodefault,fromrequest)
    mtNum=mtNum+1
    redim preserve mtArr(mtNum)
    set mtArr(mtNum) = new mapaTalentoItem
    mtArr(mtNum).id = id
    mtArr(mtNum).tipo = tipo
    mtArr(mtNum).desc = desc
    mtArr(mtNum).peso = iif(fromrequest and noDefault,reqn("peso_"&desc),pesodefault)
    mtArr(mtNum).total = 0
    mtArr(mtNum).total2 = 0
  end sub

  sub clean()
    dim i
    for i=1 to mtNum
      set mtArr = nothing
    next
    mtNum = 0
    redim mtArr(mtNum)
  end sub
  private sub class_initialize()
    mtNum=0
    clean
  end sub
  private sub class_terminate()
    clean
  end sub
end class

'================================================================================

%>