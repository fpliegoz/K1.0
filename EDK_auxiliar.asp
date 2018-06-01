<!--#include file="./edk_especial.asp"-->
<!--#include file="./EDK_PesosEspecial.asp"-->
<%
'========================================

  function edk_LoadEventos(colEvento,IdEvaluacion,IdEvaluacionK)
    dim resEvento : resEvento = 0
    dim i, auxo 
    colEvento.clean
    colEvento.keyDescFromTable "EDK_EventoPersona", "IdEventoK", "Fecha", "Puntos", "IdEventoK", "IdEvaluacion=" & IdEvaluacion
    for i=1 to colEvento.count
      set auxo = colEvento.obj(i)
      resEvento = resEvento + auxo.aux
    next
    set auxo = nothing
    edk_LoadEventos = resEvento
  end function

  sub edk_tablaEventos(colEvento,IdEvaluacion,IdEvaluacionK,adminEventos)
    dim sq, rs, i, tot, auxo, desc, estilo
      if (colEvento.count > 0) OR adminEventos then %>
      <table border="0" cellspacing="0" cellpadding="1" align="center" class="tsmall" width="400">
        <tr class="celdaTit">
          <td colspan="<%=iif(adminEventos,3,2)%>"><%=lblEDK_Evento%> / *<%=lblEDK_EventoTotal%></td>
        </tr> <%
        tot = 0
        for i=1 to colEvento.count
          set auxo = colEvento.obj(i)
          desc = descFromTable("EDK_EventoCat","IdEventoK","EventoK",auxo.key)
          estilo = switchEstilo(estilo) %>
        <tr class="<%=estilo%>">
          <td class="tsmall"><%=desc%></td>
          <td align="right"><%=iif(auxo.aux>0,"+","") & auxo.aux%></td> <%
          tot = tot + auxo.aux
          if adminEventos then %>
          <td class="tsmall">[<a href="#" onClick="return borraEvento(<%=auxo.key%>)"><%=lblFRS_Borrar%></a>]</td> <%
          end if %>
        </tr> <%
        next %>
      </table> <%
        end if
        if adminEventos then %>
        <div class="tsmall" align="center" nowrap>
            <select id="ideventok" name="ideventok" style="font-size:10px;">
              <option value="0">-</option> <%
              sq = "SELECT a.IdEventoK, a.EventoK, b.Puntos" & _
                " FROM EDK_EventoCat a INNER JOIN EDK_EventoEvaluacion b ON a.IdEventoK = b.IdEventoK AND b.IdEvaluacionK=" & IdEvaluacionK & _
                " WHERE NOT EXISTS(SELECT * FROM EDK_EventoPersona c WHERE c.IdEventoK = a.IdEventoK AND IdEvaluacion=" & IdEvaluacion & ")" & _
                " ORDER BY EventoK"
              set rs = getrs(conn,sq)
              while not rs.eof
                i = rsNum(rs,"Puntos") %>
              <option value="<%=rsNum(rs,"IdEventoK")%>"><%=rsStr(rs,"EventoK") & " (" & iif(i>0,"+","") & i & ")"%></option> <%
                rs.movenext
              wend
              rs.close %>
            </select>
          [<a href="#" onClick="return agregaEvento()"><%=lblFRS_Agregar%></a>]
        </div> <%
        end if %>
      <input type="hidden" id="totEvento" name="totEvento" value="<%=tot%>"> <%
    set auxo = nothing
    set rs = nothing
  end sub

'========================================
  dim edkGlobalPeriodoOE : edkGlobalPeriodoOE = 0
  
  function edkDescObjetivosEstrategico(regOE,tsel)
    '-- Variable sucia, declarar en khorLabelEspecial.asp o edk_especial.asp
    'edk_showOEdisplay = true '-- Muestra el objetivo estrategico como "abreviatura: descripcion"
    if edk_showOEdisplay then
      edkDescObjetivosEstrategico = regOE.ObjetivoEstrategicoDisplay()
    else  '-- default
      edkDescObjetivosEstrategico = regOE.Abreviatura
    end if
  end function

  sub edkGetObjetivosEstrategicos(colOEbase,IdUnidadAdministrativa)
    dim objC, objO
    dim oeManager : set oeManager = new oe_treeNodeManager
    '-- Puebla el arbol de Clasificacion-ObjetivoEstrategico
    dim sq : sq = "SELECT * FROM vED_ObjetivoEstrategico" & _
                  " WHERE IdUnidadAdministrativa=" & IdUnidadAdministrativa & " AND IdOEactual=0 AND StatusOE<>0"
    if khorEDoePorPeriodo() then sq = sq & " AND IdPeriodoOE = " & edkGlobalPeriodoOE
    sq = sq & " ORDER BY OrdenClasificacion, Clasificacion, ObjetivoEstrategico"
    dim rs : set rs = getrs(conn,sq)
    while not rs.eof
      set objC = colOEbase.objByKey( rsNum(rs,"IdClasificacion") )
      if objC is nothing then
        set objC = colOEbase.addTreeNodeFromRS( rs, oeManager, 1 )
      end if
      set objO = objC.col.objByKey( rsNum(rs,"IdObjetivoEstrategico") )
      if objO is nothing then
        set objO = objC.col.addTreeNodeFromRS( rs, oeManager, 2 )
      end if
      rs.movenext
    wend
    rs.close
    set rs = nothing
    set oeManager = nothing
  end sub

  sub edkPaintObjetivosEstrategicos(colOEbase)
    dim objC, objO, id, ic, io %>
        <table border="0" cellspacing="1" cellpadding="1" align="center" id="edRK">
          <tr class="celdaTit">
            <td align="center"><%=lblED_Area%></td>
            <td align="center" colspan="2"><%=lblKHOR_ObjetivoEstrategico%></td>
            <td align="center"><%=lblFRS_Descripcion%></td>
          </tr> <%
            for ic=1 to colOEbase.count
              set objC = colOEbase.obj(ic) %>
          <tr>
            <td rowspan="<%=objC.rowspan()%>"><%=objC.data.desc%></td> <%
              if objC.col.count=0 then %>
            <td colspan="3">&nbsp;</td>
          </tr> <%
              else
                for io=1 to objC.col.count
                  set objO = objC.col.obj(io)
                  id = objO.data.IdObjetivoEstrategico
                  if io>1 then %>
          <tr> <%
                  end if %>
            <td><input type="radio" id="OEradio" name="OEradio" value="<%=id%>">
              <input type="hidden" id="OEradioDisplay<%=id%>" name="OEradioDisplay<%=id%>" value="<%=serverHTMLencode(edkDescObjetivosEstrategico(objO.data,true))%>">
            </td>
            <td id="oeDisplay<%=id%>"><%=objO.data.ObjetivoEstrategicoDisplay%></td>
            <td id="oeDescription<%=id%>"><%=objO.data.Descripcion%></td>
          </tr> <%
                next
              end if
            next %>
        </table> <%
  end sub

'========================================

class EDK_Interfaz
  public persona     ' AS clsPersona
  public periodo     ' AS ED_Periodo
  public colRevision ' AS frsCollection OF ED_PeriodoRevision
  public calRevision  'Revision anterior o actual mas reciente por calendario
  public idxRevision  
  public regGP       ' AS ED_GrupoPersona
  public regEvaK     ' AS EDK_Evaluacion
  public colRubroK   ' AS frsCollection OF EDK_Rubro
  public colComment   ' AS frsCollection OF frsKeyDesc
  public colEvento    ' AS frsCollection OF frsKeyDesc
  private hayEvento
  private resEvento
  public escalaK     ' AS EDK_Escala
  private columnaPeso
  private numColumnas
  public numDec
  public hideSelfResults
  public habAutoRevision
  public logroMaxResultado

  private sub class_initialize()
    set persona = new clsPersona
    set periodo = new ED_Periodo
    set colRevision = new frsCollection 'OF ED_PeriodoRevision
    calRevision = 0
    idxRevision = 0
    set regGP = new ED_GrupoPersona
    set regEvaK = new EDK_Evaluacion
    set escalaK = new EDK_Escala
    set colRubroK = new frsCollection 'OF EDK_Rubro
    set colComment = new frsCollection 'OF frsKeyDesc
    set colEvento = new frsCollection 'OF frsKeyDesc
    hayEvento = false
    resEvento = 0
    numDec = 1
    hideSelfResults = false
    habAutoRevision = edkAutoRevision()
    logroMaxResultado = 100
  end sub
  
  private sub class_terminate()
    colEvento.clean
    set colEvento = nothing
    colComment.clean
    set colComment = nothing
    colRubroK.clean
    set colRubroK = nothing
    set regGP = nothing
    set escalaK = nothing
    set regEvaK = nothing
    colRevision.clean
    set colRevision = nothing
    set periodo = nothing
    set persona = nothing
  end sub
  
  sub getInfoPeriodo(idper)
    periodo.getFromDB conn, idper
    dim sq : sq = "SELECT * FROM ED_PeriodoRevision WHERE IdPeriodo=" & periodo.IdPeriodo & " ORDER BY numRevision"
    dim rs : set rs = getrs(conn,sq)
    dim auxrev
    colRevision.clean
    calRevision = 0
    while not rs.eof
      set auxrev = new ed_PeriodoRevision
      auxrev.getFromRS rs
      colRevision.add auxrev, auxrev.numRevision
      if Date >= auxrev.FechaIni then calRevision = auxrev.numRevision
      rs.movenext
    wend
    rs.close
    set rs = nothing
    set auxrev = nothing
  end sub
  
  sub getInfoEvaluacion(idper)
    persona.getFromDB conn, idper
    regGP.getFromDBfilterView conn, "IdEvaluado = " & persona.IdPersona & " AND IdPeriodo =" & periodo.IdPeriodo
    if regGP.IdEvaluacion > 0 AND regGP.IdEvaluacionK > 0 then
      regEvaK.getFromDBcomplete regGP.IdEvaluacionK
      if regEvaK.conObjetivoEstrategico and regGP.IdUnidadAdministrativa_EDP=0 then
        regGP.IdUnidadAdministrativa_EDP = reqn("IdUA")
        if regGP.IdUnidadAdministrativa_EDP=0 then
          regGP.IdUnidadAdministrativa_EDP = regGP.IdUnidadAdministrativaEvaluado
        end if
      end if
      hideSelfResults = (persona.IdPersona = personaSesion()) AND edHideSelfResults()
      escalaK.getFromDB conn, regEvaK.IdEscalaK
      numDec = escalaK.numDecimales
      if edk_pesodecimal then regEvaK.PesoDecimal = true
      '-- Eventos
      hayEvento = bdExists("SELECT * FROM EDK_EventoCat a INNER JOIN EDK_EventoEvaluacion b ON a.IdEventoK = b.IdEventoK AND b.IdEvaluacionK=" & regGP.IdEvaluacionK )
      resEvento = edk_LoadEventos(colEvento,regGP.IdEvaluacion,regGP.IdEvaluacionK)
      '-- Comentarios
      colComment.clean
      if not edkCompromisos() then
        colComment.keyDescFromMenuFijo(menuFijo_EDKcomentarios)
        dim c, auxo
        for c=1 to colComment.count
          set auxo = colComment.obj(c)
          auxo.desc = getBD("Comentario","SELECT Comentario FROM EDK_Comentario WHERE IdEvaluacion=" & regGP.IdEvaluacion & " AND TipoComentario=" & auxo.key)
          '-- Variable sucia: declarar en khorlabelespecial.asp o similar
          'edk_comentariosopcionales = true  '-- comentarios opcionales
          if edk_comentariosopcionales then auxo.aux = "optional"
        next
      end if
    end if
  end sub
  
  function enFechasRevision(numrev)
    enFechasRevision = enFechasRevisionExtended(numrev,-1)
  end function
  function enFechasRevisionExtended(numrev,idper)
    dim allowPermiso : allowPermiso = true
    dim ok : ok = false
    dim auxrev : set auxrev = colRevision.objByKey(numrev)
    if numrev = 0 then
      ok = (periodo.edpObjFechaIni <= Date) AND (Date <= periodo.edpObjFechaFin)
    elseif not auxrev is nothing then
      dim fini : fini = auxrev.FechaIni
      dim ffin : ffin = auxrev.FechaFin
      if habAutoRevision AND (idper > 0) then
        if idper = persona.idpersona then '-- Evaluado
          if isnull(auxrev.FechaMed) then
            fini = dateAdd("d",1,Date)  '-- fecha futura
            allowPermiso = false
          else
            ffin = auxrev.FechaMed
          end if
        elseif not isnull(auxrev.FechaMed) and not edkAutoRevisionTraslapada then  '-- Evaluador
          fini = dateAdd("d",1,auxrev.FechaMed)
        end if
      end if
      ok = (fini <= Date) AND (Date <= ffin) AND _
            bdExists("SELECT * FROM EDK_Rubro WHERE IdRubroActual=0 AND IdEvaluacion=" & regGP.IdEvaluacion)
      set auxrev = nothing
    end if
    if (numrev=0 or (numrev=calRevision and not isnull(regGP.Fecha_EDK))) and allowPermiso AND not ok then 
      dim sq : sq = "SELECT * FROM ED_GrupoPersonaPermiso WHERE IdEvaluacion=" & regGP.IdEvaluacion & _
                    " AND TipoPermiso=" & tipoCR_EDK &_
                    " AND " & formatDateSQL(Date,false) & " BETWEEN permisoFechaIni AND permisoFechaFin"
      ok = bdExists(sq)
    end if
    enFechasRevisionExtended = ok
  end function
  
  '----------------------------------------
  
  function OE2RubroK(auxOE,objR)
    dim changed : changed = false
    if auxOE.TipoOE = 0 then  '-- estructurado
      dim auxs : auxs = auxOE.VerboOE & " " & auxOE.CuantoOE & auxOE.UnidadOE & " " & auxOE.QueOE & " " & auxOE.ParaQueOE & iif(auxOE.PolaridadOE=0," (+)"," (-)")
      if objR.Rubro <> auxs then
        objR.Rubro = auxs
        changed = true
      end if
      if objR.Esperado <> auxOE.QueOE then
        objR.Esperado = auxOE.QueOE
        changed = true
      end if
      if objR.LogroEsperado <> auxOE.CuantoOE then
        objR.LogroEsperado = auxOE.CuantoOE
        changed = true
      end if
    else  '-- enunciativo
      if objR.Rubro <> auxOE.ObjetivoEstrategico then 
        objR.Rubro = auxOE.ObjetivoEstrategico
        changed = true
      end if
      if objR.Esperado <> auxOE.Descripcion then
        objR.Esperado = auxOE.Descripcion
        changed = true
      end if
    end if
    dim changecal : changecal = false
    if objR.PesoRubro <> auxOE.PesoOE then
      objR.PesoRubro = auxOE.PesoOE
      changecal = true
    end if
    if objR.rev.Calificacion <> auxOE.CalificacionOE then
      objR.rev.Calificacion = auxOE.CalificacionOE
      objR.rev.Fecha = auxOE.FechaOE
      changecal = true
    end if
    if changecal then
      objR.rev.Resultado = objR.rev.Calificacion * objR.PesoRubro / 100.0
      changed = true
    end if
    if objR.rev.numRevision > 0 and changecal then
      objR.rev.dirty = true
    end if
    OE2RubroK = changed
  end function
  
  sub verificaOErubroK(objS,objR,IdUA)
    dim rchanged : rchanged = false
    if objS.TipoSeccionK = EDK_TIPOORG OR objS.TipoSeccionK = EDK_TIPOEST then
      dim auxOE : set auxOE = new ED_ObjetivoEstrategico
      do
        auxOE.getFromDB conn, objr.IdObjetivoEstrategico
        if auxOE.IdOEactual > 0 then  '-- Hay una nueva version del OE
          objr.IdObjetivoEstrategico = auxOE.IdOEactual
          rchanged = true
        end if
      loop while auxOE.IdOEactual > 0
      if (objS.TipoSeccionK = EDK_TIPOEST) AND (auxOE.IdUnidadAdministrativa <> IdUA) AND (IdUA > 0) then
        auxOE.IdOEactual = -1 '-- Simula que el objetivo fue eliminado, para excluir el rubro
      end if
      if auxOE.IdOEactual < 0 then  '-- El OE fue eliminado
        conn.execute "DELETE FROM EDK_RubroRevision WHERE IdRubro = " & objR.IdRubro
        conn.execute "DELETE FROM EDK_Rubro WHERE IdRubro = " & objR.IdRubro
        objr.IdObjetivoEstrategico = 0
        objr.IdRubro = 0
      else
        if OE2RubroK(auxOE,objR) OR rchanged then
          objR.updateRev conn, objR.rev.numRevision, false
        end if
      end if
      set auxOE = nothing
    end if
  end sub

  sub getRubrosK()
    if regGP.IdEvaluacion <> 0 then
      dim objR, r, objS, s, c, objc, sq
      '-- Carga competencias del catalogo en las secciones correspondientes
      if regEvaK.usePuesto AND (regGP.IdPuesto_EDK = 0) AND (regGP.IdPuestoEvaluado > 0) then
        regGP.IdPuesto_EDK = regGP.IdPuestoEvaluado
        conn.execute "UPDATE ED_GrupoPersona SET IdPuesto_EDK = " & regGP.IdPuesto_EDK & " WHERE Idevaluacion = " & regGP.IdEvaluacion
      end if
      for s=1 to regEvaK.colSeccion.count
        set objs = regEvaK.colSeccion.obj(s)
        if objs.TipoSeccionK = EDK_TIPOCOT then
          objs.getCompetenciasTecnicasPuesto regGP.IdPuesto_EDK, numDec
          if (objs.TipoSeccionK = EDK_TIPOCOT) AND NOT objs.preloadCompetencias then regEvaK.allFixed = false
        end if
        if edkModoPesos() > 0 then
          sq = "SELECT PesoSeccionK FROM EDK_SeccionPeriodoPeso WHERE IdSeccionK = " & objS.IdSeccionK & _
              " AND IdPeriodo = " & regGP.IdPeriodo & " AND IdEntidadAux = " & edk_pesos_entidad_valor( regGP )
          objs.pesoPonderado = getBDnum( "PesoSeccionK", sq )
        end if
      next
      '-- Rubros
      dim numrev : numrev = idxRevision
      sq = "SELECT rub.*, rev.numRevision, rev.Calificacion, rev.Resultado, rev.Fecha, rev.Observaciones, rev.Logro, rev.ResultadoLogro, rev.ObsEvaluado" & _
          " FROM EDK_Rubro rub INNER JOIN EDK_Seccion secc ON rub.IdSeccionK = secc.IdSeccionK AND secc.IdEvaluacionK = " & regGP.IdEvaluacionK & _
          " LEFT JOIN EDK_SeccionCompetencia secom ON rub.IdSeccionCompetencia = secom.IdSeccionCompetencia" & _
          " LEFT JOIN EDK_RubroRevision rev ON rub.IdRubro = rev.IdRubro AND rev.numRevision = " & numrev & _
          " WHERE IdEvaluacion=" & regGP.IdEvaluacion & " AND IdRubroActual=0" & _
          " ORDER BY secc.SecuenciaS, secom.SecuenciaC, rub.PesoRubro DESC, rub.IdRubro"
      dim rs : set rs = getrs(conn,sq)
      colRubroK.clean
      while not rs.eof
        rubok = true
        set objR = new EDK_Rubro
        objR.getFromRSrev rs
        set objS = regEvaK.colSeccion.objByKey(objR.IdSeccionK)
        if objs.TipoSeccionK = EDK_TIPOCOT then
          set objc = objs.colCompetencia.objByKey(objR.IdCompetencia)
          if objc is nothing then '-- Ya no esta en el perfil! usa valores guardados del rubro
            set objc = new EDK_SeccionCompetencia
            objc.IdSeccionCompetencia = objR.IdCompetencia
            objc.Nivel = objR.NivelCompetencia
            objc.IdSeccionK = objR.IdSeccionK
            objc.SeccionCompetencia = getBD("Competencia","SELECT Competencia FROM CatCompetencias360 WHERE IdCompetencia=" & objR.IdCompetencia)
            objc.DescripcionCompetencia = getBD("Descripcion","SELECT Descripcion FROM ECCompetenciaNivel WHERE IdCompetencia=" & objR.IdCompetencia & " AND Nivel=" & objc.Nivel)
            objc.DescripcionAdicional = getBD("Definicion","SELECT Definicion FROM CatCompetencias360 WHERE IdCompetencia=" & objR.IdCompetencia)
            objs.colCompetencia.add objc, objR.IdCompetencia
          elseif objc.Nivel <> objR.NivelCompetencia then '-- El nivel cambió en el perfil! usa valor guardado del rubro
            objc.Nivel = objR.NivelCompetencia
            objc.DescripcionCompetencia = getBD("Descripcion","SELECT Descripcion FROM ECCompetenciaNivel WHERE IdCompetencia=" & objR.IdCompetencia & " AND Nivel=" & objc.Nivel)
          end if
        elseif objs.TipoSeccionK = EDK_TIPOORG OR objs.TipoSeccionK = EDK_TIPOEST then
          verificaOErubroK objs, objR, regGP.IdUnidadAdministrativa_EDP
        end if
        '--
        if isnull(rs("Fecha")) then
          if objS.lblRubro="" then
            objR.rev.Resultado = 100
            if escalaK.MaxResultado <> 0 then
              objr.rev.Fecha = Now
              objr.rev.Calificacion = 100 * div( 100, objr.PesoRubro )
            else
              objr.rev.Calificacion = edk_Resultado2Escala(escalaK.IdEscalaK,objr.rev.Resultado,"Calificacion")
            end if
          end if
        end if
        '--
        if objR.IdRubro > 0 then
          colRubroK.add objR, objR.IdRubro
        end if
        rs.movenext
      wend
      rs.close
      set rs = nothing
      '-- Si no hay rubros, carga las competencias predefinidas en las seciones correspondientes
      if (numrev = 0 OR regEvaK.allFixed) then
        dim auxid : auxid = 0
        for s=1 to regEvaK.colSeccion.count
          set objs = regEvaK.colSeccion.obj(s)
          if ((objs.TipoSeccionK = EDK_TIPOCOM) OR ((objs.TipoSeccionK = EDK_TIPOCOT) AND objs.preloadCompetencias)) AND (objs.colCompetencia.count > 0) then
            nr = 0
            for r=1 to colRubroK.count
              set objR = colRubroK.obj(r)
              if objR.IdSeccionK = objs.IdSeccionK then
                nr = nr + 1
              end if
            next
            if nr = 0 then
              for c=1 to objs.colCompetencia.count
                set objc = objs.colCompetencia.obj(c)
                set objR = new EDK_Rubro
                objR.IdEvaluacion = regGP.IdEvaluacion
                objR.getFromCom objc
                if objs.TipoSeccionK = EDK_TIPOCOT then
                  objR.IdCompetencia = objR.IdSeccionCompetencia
                  objR.IdSeccionCompetencia = 0
                  objR.NivelCompetencia = objc.Nivel
                end if
                if regEvaK.allFixed then
                  objR.update conn
                  auxid = objR.IdRubro
                else
                  auxid = auxid - 1 '-- negativos, solo para key de la coleccion
                end if
                colRubroK.add objR, auxid
              next
            end if
          end if
        next
        set objs = nothing
      end if
      '-- Verifica objetivos estrategicos faltantes
      dim lstOE 
      for s=1 to regEvaK.colSeccion.count
        set objs = regEvaK.colSeccion.obj(s)
        if (objs.TipoSeccionK = EDK_TIPOEST AND regGP.IdUnidadAdministrativa_EDP > 0) OR objs.TipoSeccionK = EDK_TIPOORG then
            lstOE = ""
            for r=1 to colRubroK.count
              set objR = colRubroK.obj(r)
              if objR.IdSeccionK = objs.IdSeccionK then
                lstOE = strAdd( lstOE, ",", objR.IdObjetivoEstrategico )
              end if
            next
            dim auxOE : set auxOE = new ED_ObjetivoEstrategico
            sq = "SELECT * FROM ED_ObjetivoEstrategico WHERE IdOEactual = 0 AND StatusOE<>0" & _
                " AND IdUnidadAdministrativa = " & iif(objs.TipoSeccionK = EDK_TIPOEST, iif(regGP.IdUnidadAdministrativa_EDP > 0, regGP.IdUnidadAdministrativa_EDP, -1), 0)
            if khorEDoePorPeriodo() then sq = sq & " AND IdPeriodoOE = " & regGP.IdPeriodo
            if lstOE<>"" then sq = sq & " AND IdObjetivoEstrategico NOT IN (" & lstOE & ")"
            set rs = getrs(conn,sq)
            while not rs.eof
              auxOE.getFromRS rs
              set objR = new EDK_Rubro
              objR.IdEvaluacion = regGP.IdEvaluacion
              objR.IdSeccionK = objs.IdSeccionK
              objR.IdObjetivoEstrategico = rsNum(rs,"IdObjetivoEstrategico")
              OE2RubroK auxOE, objR
              objR.update conn
              auxid = objR.IdRubro
              colRubroK.add objR, auxid
              rs.movenext
            wend
            rs.close
            set rs = nothing
            set auxOE = nothing
        end if
      next
      set objc = nothing
      set objR = nothing
      set ObjS = nothing
    end if
  end sub

  '----------------------------------------

  sub processRequest(stage)
    dim auxo : set auxo = new EDK_Rubro
    dim lId : lId = reqCSV("rubroId")
    dim rId : rId = split( lId, "," )
    dim rPeso : rPeso = split( reqs("rubroPeso"), "," )
    dim lsaved : lsaved = ""
    dim lpeso : lpeso = ""
    dim s, r, objS, whistory
    dim edk_keepVersionResults : edk_keepVersionResults = (khorConfigValue(576,true) <> 0)
    dim auxs : auxs = getBDlist("IdEvaluacion","SELECT DISTINCT IdEvaluacion FROM EDK_Rubro WHERE IdRubro IN (" & lId & ")", false)
    if (trim(auxs) = "") OR (trim(auxs) = trim(regGP.IdEvaluacion)) then
    'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
    if (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas) then  'Revision final
      for s=1 to regEvaK.colSeccion.count
        set objs = regEvaK.colSeccion.obj(s)
        objs.sumResultado = 0
        if edkModoPesos() > 0 then
          sq = "SELECT PesoSeccionK FROM EDK_SeccionPeriodoPeso WHERE IdSeccionK = " & objS.IdSeccionK & _
              " AND IdPeriodo = " & regGP.IdPeriodo & " AND IdEntidadAux = " & edk_pesos_entidad_valor( regGP )
          objs.pesoPonderado = getBDnum( "PesoSeccionK", sq )
        end if
      next
    end if
    for r=lbound(rId) to uBound(rId)
      auxo.clean
      if not auxo.getFromDB(conn,rId(r)) then
        auxo.IdEvaluacion = regGP.IdEvaluacion
        auxo.IdSeccionK = reqn("rubroSec" & r)
        auxo.IdSeccionCompetencia = reqn("rubroCom" & r)
      end if
      set objs = regEvaK.colSeccion.objByKey(auxo.IdSeccionK)
      auxo.Rubro = reqs("rubroDes" & r)
      if objs.TipoSeccionK = EDK_TIPOCOT then   
        auxo.IdCompetencia = reqn("rubroIdC" & r)
        auxo.Rubro = getBD("Competencia","SELECT Competencia FROM vCompetencia WHERE IdCompetencia = " & auxo.IdCompetencia)
        auxo.NivelCompetencia = getBDnum("Nivel","SELECT Nivel FROM PuestoTecnicas WHERE IdPuesto=" & regGP.IdPuesto_EDK & " AND IdCompetencia=" & auxo.IdCompetencia)
      end if
      if auxo.Rubro <> "" then
        auxo.PesoRubro = getVal( rPeso(r), true )
        if stage = 0 then 'Establecimiento de objetivos
          auxo.IdSeccionTipoRubro = reqn("rubroTip" & r)
          auxo.Esperado = reqs("rubroEsp" & r)
          auxo.ModoCalculo = reqs("rubroMod" & r)
          auxo.FechaCompromiso = getDateFromDMYstr(request("rubroFC" & r))
          if NOT (objs.TipoSeccionK = EDK_TIPOORG OR objs.TipoSeccionK = EDK_TIPOEST) then
            auxo.IdObjetivoEstrategico = reqn("rubroOE" & r)
          end if
          auxo.logroMinimo = reqd("rubroLmin" & r)
          auxo.logroEsperado = reqd("rubroLesp" & r)
          auxo.logroMaximo = reqd("rubroLmax" & r)
          auxo.logroPolaridad = reqn("rubroLpol" & r)
        else
          if habAutoRevision then
            if personaSesion() = persona.idpersona then
              auxo.rev.getFromDB conn, auxo.IdRubro, idxRevision
            end if
            auxo.rev.ObsEvaluado = reqs("rubroObA" & r)
          elseif edk_useObsAdicional then
            auxo.rev.ObsEvaluado = reqs("rubroObA" & r)
          end if
          auxo.rev.Observaciones = reqs("rubroObs" & r)
          if objs.TipoSeccionK = EDK_TIPOLOG then
            auxo.rev.Logro = reqd("rubroLog" & r)
            auxo.rev.ResultadoLogro = reqd("rubroLres" & r)
            auxo.rev.Resultado = auxo.rev.ResultadoLogro * auxo.PesoRubro / 100.0
            auxo.rev.dirty = true
            objs.sumResultado = objs.sumResultado + auxo.rev.Resultado
          'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
          elseif (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas) then  'Revision final
            if NOT (habAutoRevision and (personaSesion() = persona.idpersona)) then
              if objs.allowEditResultado<>0 then
                auxo.rev.Resultado = reqd("rubroResTxt" & r)
                auxo.rev.dirty = true
              else
                auxo.rev.Calificacion = reqn("rubroCal" & r)
                auxo.rev.Resultado = escalaK.calResultadoX(auxo.rev.Calificacion,objs.IdSeccionK) * auxo.PesoRubro
              end if
            end if
            objs.sumResultado = objs.sumResultado + auxo.rev.Resultado
          end if
          auxo.rev.dirty = auxo.rev.dirty OR (auxo.rev.Observaciones <> "") OR (auxo.rev.ObsEvaluado <> "") OR (auxo.rev.Calificacion <> 0)
        end if
        whistory = (reqn("rubroHis"&r) <> 0)
        auxo.updateRev conn, idxRevision, whistory
        if whistory AND edk_keepVersionResults then
          '-- Recupera resultados de revisiones en la nueva version del rubro
          dim auxrev : set auxrev = new EDK_RubroRevision
          dim sq : sq = "SELECT * FROM EDK_RubroRevision WHERE IdRubro=" & rId(r)
          dim rsrev : set rsrev = getrs(conn,sq)
          while not rsrev.eof
            auxrev.getFromRS rsrev
            auxrev.IdRubro = auxo.IdRubro
            auxrev.update conn
            rsrev.movenext
          wend
          rsrev.close
          set rsrev = nothing
          set auxrev = nothing
        end if
        lsaved = strAdd( lsaved, ",", auxo.IdRubro )
        lpeso = strAdd( lpeso, ",", auxo.PesoRubro )
      end if
    next
    dim sumaResultado : sumaResultado = 0
    'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
    if (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas) then  'Revision final
      for s=1 to regEvaK.colSeccion.count
        set objs = regEvaK.colSeccion.obj(s)
        if edkModoPesos() = 0 then
          sumaResultado = sumaResultado + objs.sumResultado
        else
          sumaResultado = sumaResultado + (objs.sumResultado * objs.pesoPonderado / 100.0)
        end if
      next
      if regEvaK.modoCalK = 1 then
        sumaResultado = sumaResultado / regEvaK.colSeccion.count
      end if
    end if
    set objs = nothing
    if lsaved<>"" then
      conn.execute "UPDATE EDK_Rubro SET IdRubroActual=-1, FechaUltimoCambio=" & formatDateSQL(Now,true) & _
                  " WHERE IdEvaluacion=" & regGP.IdEvaluacion & " AND IdRubroActual=0 AND IdRubro NOT IN (" & lsaved & ")"
      if stage = -1 then
        if edkForceAD then
          regGP.AccionesD_EDK = reqs("accionesdTxt")
        else
          regGP.AccionesD_EDK = ""
        end if
        conn.execute "DELETE FROM EDK_Comentario WHERE IdEvaluacion=" & regGP.IdEvaluacion
        for r=1 to colComment.count
          set auxo = colComment.obj(r)
          auxo.desc = reqs("comment" & auxo.key)
          if auxo.desc <> "" then
            conn.execute "INSERT INTO EDK_Comentario (IdEvaluacion,TipoComentario,IdPersona,Comentario) VALUES (" & regGP.IdEvaluacion & "," & auxo.key & ",0,'" & sqsf(auxo.desc) & "')"
          end if
        next
      end if
      regGP.Fecha_EDK = Now
      dim strGP : strGP = "Fecha_EDK=" & formatDateSQL(regGP.Fecha_EDK,true)
      if regGP.Status_EDK < idxRevision then
        regGP.Status_EDK = idxRevision
        strGP = strGP & ", Status_EDK=" & regGP.Status_EDK
      end if
      'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
      if (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas) then
        regGP.Resultado_EDK = sumaResultado + resEvento
        if regGP.Resultado_EDK < 0 then regGP.Resultado_EDK = 0
        if regGP.Resultado_EDK > escalaK.maxEscala and not eval(edk_allowexcess) then regGP.Resultado_EDK = escalaK.maxEscala
        strGP = strGP & ", Resultado_EDK=" & regGP.Resultado_EDK
      end if
      if stage = -1 then
				strGP = strGP & ", AccionesD_EDK='" & sqsf(regGP.AccionesD_EDK) & "'"
      end if
      if regEvaK.conObjetivoEstrategico and regGP.IdUnidadAdministrativa_EDP>0 then
        strGP = strGP & ", IdUnidadAdministrativa_EDP=" & regGP.IdUnidadAdministrativa_EDP
      end if
      conn.execute "UPDATE ED_GrupoPersona SET " & strGP & " WHERE IdEvaluacion=" & regGP.IdEvaluacion
      logAcceso LOG_CAMBIO, lblED_EDK_titulo, " IdEvaluacion (" & regGP.IdEvaluacion & ") numRevision(" & idxRevision & ") IdRubro (" & lsaved & ") P (" & lpeso & ")"
      if (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas) then
        khorBorraCompatibilidadPersonaRubro conn, regGP.IdEvaluado, tipoCR_EDK, 0, ""
      end if
    end if
    end if
    set auxo = nothing
  end sub
  
  '----------------------------------------

  function jsonStrSeccion(objS)
    dim auxtr : auxtr = "["
    if objS.conTipoRubro then
      dim j
      for j=1 to objS.colTipoRubro.count
        set objTR = objS.colTipoRubro.obj(j)
        auxtr = auxtr & iif(j=1,"",", ") & "{ ""id"":""" & objTR.IdSeccionTipoRubro & """"
        auxtr = auxtr & ", ""title"":""" & strJS(objTR.TipoRubro) & """"
        auxtr = auxtr & ", ""minR"":""" & objTR.minRubrosTipo & """" 
        auxtr = auxtr & ", ""maxR"":""" & objTR.maxRubrosTipo & """" 
        auxtr = auxtr & ", ""numR"":""" & 0 & """"
        auxtr = auxtr & "}"
      next
    end if
    auxtr = auxtr & "]"
    dim retval : retval = "{ ""id"":""" & objS.IdSeccionK & """"
    retval = retval & ", ""title"":""" & strJS(objS.SeccionK) & """"
    retval = retval & ", ""tipo"":""" & objS.TipoSeccionK & """"
    retval = retval & ", ""minR"":""" & objS.MinRubros & """"
    retval = retval & ", ""maxR"":""" & objS.MaxRubros & """"
    retval = retval & ", ""minP"":""" & objS.MinPesoTotal & """"
    retval = retval & ", ""maxP"":""" & objS.MaxPesoTotal & """"
    retval = retval & ", ""fixP"":""" & objS.PesoFijoRubro & """"
    retval = retval & ", ""allEdRes"":""" & objS.allowEditResultado & """"
    retval = retval & ", ""msgR"":""" & strJS(seccionLeyendaRubros(objS)) & """"
    retval = retval & ", ""msgP"":""" & strJS(seccionLeyendaPesos(objS)) & """"
    retval = retval & ", ""numR"":""" & 0 & """"
    retval = retval & ", ""sumR"":""" & 0 & """"
    retval = retval & ", ""lblRubro"":""" & strJS(objS.lblRubro) & """"
    retval = retval & ", ""lblEsperado"":""" & strJS(objS.lblEsperado) & """"
    retval = retval & ", ""lblModoCalculo"":""" & strJS(objS.lblModoCalculo) & """"
    retval = retval & ", ""lblTipoRubro"":""" & strJS(objS.lblTipoRubro) & """"
    retval = retval & ", ""lblObjetivoEstrategico"":""" & strJS(objS.lblObjetivoEstrategico) & """"
    retval = retval & ", ""lblFechaCompromiso"":""" & strJS(objS.lblFechaCompromiso) & """"
    retval = retval & ", ""arrTR"": " & auxtr
    retval = retval & ", ""peso"":""" & objS.pesoPonderado & """"
    retval = retval & ", ""minPR"":""" & objS.minPesoRubro & """"
    retval = retval & ", ""maxPR"":""" & objS.maxPesoRubro & """"
    retval = retval & "}"
    jsonStrSeccion = retval
  end function
  
  sub jsEdicionEDK(stage)
    dim auxo
    dim s, objS
    dim e, oe %>
  <script languaje="javascript">
    function escalaShowToolTip(ids,idx) {
      if (typeof(myEscalaShowToolTip) == typeof(Function)) myEscalaShowToolTip(ids,idx);
    }
    function strResultado(res) {
      var s = ''; <%
      FOR e=1 to escalaK.colDetalle.count
        SET oe = escalaK.colDetalle.obj(e) %>
      if ( res >= <%=oe.MinResultado%> ) s = '<%=strJS(oe.CalTitulo)%>'; <%
      NEXT %>
      return s;
    }
    function calResultado(cal,ids) {
      var retval = 0;
      if (typeof(myCalResultado) == typeof(Function)) {
        retval = myCalResultado(cal,ids);
      } else { <%
      '-- function myCalResultado must be defined somewhere where edk_revision.asp (or its surrogate) can use it
      IF escalaK.MaxResultado <> 0 THEN %>
        retval = cal / 100; <%
      ELSE
        FOR e=1 to escalaK.colDetalle.count
          SET oe = escalaK.colDetalle.obj(e) %>
        if ( cal == <%=oe.Calificacion%> ) retval = <%=(oe.CalResultado/100)%>; <%
        NEXT
      END IF %>
      }
      return retval;
    }
    function objCompetencia(id,nom,peso,descom,desadd) {
      this.id = id
      this.nom = nom
      this.peso = peso
      this.descom = descom
      this.desadd = desadd
    }
    var arrCom = new Array();
    arrCom.push( new objCompetencia(0,'',0,'','') );
    var arrSecc = new Array(); <%
    for s=1 to regEvaK.colSeccion.count
      set objS = regEvaK.colSeccion.obj(s)
      if objS.TipoSeccionK = EDK_TIPOCOT then
        for j=1 to objS.colCompetencia.count
          set objC = objS.colCompetencia.obj(j) %>
    arrCom.push( new objCompetencia(<%=objC.IdSeccionCompetencia%>,'<%=strJS(objC.SeccionCompetencia)%>',<%=objC.PesoCompetencia%>,'<%=strJS(replace(objC.DescripcionCompetencia,vbCRLF,"<br />"))%>','<%=strJS(replace(objC.DescripcionAdicional,vbCRLF,"<br />"))%>') ); <%
        next
      end if %>
    arrSecc.push( <%=jsonStrSeccion(objS)%> ); <%
    next
    set objS = nothing %>
    function idxTipoRubro(idxs,idtr) {
      for (var t=0; t<arrSecc[idxs].arrTR.length; t++) if (arrSecc[idxs].arrTR[t].id == getNumero(idtr,'int')) break;
      return t;
    }
    function seccionConTipo(idxs) {
      return (arrSecc[idxs].lblTipoRubro != '') && (arrSecc[idxs].arrTR.length > 0);
    }
    function idxSeccion(idsecc) {
      for (var s=0; s<arrSecc.length; s++) if (arrSecc[s].id == getNumero(idsecc,'int')) break;
      return s;
    }
    function seccionConPeso(idxs) {
      return (arrSecc[idxs].minP > 0 || arrSecc[idxs].maxP > 0);
    }
    var stage = <%=stage%>;
    function verificaItem(idsecc,des,r,peso,edkMsg) {
      var idxs = idxSeccion(idsecc);
      if (isWhitespace(des.value) && (getNumero(peso.value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>') != 0)) {
        edkMsg.push(arrSecc[idxs].title + ': <%=strJS(strLang(lblEDK_RubroVacioCon_X,lblEDK_Peso))%>');
        des.focus();
        return false;
      }
      if ( !isWhitespace(des.value) ) {
        var obj;
        var pesoval = getNumero(peso.value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>');
        if (seccionConPeso(idxs)) {
          if (pesoval == 0) {
            edkMsg.push(arrSecc[idxs].title + ': <%=strJS(strJS(strLang(lblFRS_ElDato_X_EsRequerido,lblEDK_Peso)))%>');
            peso.focus();
            return false;
          }
          if (pesoval < arrSecc[idxs].minPR) {
            edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_Dato_X_DebeSerNumeroMayorIgual_Y)%>','<%=strJS(lblEDK_Peso)%>',arrSecc[idxs].minPR) );
            peso.focus();
            return false;
          }
          if (pesoval > arrSecc[idxs].maxPR) {
            edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_Dato_X_DebeSerNumeroMenorIgual_Y)%>','<%=strJS(lblEDK_Peso)%>',arrSecc[idxs].maxPR) );
            peso.focus();
            return false;
          }
        }
        if (stage==0) {
          if ( arrSecc[idxs].lblObjetivoEstrategico != '' ) {
            if ( getValor('rubroOE'+r,'int')==0 ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_ElDato_X_EsRequerido)%>',arrSecc[idxs].lblObjetivoEstrategico) );
              return false;
            }
          }
          if ( seccionConTipo(idxs) ) {
            obj = MM_findObj('rubroTip'+r);
            if ( getNumero(obj.value,'int') == 0 ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_ElDato_X_EsRequerido)%>',arrSecc[idxs].lblTipoRubro) );
              obj.focus();
              return false;
            }
          }
          if ( arrSecc[idxs].lblFechaCompromiso != '' ) {
            obj = MM_findObj('rubroFC'+r);
            if ( isWhitespace(obj.value) ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_ElDato_X_EsRequerido)%>',arrSecc[idxs].lblFechaCompromiso) );
              obj.focus();
              return false;
            }
          }
          if ( arrSecc[idxs].lblEsperado != '' ) {
            obj = MM_findObj('rubroEsp'+r);
            if ( isWhitespace(obj.value) && (<%=iif(edk_lblEsperado_jsEmptyRejectCondition="","true",edk_lblEsperado_jsEmptyRejectCondition)%>) ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_ElDato_X_EsRequerido)%>',arrSecc[idxs].lblEsperado) );
              obj.focus();
              return false;
            }
          }
          if ( arrSecc[idxs].lblModoCalculo != '' ) {
            obj = MM_findObj('rubroMod'+r);
            if (isWhitespace(obj.value) && (<%=iif(edk_lblModoCalculo_jsEmptyRejectCondition="","true",edk_lblModoCalculo_jsEmptyRejectCondition)%>) ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_ElDato_X_EsRequerido)%>',arrSecc[idxs].lblModoCalculo) );
              obj.focus();
              return false;
            }
          } <%
          '-- Variable sucia, de los dos if's anteriores
          'edk_lblEsperado_jsEmptyRejectCondition = "arrSecc[idxs].tipo != " & EDK_TIPOLOG '-- Condicion (javascript) para considerar invalido un valor vacío
          'edk_lblModoCalculo_jsEmptyRejectCondition = "arrSecc[idxs].tipo != " & EDK_TIPOLOG '-- Condicion (javascript) para considerar invalido un valor vacío
          IF lblEDK_logroMinimo<>"" AND lblEDK_logroMaximo<>"" THEN %>
          if (arrSecc[idxs].tipo == <%=EDK_TIPOLOG%>) {
            var lmin = getNumero($("#rubroLmin"+r).val(),"float");
            var lesp = getNumero($("#rubroLesp"+r).val(),"float");
            var lmax = getNumero($("#rubroLmax"+r).val(),"float");
            if ( !((lmin < lesp && lesp < lmax) || (lmin > lesp && lesp > lmax)) ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + '<%=strJS(lblEDK_logroRangoInvalido)%>' );
              des.focus();
              return false;
            }
          } <%
          END IF %>
        } else {
          if (seccionConPeso(idxs)) {  <%
            '-- Variable sucia, declarar en khorlabelespecial.asp o similar
            'edk_obsrubroopcional = true '-- Hace opcional el registro de los comentarios por rubro
            IF NOT edk_obsrubroopcional THEN %>
            obj = MM_findObj('<%=iif(habAutoRevision AND personaSesion() = persona.idpersona, "rubroObA", "rubroObs")%>'+r);
            if ( obj && isWhitespace(obj.value) && (arrSecc[idxs].lblRubro != '') ) {
              edkMsg.push(arrSecc[idxs].title + ': ' + '<%=strJS(strLang(lblFRS_ElDato_X_EsRequerido,lblEDK_Observaciones))%>' );
              obj.focus();
              return false;
            } <%
            END IF %>
            if ( ((stage == -1)<%=iif(edk_parciales_cuantitativas," || (stage > 0)","")%>) && (arrSecc[idxs].tipo != <%=EDK_TIPOLOG%>)) {
              obj = MM_findObj('rubroCal'+r);
              if ( obj && obj.value == '' ) {
                edkMsg.push(arrSecc[idxs].title + ': ' + '<%=strJS(strLang(lblFRS_ElDato_X_EsRequerido,lblEDK_Calificacion))%>' );
                obj.focus();
                return false;
              } 
              if (arrSecc[idxs].allEdRes != 0) {
                obj = MM_findObj('rubroResTxt'+r);
                var lims = getNumero(peso.value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>');
                var limi = (arrSecc[idxs].allEdRes < 0) ? (-1 * lims) : 0;
                var resv = getNumero(obj.value,'float');
                if ( (resv < limi) || (resv > lims) ) {
                  edkMsg.push(arrSecc[idxs].title + ': ' + strLang('<%=strJS(lblFRS_Dato_X_DebeSerNumeroEntre_Y_Z)%>','<%=strJS(lblEDK_Resultado)%>',limi,lims) );
                  obj.focus();
                  return false;
                }
              }
            }
          }
        }
      }
      return true;
    }
    function verificaRubros() {
      var edkMsg = new Array();
      var ok = true;
      var pobj = $("input[name=rubroPeso]");
      var suma = 0;
      var s, r, t, idsec, idxs, dobj;
      var invisibler = false;
      for (s=0; s<arrSecc.length; s++) {
        arrSecc[s].numR = 0;
        for (t=0; t<arrSecc[s].arrTR.length; t++) arrSecc[s].arrTR[t].numR = 0;
        if (arrSecc[s].lblRubro == '') invisibler = true;
      }
      // Datos
      for (r=0; r<pobj.length; r++) {
        idsec = getValor('rubroSec'+r,'int');
        idxs = idxSeccion(idsec);
        dobj = MM_findObj('rubroDes'+r);
        if (!isWhitespace(dobj.value)) { <%
          IF edkModoPesos() = 0 THEN %>
          if (getValor('secExcPeso'+idsec,'int') == 0) {
            suma += getNumero(pobj[r].value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>'); 
          } <%
          END IF %>
          arrSecc[idxs].numR++;
          if (stage==0) {
            t = getValor('rubroTip'+r,'int');
            if ( seccionConTipo(idxs) && t > 0 ) {
              arrSecc[idxs].arrTR[idxTipoRubro(idxs,t)].numR++;
            }
          }
        }
      }
      for (r=0; r<pobj.length; r++) {
        idsec = getValor('rubroSec'+r,'int');
        dobj = MM_findObj('rubroDes'+r);
        if ( !verificaItem(idsec,dobj,r,pobj[r],edkMsg) ) { 
          ok = false; 
          break; 
        } 
      } <%
      IF edkModoPesos() = 0 THEN %>
      // Peso total
      if (ok && (suma != 100) && !invisibler) {
        edkMsg.push('<%=strJS(strLang(lblEDK_ElPesoDebSumar_X,"100"))%>');
        ok = false;
      } <%
      END IF %>
      // Secciones
      var peso, auxs;
      for (s=0; s<arrSecc.length; s++) {
        if ( (arrSecc[s].numR < arrSecc[s].minR) && (arrSecc[s].msgR != '') ) {
          ok = false; edkMsg.push( arrSecc[s].title + ': ' + arrSecc[s].msgR );
        }
        peso = getValor('seccionPeso'+arrSecc[s].id,'<%=iif(regEvaK.PesoDecimal,"float","int")%>'); <%
        IF edkModoPesos() <> 0 THEN %>
        if ( peso != 100 ) {
          ok = false; edkMsg.push( arrSecc[s].title + ': <%=strJS(strLang(lblEDK_ElPesoDebSumar_X,"100"))%>' );
        } <%
        ELSE %>
        if ( ((peso < arrSecc[s].minP) || (peso > arrSecc[s].maxP)) && (arrSecc[s].msgP != '') ) {
          ok = false; edkMsg.push( arrSecc[s].title + ': ' + arrSecc[s].msgP );
        } <%
        END IF %>
        if ((stage==0) && seccionConTipo(s) ) {  // Tipo de rubro
          auxs = '';
          for (t=0; t<arrSecc[s].arrTR.length; t++) {
            if ((arrSecc[s].arrTR[t].numR < arrSecc[s].arrTR[t].minR) || (arrSecc[s].arrTR[t].numR > arrSecc[s].arrTR[t].maxR)) {
              auxs = strAdd( auxs, ', ', strLang( '<%=strJS(lblEDK_Entre_A_y_B_con_Y_Z)%>', arrSecc[s].arrTR[t].minR, arrSecc[s].arrTR[t].maxR, arrSecc[s].lblTipoRubro, arrSecc[s].arrTR[t].title ) );
            }
          }
          if ( auxs != '' ) {
            ok = false; edkMsg.push( arrSecc[s].title + ': ' + auxs );
          }
        }
      } <%
      IF stage=-1 THEN
        'Acciones de desarrollo
        IF edkForceAD THEN %>
      dobj = MM_findObj('accionesdTxt');
      if ( (getValor('resumenBareTotal','float') < <%=regGP.MinResAD_EDK%>) && isWhitespace(dobj.value) ) {
        edkMsg.push('<%=strJS(lblFRS_FaltaIngresar_ & lblEDK_AccionesD)%>');
        if (ok) dobj.focus();
        ok = false;
      } <%
        END IF
        'Comentarios
        FOR i=1 to colComment.count
          set auxo = colComment.obj(i)
          if (auxo.aux <> "optional") AND (auxo.aux <> "readonly") then %>
      dobj = MM_findObj('comment<%=auxo.key%>');
      if ( isWhitespace(dobj.value) ) {
        edkMsg.push('<%=strJS(lblFRS_FaltaIngresar_ & regEvaK.edkCommentLabel(auxo.key))%>');
        if (ok) dobj.focus();
        ok = false;
      } <%
          end if
        NEXT
      END IF %>
      var msg = '';
      for (s=0; s<edkMsg.length; s++) msg = strAdd( msg, '\n', edkMsg[s] )
      if (msg != '') alert(msg);
      return ok;
    }
    function pTotaliza(ids) {
      var pobj = $("input[name=rubroPeso]");
      var suma = 0;
      var sres = 0;
      for (var r=0; r<pobj.length; r++)
        if (getValor('rubroSec'+r,'int') == parseInt(ids,10)) {
          suma += getNumero(pobj[r].value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>');
          sres += getValor('rubroLres'+r,'float');
        }
      setValor( 'seccionPeso'+ids, suma );
      return suma;
    }
    function pChanged(obj,idx) {
      var idsec = getValor('rubroSec'+idx,'int');
      var idxs = idxSeccion(idsec);
      valida(obj,'<%=iif(regEvaK.PesoDecimal,"float","int")%>',0,100);
      if (arrSecc[idxs].tipo == <%=EDK_TIPOLOG%>) {
        lRecalcula(idx);
      }
      pTotaliza(idsec);
      setDirty(); <%
      IF calRevision > 0 THEN %>
      setValor('rubroHis'+idx,1); <%
      END IF %>
      return false;
    }
    function lRecalcula(idx) {
      var peso = getNumero($("input[name=rubroPeso]")[idx].value,"float");
      var logro = getNumero($("#rubroLog"+idx).val(),"float");
      var lesp = getNumero($("#rubroLesp"+idx).val(),"float");
      var res = null; <%
      IF lblEDK_logroMinimo<>"" AND lblEDK_logroMaximo<>"" THEN %>
      var lmin = getNumero($("#rubroLmin"+idx).val(),"float");
      var lmax = getNumero($("#rubroLmax"+idx).val(),"float");
      if ((lmin < lesp && lesp < lmax) || (lmin > lesp && lesp > lmax)) {
        res = 0;
        var factor;
        if (lmin < lmax) {
          factor = (25.0/((logro<lesp)?lesp-lmin:lmax-lesp));
          res = ((logro<lmin)?75.0:100.0) + (factor * (((logro>lmax)?lmax:logro)-((logro<lmin)?lmin:lesp)));
        } else if (lmin > lmax) {
          factor = (25.0/((logro>lesp)?lesp-lmin:lmax-lesp));
          res = ((logro>lmin)?75.0:100.0) + (factor * (((logro<lmax)?lmax:logro)-((logro>lmin)?lmin:lesp)));
        }
      } <%
      ELSEIF lblEDK_logroMinimo<>"" THEN %>
      var lmin = getNumero($("#rubroLmin"+idx).val(),"float");
      var pol = getValor("rubroLpol"+idx,'int');
      if (lesp==0) {
        res = (logro == 0) ? 100 : 0;
      } else if (pol != 0) {
        res = 100 * ( (logro!=0) ? (lesp/logro) : 1 );
      } else {
        res = 100 * ( (lesp!=0) ? (logro/lesp) : 0 );
      }
      if (res > <%=logroMaxResultado%>) res = <%=logroMaxResultado%>;
      if (res < 0) res = 0; <%
      END IF %>
      if (res != null) {
        if (res < 0) res = 0;
        setValor('rubroLres'+idx,res); <%
        IF stage <> 0 THEN %>
        $("#rubroLresDiv"+idx).html(res.toFixed(<%=numDec%>)); <%
        END IF %>
      }
    }
    function lChanged(obj,idx) {
      var auxval = stripCharsNotInBag(obj.value,'0123456789.-');
      if (obj.value != auxval) obj.value = auxval;
      if ( !valida(obj,'float') ) obj.value='';
      setDirty(); <%
      IF calRevision > 0 THEN %>
      if (obj.name != ('rubroLog'+idx)) setValor('rubroHis'+idx,1); <%
      END IF %>
      lRecalcula(idx);
      cRecalcula();
    }
    function tChanged(obj,idx,his) {
      setDirty();
      if (his) setValor('rubroHis'+idx,1);
      return false;
    }
    function cidChanged(idx,his) {
      tChanged(null,idx,his);
      var idc = getValor('rubroIdC'+idx,'int');
      var sid = getValor('rubroSec'+idx,'int');
      var sidx = idxSeccion(sid);
      var pobj = $("input[name=rubroPeso]");
      var cidx;
      for (cidx=0; cidx<arrCom.length; cidx++) if (arrCom[cidx].id == idc) break;
      setValor( "rubroDes"+idx, arrCom[cidx].nom );
      $('#rubroAddDiv'+idx).html( arrCom[cidx].desadd );
      $('#rubroEjeDiv'+idx).html( arrCom[cidx].descom );
      var c, obj, aux;
      var nc = 0;
      var lc = "";
      for (c=0; c<pobj.length; c++) {
        if ((obj = MM_findObj("rubroIdC"+c)) && (obj.value > 0)) {
          if ( (c != idx) && (obj.value == idc) ) {
            obj.value = 0;
            cidChanged(c,his);
          } else if ( getValor('rubroSec'+c,'int') == sid) {
            nc++;
            lc += ((lc=="")?"":",") + c;
          }
        }
      }
      if ( arrSecc[sidx].fixP == 0 ) {
        if (idc == 0) pobj[idx].value = 0;
        if (nc > 0) {
          var p = arrSecc[sidx].maxP / nc;
          p = p.toFixed(<%=numDec%>);
          var xs = arrSecc[sidx].maxP - (p * nc);
          xs = xs.toFixed(<%=numDec%>);
          obj = lc.split(",");
          for (c=0; c<obj.length; c++) { 
            aux = parseFloat(p) + parseFloat(((c==0)?xs:0));
            pobj[obj[c]].value = aux.toFixed(<%=numDec%>);
          }
        }
        pTotaliza(sid);
      }
      return false;
    }
    function cRecalcula() {
      var robj = $("input[name=rubroRes]");
      var pobj = $("input[name=rubroPeso]");
      var ids, s, cal, rubroResTxt;
      for (s=0; s<arrSecc.length; s++) arrSecc[s].sumR = 0;
      for (r=0; r<pobj.length; r++) {
        ids = getValor('rubroSec'+r,'int');
        s = idxSeccion(ids);
        if ( seccionConPeso(s) ) {
          if (arrSecc[s].tipo == <%=EDK_TIPOLOG%>) {
            cal = getValor('rubroLres'+r,'float') * getNumero(pobj[r].value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>') / 100.0; <%
            IF stage <> 0 THEN %>
            MM_setTextOfLayer('rubroResDiv'+r,0,cal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>'); <%
            END IF %>
          } else {
            rubroResTxt = MM_findObj('rubroResTxt'+r);
            if (rubroResTxt) {
              cal = getNumero(rubroResTxt.value,'float');
            } else {
              cal = calResultado(getValor('rubroCal'+r,'int'),ids) * getNumero(pobj[r].value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>');
              MM_setTextOfLayer('rubroResDiv'+r,0,cal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>');
            }
            robj[r].value = cal;
          }
          arrSecc[s].sumR += cal;
        }
      }
      var auxRes;
      var sumatotal = 0;
      var numsecc = 0;
      for (s=0; s<arrSecc.length; s++) {
        cal = arrSecc[s].sumR;
        MM_setTextOfLayer('seccionRes'+arrSecc[s].id,0, cal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>');
        MM_setTextOfLayer('resumen'+arrSecc[s].id,0, cal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>'); <%
        IF edkModoPesos() <> 0 THEN %>
        cal = (cal * arrSecc[s].peso / 100.0);
        MM_setTextOfLayer('restotal'+arrSecc[s].id,0, cal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>'); <%
        END IF %>
        sumatotal += cal;
        if ( seccionConPeso(s) ) numsecc++;
      } <%
      IF regEvaK.modoCalK = 1 THEN %>
      sumatotal = sumatotal / numsecc; <%
      END IF %>
      setValor('resumenBareTotal',sumatotal); <%
      IF hayEvento THEN %>
      sumatotal += getValor('totEvento','float');
      if (sumatotal < 0) sumatotal = 0;
      if (sumatotal > <%=escalaK.maxEscala%>) sumatotal = <%=escalaK.maxEscala%>; <%
      END IF %>
      MM_setTextOfLayer('resumenTotal',0, sumatotal.toFixed(<%=numDec%>) + '<%=lblEDK_PercentSign%>');
      MM_setTextOfLayer('resumenEscala',0, strResultado(sumatotal));
    }
    function cChanged(obj,idx) { <%
      IF escalaK.MaxResultado <> 0 THEN %>
      if ( !valida(obj,'int',<%=escalaK.minCalificacion%>,<%=escalaK.maxEscalaCal()%>) ) obj.value=''; <%
      END IF %>
      setDirty();
      cRecalcula();
    }
    function cRadioSet(idx,v) {
      var obj = MM_findObj("rubroCal"+idx);
      obj.value = v;
      setRadio("radioCal"+idx,v);
      cChanged(obj,idx);
    }
    function rChanged(obj,idx,dir) {
      var pobj = MM_findObj('rubroPeso');
      var peso = getNumero(pobj[idx].value,'<%=iif(regEvaK.PesoDecimal,"float","int")%>');
      if ( !valida(obj,'float',((dir<0)?(-1*peso):0),peso) ) obj.value=0;
      setDirty();
      cRecalcula();
    }
    function showHistory(idrub) {
      popUp_show('rubroHistory'+idrub);
      return false;
    }
    function setCalendar() {
      setDirty();
    }
    <%if stage=0 AND regEvaK.conObjetivoEstrategico then%>
    var oeIndex;
    function selOE(idx) {
      oeIndex = idx;
      var idOE = getValor('rubroOE'+oeIndex);
      if (idOE == 0) {
        $('input[name="OEradio"]').attr('checked', false);
      } else {
        setRadio('OEradio',idOE);
      }
      popUp_show('popUpSelOE');
      return false;
    }
    function setOE() {
      var idOE = getMultiSelCheck('OEradio');
      setValor( 'rubroOE'+oeIndex, idOE );
      $('#divOE'+oeIndex).html( $('#OEradioDisplay'+idOE).val() );
      popUp_hide('popUpSelOE');
    }
    <%end if%>
  </script> <%
  end sub
  
  '----------------------------------------

  sub evaluacionSelOE(stage,allowEdit)
    if regEvaK.conObjetivoEstrategico AND regGP.IdUnidadAdministrativa_EDP>0 AND not khorPage_isPDF() then
      dim colOE : set colOE = new frsCollection
      edkGlobalPeriodoOE = regGP.IdPeriodo
      edkGetObjetivosEstrategicos colOE, regGP.IdUnidadAdministrativa_EDP
      popUpBegin "popUpSelOE", lblKHOR_ObjetivoEstrategico, "", "" %>
      <style>
        div.oeScroll {
          height: 300px;
          width: 100%;
          overflow: auto;
          border: 1px solid #ddd;
          background-color: #fff;
          padding: 10px;
          text-align: center;
        }
      </style>
      <script>
        function oeShowToolTip(idx) {
          var txt = $("#oeDisplay"+$("#rubroOE"+idx).val()).html();
          setToolTipDimensions(200,Math.ceil(txt.length/40)*15,10);
          showToolTip(txt, true);
        }
      </script>
      <div class="oeScroll">
      <% edkPaintObjetivosEstrategicos(colOE) %>
      </div>
      <div align="center">
        <INPUT type="button" value="<%=lblFRS_Aceptar%>"  onclick="setOE();" class="whitebtn" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this); onmouseout=inOut(this);">
        <INPUT type="button" value="<%=lblFRS_Cancelar%>" onclick="popUp_hide('popUpSelOE')" class="whitebtn" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this); onmouseout=inOut(this);">
      </div> <%
      popUpEnd
      set colOE = nothing
    end if
  end sub

  '----------------------------------------

  sub evaluacionHistory()
    dim s, objS, r, objR, estilo, h
    for s=1 to regEvaK.colSeccion.count
      set objS = regEvaK.colSeccion.obj(s)
      for r=1 to colRubroK.count
        set objr = colRubroK.obj(r)
        if objr.IdSeccionK = objS.IdSeccionK then
          if objr.history.count > 0 then
            popUpBegin "rubroHistory"&objr.idrubro, lblFRS_Historial, khorWinWidthPix()&"px", "" %>
            <table border="0" cellspacing="0" cellpadding="1" align="center" id="edRK" style="page-break-inside:avoid;"> <%
            seccionTableTitles 0, objS, true, true
            seccionTableRow 0, objS, objR, -1, false, false, estilo
            for h=1 to objr.history.count
              estilo = switchEstilo(estilo)
              seccionTableRow 0, objS, objR.history.obj(h), -1-h, false, false, estilo
            next %>
            </table> <%
            popUpEnd
          end if
        end if
      next
    next
  end sub
  
  '----------------------------------------

  sub evaluacionComments(stage)
    dim i, auxo, w
    '-- Variable sucia, declarar en khorlabelespecial.asp o edk_especial.asp
    'edk_commentsXrow = 2 '-- numero de comentarios por renglon
    if edk_commentsXrow > 0 then
      cxr = edk_commentsXrow
    else
      cxr = colComment.count
    end if
    if stage=-1 and colComment.count > 0 then
      w = floor( 100 / cxr ) %>
    <br />
    <table border="0" cellspacing="2" cellpadding="0" align="center" style="width:100%;page-break-inside:avoid;">
      <tr> <%
      for i=1 to colComment.count
        set auxo = colComment.obj(i) %>
        <td width="<%=w%>%" class="edRKborder" valign="top">
          <table border="0" cellspacing="0" cellpadding="1" align="center"  style="width:100%;">
            <tr class="celdaTit"><td><%=regEvaK.edkCommentLabel(auxo.key)%></td></tr>
            <tr>
              <td> <%
                if allowEvaluate then %>
                <textarea id="comment<%=auxo.key%>" name="comment<%=auxo.key%>" rows="5" maxlength="4000" <%=iif(auxo.aux="readonly", "readOnly=""readOnly""", "onChange=""setDirty()""")%> style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(auxo.desc)%></textarea> <%
                else
                  response.write serverHTMlencode(auxo.desc)
                end if %>
              </td>
            </tr>
          </table>
        </td> <%
        if i<colComment.count and (i mod cxr) = 0 then
          response.write "</tr><tr>"
        end if
      next %>
      </tr>
    </table> <%
    end if
  end sub
    
  '----------------------------------------

  sub evaluacionEventos(stage,adminEventos)
    if hayEvento then
      if adminEventos then %>
    <script languaje="javascript">
      function refreshEvento(eid,response) {
        document.getElementById('divEventos').innerHTML = response; <%
        if stage = -1 then %>
        var tot = getValor('totEvento','float')
        document.getElementById('resumen0').innerHTML = (tot>0?'+':'') + tot.toFixed(<%=numDec%>) + '%';
        cRecalcula(); <%
        end if %>
      }
      function actionEvento(action,ideve) {
        var params = "mov="+action+"&ideve="+ideve+"&ideva=<%=regGP.IdEvaluacion%>&stage=<%=stage%>&adm=<%=bool2num(adminEventos)%>";
        var ajaxObj = createAjaxObject();
        callAjaxFunction(ajaxObj, "edk_ajax.asp", "evento", params, refreshEvento);
      }
      function borraEvento(ideve) {
        actionEvento('B', ideve);
        return false;
      }
      function agregaEvento() {
        var pobj = MM_findObj('ideventok');
        if (pobj.value == '0') {
          pobj.focus();
        } else {
          actionEvento('A', pobj.value)
        }
        return false;
      }
    </script> <%
      end if %>
    <div id="divEventos" align="center">
    <% edk_tablaEventos colEvento, regGP.IdEvaluacion, regGP.IdEvaluacionK, adminEventos %>
    </div>
    <br /> <%
    end if
  end sub
    
  '----------------------------------------

  function evaluacionResumenSimple()
    dim modoPesos : modoPesos = edkModoPesos()
    dim resTotal : resTotal = 0
    dim s, objS, auxRes  %>
    <table border="0" cellspacing="0" cellpadding="1" align="center">
      <tr class="celdaTit">
        <td colspan="<%=(2+iif(modoPesos<>0,2,0))%>" style="text-align:center;"><%=lblEDK_ResumenEvaluacion%></td>
      </tr> <%
      if modoPesos<>0 then %>
      <tr class="celdaTit">
        <td><%=lblEDK_Seccion%></td>
        <td><%=lblEDK_Resultado%></td>
        <td><%=lblEDK_Peso%></td>
        <td><%=lblEDK_CalificacionFinal%></td>
      </tr> <%
      end if
      dim numsecc : numsecc = 0
      for s=1 to regEvaK.colSeccion.count
        set objS = regEvaK.colSeccion.obj(s)
        if modoPesos = 0 then
          resTotal = resTotal + objS.sumResultado
        else
          auxRes = (objS.sumResultado * objS.pesoPonderado / 100.0)
          resTotal = resTotal + auxRes
        end if
        if objS.conPeso then %>
      <tr>
        <td><%=objS.SeccionK%></td>
        <td id="resumen<%=objS.IdSeccionK%>" style="text-align:right;"><%=formatNumber(objS.sumResultado,numDec)%><%=lblEDK_PercentSign%></td> <%
          if modoPesos<>0 then %>
        <td style="text-align:right;"><%=formatNumber(objS.pesoPonderado,numDec)%><%=lblEDK_PercentSign%></td>
        <td id="restotal<%=objS.IdSeccionK%>" style="text-align:right;"><%=formatNumber(auxRes,numDec)%><%=lblEDK_PercentSign%></td> <%
          end if %>
      </tr> <%
          numsecc = numsecc + 1
        end if
      next
      if modoPesos = 0 AND regEvaK.modoCalK = 1 then
        resTotal = div( resTotal, numsecc )
      end if %>
      <input type="hidden" id="resumenBareTotal" name="resumenBareTotal" value="<%=formatNumber(resTotal,numDec)%>"> <%
      if hayEvento then %>
      <tr>
        <td colspan="<%=iif(modoPesos<>0,3,1)%>">*<%=lblEDK_EventoTotal%></td>
        <td id="resumen0" style="text-align:right;"><%=iif(resEvento>0,"+","")%><%=formatNumber(resEvento,numDec)%><%=lblEDK_PercentSign%></td>
      </tr> <%
        resTotal = resTotal + resEvento
        if resTotal < 0 then resTotal = 0
        if resTotal > escalaK.maxEscala then resTotal = escalaK.maxEscala
      end if %>
      <tr class="celdaTit">
        <td colspan="<%=iif(modoPesos<>0,3,1)%>"><%=lblEDK_CalificacionFinal%></td>
        <td id="resumenTotal" style="text-align:right;"><%=formatNumber(resTotal,numDec)%><%=lblEDK_PercentSign%></td>
      </tr>
    </table> <%
    evaluacionResumenSimple = resTotal
  end function

  sub evaluacionResumen(stage,adminEventos)
    dim resTotal : resTotal = 0
    if not hideSelfResults then
    if stage=-1 then %>
    <br />
    <table border="0" cellspacing="0" cellpadding="5" align="center" style="width:90%;page-break-inside:avoid;"> <%
      if (regGP.IdEvaluado = personaSesion()) AND enFechasRevision(idxRevision) then %>
      <tr><td colspan="2" class="alerta"><%=lblEDK_AvisoResultadoNoDefinitivo%></td></tr> <%
      end if %>
      <tr>
        <td valign="middle">
        <% resTotal = evaluacionResumenSimple() %>
        </td>
        <td valign="middle" style="text-align:center;"> <% 
          evaluacionEventos stage, adminEventos
          if (escalaK.colDetalle.count > 0) AND NOT edk_resumenraw then %>
          <strong><%=lblEDK_Calificacion%></strong>
          <div id="resumenEscala" style="text-align:center;"><%=escalaK.strResultado(round(resTotal,numDec))%></div> <%
          end if %>
        </td>
      </tr>
    </table> <%
      if edkForceAD then %>
    <div id="divAD" style="width:100%;page-break-inside:avoid;"> <%
        dim msg: msg = "<b>" & lblEDK_AccionesD & " </b> "
        if allowEvaluate then
          if regGP.MinResAD_EDK > 0 then
            msg = msg & " <span class=""tsmall"">(" & strLang(lblEDK_AccionesDinstruccion,regGP.MinResAD_EDK) & ")</span>"
          end if %>
      <div class="celdaTit"><%=msg%></div> 
      <textarea id="accionesdTxt" name="accionesdTxt" rows="5" maxlength="8000" onChange="setDirty()" style="width:100%" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(regGP.AccionesD_EDK)%></textarea> <%
        else
          response.write "<div class=""celdaTit"">" & msg & "</div>" & vbCRLF
          response.write "<br/>" & serverHTMLencode(regGP.AccionesD_EDK)
        end if %>
    </div> <%
      end if
    else
      evaluacionEventos stage, adminEventos
    end if
    end if
  end sub
  
  '----------------------------------------

  private function seccionLeyendaRubros(objS)
    dim auxs : auxs = ""
    if objS.PesoFijoRubro=0 AND (objS.MinRubros<>0 OR objS.MaxRubros<>0) then
      if objS.MinRubros = objS.MaxRubros then
        auxs = objS.MaxRubros
      else
        auxs = strAdd( iif(objS.MinRubros<>0, lblFRS_Minimo & " " & objS.MinRubros,""), lblFRS_And_, iif(objS.MaxRubros<>0, lblFRS_Maximo & " " & objS.MaxRubros,"") )
      end if
      auxs = strLang( lblEDK_DebeRegistrar_X, auxs )
    end if
    seccionLeyendaRubros = auxs
  end function
  
  private function seccionLeyendaPesos(objS)
    dim auxs : auxs = ""
    if objS.PesoFijoRubro=0 AND (objS.MinPesoTotal<>0 OR objS.MaxPesoTotal<>0) AND not edk_hidepesos then
      if objS.MinPesoTotal = objS.MaxPesoTotal then
        auxs = objS.MaxPesoTotal
      else
        auxs = strAdd( iif(objS.MinPesoTotal<>0, lblFRS_Minimo & " " & objS.MinPesoTotal,""), lblFRS_And_, iif(objS.MaxPesoTotal<>0, lblFRS_Maximo & " " & objS.MaxPesoTotal,"") )
      end if
      auxs = strLang( lblEDK_ElPesoDebSumar_X, auxs )
    end if
    seccionLeyendaPesos = auxs
  end function
  
  sub seccionTotalRow(stage, objS, colSpan, allowEvaluate, allowEdit)
    dim sFixed : sFixed = ((objS.TipoSeccionK = EDK_TIPOORG) OR (objS.TipoSeccionK = EDK_TIPOEST))
    dim prevSpan, postSpan
    dim stageEval : stageEval = (stage = -1)
    if objS.TipoSeccionK = EDK_TIPOLOG then
      prevSpan = columnaPeso - 1
      postSpan = 0
    else
      'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
      stageEval = stageEval OR (stage > 0 AND edk_parciales_cuantitativas)
      postSpan = iif( stage=1 AND NOT stageEval, iif(objS.lblRevision<>"" AND not sFixed,1+bool2num((habAutoRevision AND NOT (personaSesion() = persona.idpersona)) OR (edk_useObsAdicional AND NOT habAutoRevision)),0), 0 )
      prevSpan = iif( stage=0, (colSpan-1), (colSpan - postSpan - bool2num(not edk_hidepesos) - bool2num(stageEval AND not hideSelfResults)) ) 
    end if
    dim auxs : auxs = ""
    if objS.conPeso AND (stageEval OR not edk_hidepesos) then
      if allowEdit AND (not sFixed) then
        auxs = seccionLeyendaPesos(objS)
      end if %>
          <tr>
            <td colspan="<%=prevSpan%>" style="text-align:right;font-weight:bold;"> <%
              if edk_hidepesos then
                response.write "&nbsp;"
              else
                response.write strAdd( lblEDK_Peso & ":", "<br />", auxs )
              end if %>
            </td> <%
            if edk_hidepesos then %>
            <input type="hidden" id="seccionPeso<%=objS.IdSeccionK%>" name="seccionPeso<%=objS.IdSeccionK%>" value="<%=objS.sumPeso%>"> <%
            else %>
            <td class="celdaTit" style="text-align:right;"> <%
              if allowEdit AND NOT sFixed then %>
              <input type="text" id="seccionPeso<%=objS.IdSeccionK%>" name="seccionPeso<%=objS.IdSeccionK%>" value="<%=objS.sumPeso%>" readOnly="readOnly" style="width:40px; text-align:right;" class="whiteblur"><%=lblEDK_PercentSign%> <%
              else
                response.write objS.sumPeso & "&nbsp;" & lblEDK_PercentSign %>
                <input type="hidden" id="seccionPeso<%=objS.IdSeccionK%>" name="seccionPeso<%=objS.IdSeccionK%>" value="<%=objS.sumPeso%>"> <%
              end if %>
            </td> <%
            end if
            if objS.TipoSeccionK = EDK_TIPOLOG then
              response.write "<td colspan=""" & (colSpan - columnaPeso - bool2num(stageEval)) & """>&nbsp;</td>"
            end if
            if not hideSelfResults then
              if stageEval then %>
            <td id="seccionRes<%=objS.IdSeccionK%>" class="celdaTit" style="text-align:right;"><%=formatNumber(objS.sumResultado,numDec)%><%=lblEDK_PercentSign%></td> <%
              end if
              if postSpan>0 then
                response.write "<td colspan=""" & postSpan & """>&nbsp;</td>"
              end if
            end if %>
          </tr> <%
    else %>
          <input type="hidden" id="seccionPeso<%=objS.IdSeccionK%>" name="seccionPeso<%=objS.IdSeccionK%>" value="<%=objS.sumPeso%>"> <%
    end if
  end sub

  '----------------------------------------
  
  private sub pesoCell(objS, objR, idx, allowEdit)
    dim sFixed : sFixed = ((objS.TipoSeccionK = EDK_TIPOORG) OR (objS.TipoSeccionK = EDK_TIPOEST))
    if objS.conPeso AND not edk_hidepesos then %>
            <td style="text-align:right;" nowrap> <%
              dim dvalue : dvalue = iif( objS.TipoSeccionK = EDK_TIPOLOG, formatNumber(objR.PesoRubro,numDec), objR.PesoRubro )
              if allowEdit AND objS.PesoFijoRubro=0 AND NOT sFixed then %>
              <input type="text" id="rubroPeso" name="rubroPeso" value="<%=dvalue%>" onChange="pChanged(this,<%=idx%>)" maxlength="<%=iif(regEvaK.PesoDecimal,(3+numDec),3)%>" style="width:<%=(40+iif(regEvaK.PesoDecimal,(5*numDec),0))%>px; text-align:right;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onFocus="inFocus(this);" onBlur="inBlur(this);"><%=lblEDK_PercentSign%> <%
              else
                response.write "<span id=""spanPeso" & idx & """>" & strAdd(dvalue, "&nbsp;", lblEDK_PercentSign) & "</span>"
                if idx >=0 then %>
              <input type="hidden" id="rubroPeso" name="rubroPeso" value="<%=objR.PesoRubro%>"> <%
                end if
              end if %>
            </td> <%
    elseif allowEdit or (idx >=0) then %>
              <input type="hidden" id="rubroPeso" name="rubroPeso" value="<%=objR.PesoRubro%>"> <%
    end if
  end sub

  '----------------------------------------

  private sub resultadoCell(objS, objR, idx, allowEvaluate, noShow)
    if NOT (edk_hideresrubro OR hideSelfResults) then %>
            <td style="text-align:right;"> <%
      if allowEvaluate AND objS.allowEditResultado<>0 then %>
              <input type="text" id="rubroResTxt<%=idx%>" name="rubroResTxt<%=idx%>" value="<%=formatNumber(objR.rev.Resultado,numDec)%>" onChange="rChanged(this,<%=idx%>,<%=objS.allowEditResultado%>)" maxlength="5" style="width:40px; text-align:right;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onFocus="inFocus(this);" onBlur="inBlur(this);"><%=lblEDK_PercentSign%> <%
      else
        dim dvalue : dvalue = iif( noShow, "-", formatNumber(objR.rev.Resultado,numDec) ) %>
              <div id="rubroResDiv<%=idx%>" style="text-align:right;"><%=dvalue%><%=lblEDK_PercentSign%></div> <%
      end if %>
            </td> <%
    end if
    if idx >=0 then %>
            <input type="hidden" id="rubroRes" name="rubroRes" value="<%=objR.rev.Resultado%>"> <%
    end if
  end sub

  '----------------------------------------

  private sub logroCell(idx, name, value, allowEdit, noShow, usePol, polValue)
    dim inputpainted : inputpainted = false
    if hideSelfResults then
      if idx >=0 then %>
              <input type="hidden" id="<%=(name&idx)%>" name="<%=(name&idx)%>" value="<%=value%>"> <%
      end if
    else %>
            <td style="text-align:right;"> <%
      if allowEdit then %>
                <input type="text" id="<%=(name&idx)%>" name="<%=(name&idx)%>" value="<%=formatNumber(value,numDec)%>" onChange="lChanged(this,<%=idx%>)" maxlength="<%=(12+numDec)%>" style="width:<%=(90+3*numDec)%>px; text-align:right;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onFocus="inFocus(this);" onBlur="inBlur(this);"> <%
        inputpainted = true
        if usePol then %><br/>
                <select id="rubroLpol<%=idx%>" name="rubroLpol<%=idx%>" onChange="setDirty()" style="font-size:10px;">
                <% optionFromMenuFijo menufijo_EDKpolaridad, polValue %>
                </select><%
        end if
      else
        dim dvalue
        if noShow then
          dvalue = "-"
        else
          dvalue = formatNumber(value,numDec)
        end if %>
              <div id="<%=("div"&name&idx)%>" style="text-align:right;"><%=dvalue%></div> <%
        if usePol then %>
              <div class="tsmall" style="text-align:right;"><%=descFromMenuFijo(menufijo_EDKpolaridad,polValue)%></div> <%
        end if
        if idx >=0 then %>
              <input type="hidden" id="<%=(name&idx)%>" name="<%=(name&idx)%>" value="<%=value%>"> <%
        end if
      end if %>
            </td> <%
    end if
    if usePol AND idx>=0 AND not inputpainted then %>
            <input type="hidden" id="rubroLpol<%=idx%>" name="rubroLpol<%=idx%>" value="<%=polValue%>"> <%
    end if
  end sub
  
  '----------------------------------------

  sub seccionTableRow(stage, objS, objR, idx, allowEvaluate, allowEdit, estilo)
    dim textrows : textrows = iif( ((objS.TipoSeccionK = EDK_TIPOCOM) OR (objS.TipoSeccionK = EDK_TIPOCOT)), 5, 3 )
    dim sFixed : sFixed = ((objS.TipoSeccionK = EDK_TIPOORG) OR (objS.TipoSeccionK = EDK_TIPOEST))
    'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
    dim stageEval : stageEval = (stage = -1) OR (objS.TipoSeccionK <> EDK_TIPOLOG AND stage > 0 AND edk_parciales_cuantitativas)
    dim auxo, auxtxt
    if idx >=0 then
      if calRevision > 0 then %>
          <input type="hidden" id="rubroHis<%=idx%>" name="rubroHis<%=idx%>" value="0"> <%
      end if %>
          <input type="hidden" id="rubroId" name="rubroId" value="<%=objR.IdRubro%>">
          <input type="hidden" id="rubroSec<%=idx%>" name="rubroSec<%=idx%>" value="<%=objR.IdSeccionK%>">
          <input type="hidden" id="rubroCom<%=idx%>" name="rubroCom<%=idx%>" value="<%=objR.IdSeccionCompetencia%>"> <%
    end if %>
          <tr class="<%=estilo%>"> <%
            '-- Fecha (historico)
            if idx < 0 then %>
            <td align="center"><%=formatDateDisp(objr.FechaRegistro,true)%></td> <%
            end if
            '-- Objetivo estrategico
            if objS.conObjetivoEstrategico then
              dim regOE : set regOE = new ed_ObjetivoEstrategico
              regOE.getFromDB conn, objr.IdObjetivoEstrategico
              auxt = edkDescObjetivosEstrategico(regOE,false)
              set regOE = nothing %>
            <td>
              <div id="divOE<%=idx%>" style="text-align:center;"<%=iif(idx<0 OR khorPage_isPDF(),""," onMouseOver=""oeShowToolTip(" & idx & ");"" onMouseOut=""hideToolTip();""")%>>
              <%=auxt%>
              </div> <%
              if idx >= 0 then %>
              <input type="hidden" id="rubroOE<%=idx%>" name="rubroOE<%=idx%>" value="<%=objR.IdObjetivoEstrategico%>"> <%
              end if
              if stage=0 AND allowEdit then %>
              <div class="tsmall" style="text-align:center;">[<a href="#" onClick="return selOE(<%=idx%>);"><%=lblFRS_Seleccionar%></a>]</div> <%
              end if %>
            </td> <%
            end if
            '-- Tipo de rubro
            if objS.conTipoRubro then %>
            <td id="tdRubroTip<%=idx%>"><%
              '-- variable sucia, declarar en khorlabelepecial o similar
              'edk_fixTipoRubro = true  '-- Los tipos de rubro estan preasignados y fijos
              if stage=0 AND allowEdit AND NOT (edk_fixTipoRubro AND objR.IdSeccionTipoRubro>0) then %>
              <select id="rubroTip<%=idx%>" name="rubroTip<%=idx%>" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="font-size:10px;">
                <option value="0">-</option> <%
                for i=1 to objS.colTipoRubro.count
                  set objTR = objS.colTipoRubro.obj(i)
                  if objTR.TipoRubro <> "" then %>
                <option value="<%=objTR.IdSeccionTipoRubro%>"<%=optionSelIf(objR.IdSeccionTipoRubro = objTR.IdSeccionTipoRubro)%>><%=objTR.TipoRubro%></option> <%
                  end if
                next %>
              </select> <%
              else
                response.write objS.tipoRubro(objR.IdSeccionTipoRubro)
                if idx >= 0 then %>
              <input type="hidden" id="rubroTip<%=idx%>" name="rubroTip<%=idx%>" value="<%=objR.IdSeccionTipoRubro%>"> <%
                end if
              end if
              '-- variable sucia, declarar en khorlabelepecial o similar
              'edk_txtTipoRubro = "texto adicional para la celda del tipo de rubro"
              if (""&edk_txtTipoRubro) <> "" then
                response.write edk_txtTipoRubro
              end if %>
            </td> <%
            end if
            '-- Rubro
            if edk_usetooltips AND ((objS.TipoSeccionK = EDK_TIPOCOM) OR ((objs.TipoSeccionK = EDK_TIPOCOT) AND objS.preloadCompetencias)) and objS.lblEjemplo<>"" then %>
            <td id="tdRubro<%=objR.IdRubro%>" onmouseover="edkToolTip(<%=objR.IdRubro%>,<%=iif(objS.TipoSeccionK = EDK_TIPOCOT, -1 * objR.IdCompetencia, objR.IdSeccionCompetencia)%>);" onMouseOut="hideToolTip();"> <%
            else %>
            <td> <%
            end if
              if allowEdit AND (objS.TipoSeccionK = EDK_TIPOCOT) AND NOT objS.preloadCompetencias then %>
              <select id="rubroIdC<%=idx%>" name="rubroIdC<%=idx%>" onChange="cidChanged(<%=idx%>,<%=bool2num(calRevision>0)%>)" style="font-size:10px;">
                <option value="0"></option> <%
                for i=1 to objS.colCompetencia.count
                  set objC = objS.colCompetencia.obj(i) %>
                <option value="<%=objC.IdSeccionCompetencia%>"<%=optionSelIf(objR.IdCompetencia=objC.IdSeccionCompetencia)%>><%=objC.SeccionCompetencia%></option> <%
                next %>
              </select>
              <input type="hidden" id="rubroDes<%=idx%>" name="rubroDes<%=idx%>" value="<%=serverHTMLencode(objR.Rubro)%>"> <%
              elseif allowEdit AND (objS.TipoSeccionK <> EDK_TIPOCOM) AND (objS.TipoSeccionK <> EDK_TIPOCOT) AND NOT sFixed then %>
              <textarea id="rubroDes<%=idx%>" name="rubroDes<%=idx%>" rows="<%=textrows%>" maxlength="2048" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.Rubro)%></textarea> <%
              else
                response.write replace( serverHTMLencode(objR.Rubro), vbCRLF, "<br />" )
                if (stage <> 0) AND ((objS.TipoSeccionK = EDK_TIPOCOM) OR (objS.TipoSeccionK = EDK_TIPOCOT)) and (objS.lblEjemplo<>"" OR objS.lblRubroAdicional<>"") then
                  set auxo = objS.colCompetencia.objByKey( iif(objS.TipoSeccionK = EDK_TIPOCOT, objR.IdCompetencia, objR.IdSeccionCompetencia) )
                  if objS.lblRubroAdicional<>"" then
                    response.write "<div class=""tsmall"">" & replace(auxo.DescripcionAdicional,vbCRLF,"<br />") & "</div>"
                  end if
                  if objS.lblEjemplo<>"" then
                    response.write "<div class=""tsmall"">" & replace(auxo.DescripcionCompetencia,vbCRLF,"<br />") & "</div>"
                  end if
                end if
                if idx >= 0 then %>
              <input type="hidden" id="rubroDes<%=idx%>" name="rubroDes<%=idx%>" value="<%=serverHTMLencode(objR.Rubro)%>"> <%
                  if objS.TipoSeccionK = EDK_TIPOCOT then %>
              <input type="hidden" id="rubroIdC<%=idx%>" name="rubroIdC<%=idx%>" value="<%=objR.IdCompetencia%>"> <%
                  end if
                end if
              end if
              if idx >= 0 AND objr.history.count > 0 AND not edk_hideHistory then %>
              <div class="noshowimp">[<a href="#" onClick="return showHistory(<%=objR.IdRubro%>);"><%=lblFRS_Historial%></a>]</div><%
              end if %>
            </td> <%
          IF objS.TipoSeccionK = EDK_TIPOLOG THEN
            pesoCell objS, objR, idx, allowEdit
            if lblEDK_logroMinimo<>"" then logroCell idx, "rubroLmin", objR.logroMinimo, allowEdit, false, false, 0
            logroCell idx, "rubroLesp", objR.logroEsperado, allowEdit, false, objS.usePolaridad, objR.logroPolaridad
            if lblEDK_logroMaximo<>"" then logroCell idx, "rubroLmax", objR.logroMaximo, allowEdit, false, false, 0
            '-- Logro
            logroCell idx, "rubroLog", objR.rev.Logro, allowEvaluate AND (stage<>0), (stage=0), false, 0
            '-- Resultado crudo
            if NOT (edk_hideresrubro OR hideSelfResults) then %>
            <td id="rubroLresDiv<%=idx%>" style="text-align:right;">
            <%=iif( (stage=0), "-", formatNumber(objR.rev.ResultadoLogro,numDec) ) & lblEDK_PercentSign%>
            </td> <%
            end if
            if idx >=0 then %>
            <input type="hidden" id="rubroLres<%=idx%>" name="rubroLres<%=idx%>" value="<%=objR.rev.ResultadoLogro%>"> <%
            end if
            '-- Resultado ponderado
            resultadoCell objS, objR, idx, false, (stage=0)
          ELSE
            '-- Informacion de establecimiento de objetivos o revisiiones anteriore
            tdPaintPastInfo stage, objS, objR
            '-- Informacion de acuerdo a la etapa
            if stage = 0 then 'Establecimiento de objetivos
              if ((objS.TipoSeccionK = EDK_TIPOCOM) OR (objS.TipoSeccionK = EDK_TIPOCOT)) and objS.lblRubroAdicional<>"" then
                set auxo = objS.colCompetencia.objByKey( iif(objS.TipoSeccionK = EDK_TIPOCOT, objR.IdCompetencia, objR.IdSeccionCompetencia) )
                if auxo is nothing then
                  auxtxt = ""
                else
                  auxtxt = replace(auxo.DescripcionAdicional,vbCRLF,"<br />")
                end if %>
            <td id="rubroAddDiv<%=idx%>" class="tsmall"><%=auxtxt%></td> <%
              end if
              if ((objS.TipoSeccionK = EDK_TIPOCOM) OR (objS.TipoSeccionK = EDK_TIPOCOT)) and objS.lblEjemplo<>"" then
                set auxo = objS.colCompetencia.objByKey( iif(objS.TipoSeccionK = EDK_TIPOCOT, objR.IdCompetencia, objR.IdSeccionCompetencia) )
                if auxo is nothing then
                  auxtxt = ""
                else
                  auxtxt = replace(auxo.DescripcionCompetencia,vbCRLF,"<br />")
                end if %>
            <td id="rubroEjeDiv<%=idx%>" class="tsmall"><%=auxtxt%></td> <%
              end if
              if objS.lblEsperado<>"" then %>
            <td> <%
                if allowEdit AND NOT sFixed then %>
              <textarea id="rubroEsp<%=idx%>" name="rubroEsp<%=idx%>" rows="<%=textrows%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.Esperado)%></textarea> <%
                else
                  response.write replace( serverHTMLencode(objR.Esperado), vbCRLF, "<br />" )
                  if idx >= 0 then %>
              <input type="hidden" id="rubroEsp<%=idx%>" name="rubroEsp<%=idx%>" value="<%=serverHTMLencode(objR.Esperado)%>"> <%
                  end if
                end if %>
            </td> <%
              end if
              if (objS.TipoSeccionK <> EDK_TIPOCOM) AND (objS.TipoSeccionK <> EDK_TIPOCOT) and objS.lblFechaCompromiso<>"" then %>
            <td class="tsmall" align="center" nowrap> <%
                if allowEdit then %>
              <input id="rubroFC<%=idx%>" name="rubroFC<%=idx%>" value="<%=formatDateDMAnull(objR.FechaCompromiso)%>" class="whiteblur" style="width:80px;" readonly="true">
              <A href="#" onclick="khorCalendar('rubroFC<%=idx%>'); return false;"><IMG src="khorImg/ico_calendar.gif" align="middle" border="0" height="16" width="16" alt="<%=lblFRS_ClickParaSeleccionarFecha%>"></A> <%
                else
                  response.write formatDateDisp(objR.FechaCompromiso,false)
                end if %>
            </td> <%
              end if
              if objS.lblModoCalculo<>"" then
                if objS.TipoSeccionK = EDK_TIPOKPI then %>
            <td> <%
                  if allowEdit then %>
              <textarea id="rubroMod<%=idx%>" name="rubroMod<%=idx%>" rows="<%=textrows%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.ModoCalculo)%></textarea> <%
                  else
                    response.write replace( serverHTMLencode(objR.ModoCalculo), vbCRLF, "<br />" )
                    if idx >= 0 then %>
              <input type="hidden" id="rubroMod<%=idx%>" name="rubroMod<%=idx%>" value="<%=serverHTMLencode(objR.ModoCalculo)%>"> <%
                    end if
                  end if %>
            </td> <%
                elseif idx >= 0 then %>
            <input type="hidden" id="rubroMod<%=idx%>" name="rubroMod<%=idx%>" value=""> <%
                end if
              end if
              pesoCell objS, objR, idx, allowEdit
            else  '-- Revisiones
              if not stageEval then
                pesoCell objS, objR, idx, allowEdit
              end if
              if objS.conPeso then
                if hideSelfResults OR (objS.lblRevision="" AND NOT stageEval) OR objS.evalSinObservaciones then
                  if idx >= 0 then %>
              <input type="hidden" id="rubroObs<%=idx%>" name="rubroObs<%=idx%>" value="<%=serverHTMLencode(objR.rev.Observaciones)%>"> <%
                  end if
                elseif not sFixed then
                  dim showObs : showObs = true
                  if habAutoRevision then %>
            <td> <%
                    if allowEvaluate AND (personaSesion() = persona.idpersona) then %>
              <textarea id="rubroObA<%=idx%>" name="rubroObA<%=idx%>" rows="<%=textrows%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision<>idxRevision)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.rev.ObsEvaluado)%></textarea> <%
                      showObs = false
                    else
                      showObs = (personaSesion() <> persona.idpersona)
                      response.write replace( serverHTMLencode(objR.rev.ObsEvaluado), vbCRLF, "<br />" )
                      if idx >= 0 then %>
              <input type="hidden" id="rubroObA<%=idx%>" name="rubroObA<%=idx%>" value="<%=serverHTMLencode(objR.rev.ObsEvaluado)%>"> <%
                      end if
                    end if %>
            </td> <%
                  end if
                  if showObs then
                    dim revprefix : revprefix = ""
                    dim rev : set rev = new EDK_RubroRevision
                    dim rsrev
                    if edk_useObsAdicional AND NOT habAutoRevision then %>
            <td> <%
                      if edk_showPastRevisions AND objR.IdRubro>0 AND idxRevision>1 AND idx>=0 then
                        set rsrev = getrs(conn,"SELECT * FROM EDK_RubroRevision WHERE IdRubro=" & objR.IdRubro & " AND numRevision < " & idxRevision & " ORDER BY numRevision")
                        while not rsrev.eof
                          rev.getFromRS rsrev %>
              <div class="tsmall"><%="<b>" & lblFRS_Revision & " " & rev.numRevision & ":</b> " & replace( serverHTMLencode(rev.ObsEvaluado), vbCRLF, "<br />" )%></div> <%
                          revprefix = "*"
                          rsrev.movenext
                        wend
                        rsrev.close
                        set rsrev = nothing
                        if revprefix<>"" then response.write "<b>" & iif( idxRevision<colRevision.count, lblFRS_Revision & " " & idxRevision, lblED_RevisionFinal ) & ":</b> "
                      end if
                      if allowEvaluate then %>
              <textarea id="rubroObA<%=idx%>" name="rubroObA<%=idx%>" rows="<%=textrows%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision<>idxRevision)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.rev.ObsEvaluado)%></textarea> <%
                      else
                        response.write replace( serverHTMLencode(objR.rev.ObsEvaluado), vbCRLF, "<br />" )
                        if idx >= 0 then %>
              <input type="hidden" id="rubroObA<%=idx%>" name="rubroObA<%=idx%>" value="<%=serverHTMLencode(objR.rev.ObsEvaluado)%>"> <%
                        end if
                      end if %>
            </td> <%
                    end if %>
            <td> <%
                    if edk_showPastRevisions AND objR.IdRubro>0 AND idxRevision>1 AND idx>=0 then
                      set rsrev = getrs(conn,"SELECT * FROM EDK_RubroRevision WHERE IdRubro=" & objR.IdRubro & " AND numRevision < " & idxRevision & " ORDER BY numRevision")
                      while not rsrev.eof
                        rev.getFromRS rsrev %>
              <div class="tsmall"><%="<b>" & lblFRS_Revision & " " & rev.numRevision & ":</b> " & replace( serverHTMLencode(rev.Observaciones), vbCRLF, "<br />" )%></div> <%
                        revprefix = "*"
                        rsrev.movenext
                      wend
                      rsrev.close
                      set rsrev = nothing
                      if revprefix<>"" then response.write "<b>" & iif( idxRevision<colRevision.count, lblFRS_Revision & " " & idxRevision, lblED_RevisionFinal ) & ":</b> "
                    end if
                    if allowEvaluate then %>
              <textarea id="rubroObs<%=idx%>" name="rubroObs<%=idx%>" rows="<%=textrows%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision<>idxRevision)%>)" style="width:100%" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.rev.Observaciones)%></textarea> <%
                    else
                      response.write replace( serverHTMLencode(objR.rev.Observaciones), vbCRLF, "<br />" )
                      if idx >= 0 then %>
              <input type="hidden" id="rubroObs<%=idx%>" name="rubroObs<%=idx%>" value="<%=serverHTMLencode(objR.rev.Observaciones)%>"> <%
                      end if
                    end if %>
            </td> <%
                    set rev = nothing
                  elseif idx >= 0 then
                    if edk_useObsAdicional AND NOT habAutoRevision then %>
              <input type="hidden" id="rubroObA<%=idx%>" name="rubroObA<%=idx%>" value="<%=serverHTMLencode(objR.rev.ObsEvaluado)%>"> <%
                    end if %>
              <input type="hidden" id="rubroObs<%=idx%>" name="rubroObs<%=idx%>" value="<%=serverHTMLencode(objR.rev.Observaciones)%>"> <%
                  end if
                end if
              end if
              if stageEval then
                if objS.conPeso then
                  if hideSelfResults then
                    if (idx >=0) AND not allowEvaluate then %>
                <input type="hidden" id="rubroCal<%=idx%>" name="rubroCal<%=idx%>" value="<%=objR.rev.Calificacion%>"> <%
                    end if
                  elseif objS.allowEditResultado=0 then
                    if escalaK.desglosada then
                      dim i
                      for i=1 to escalaK.colDetalle.count
                        set auxo = escalaK.colDetalle.obj(i)
                        if allowEvaluate then
                          response.write vbCRLF & "<td align=""center"" onClick=""cRadioSet(" & idx & "," & i & ");"" title=""" & strAdd(auxo.calTitulo, " - ", auxo.CalDescripcion) & """>" & _
                                         "<input type=""radio"" id=""radioCal" & idx & """ name=""radioCal" & idx & """ value=""" & i & """ onClick=""cRadioSet(" & idx & "," & i & ")""" & checkedIf(objR.rev.Calificacion=i) & ">" & _
                                         "</td>"
                        else
                          response.write vbCRLF & "<td style=""text-align:center;font-size:120%;font-weight:bold;"">" & iif( objR.rev.Calificacion = i, "X", "&nbsp;" ) & "</td>"
                        end if
                      next
                      if allowEvaluate then
                        response.write "<input type=""hidden"" id=""rubroCal" & idx & """ name=""rubroCal" & idx & """ value=""" & objR.rev.Calificacion & """>"
                      end if
                    else %>
            <td<%=iif(escalaK.MaxResultado <> 0," style=""text-align:right;""","")%>> <%
                      if allowEvaluate AND NOT sFixed then
                        if escalaK.MaxResultado <> 0 then %>
              <input type="text" id="rubroCal<%=idx%>" name="rubroCal<%=idx%>" value="<%=iif(isnull(objR.rev.Fecha),"",objR.rev.Calificacion)%>" onChange="cChanged(this,<%=idx%>)" maxlength="3" style="width:40px; text-align:right;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onFocus="inFocus(this);" onBlur="inBlur(this);"> <%
                        else %>
              <select id="rubroCal<%=idx%>" name="rubroCal<%=idx%>" onChange="cChanged(this,<%=idx%>);" style="font-size:10px;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"> <% 
                '-- Variable sucia en formato menufijo
                'edk_optionEscalaCal = "0:No cumple|5:Cumple"
                if (edk_optionEscalaCal&"") <> "" then
                  optionFromMenuFijo edk_optionEscalaCal, objR.rev.Calificacion
                else
                  response.write "<option value="""">-</option>"
                  response.write escalaK.strOptionFrom(objR.rev.Calificacion)
                end if %>
              </select> <%
                        end if
                      else
                        if escalaK.MaxResultado <> 0 then
                          response.write objR.rev.Calificacion
                        else
                          response.write escalaK.strCalificacion(objR.rev.Calificacion)
                        end if
                        if (idx>=0) AND sFixed then %>
                <input type="hidden" id="rubroCal<%=idx%>" name="rubroCal<%=idx%>" value="<%=objR.rev.Calificacion%>"> <%
                        end if
                      end if %>
            </td> <%
                    end if
                    if (idx >=0) AND not allowEvaluate then %>
                <input type="hidden" id="rubroCal<%=idx%>" name="rubroCal<%=idx%>" value="<%=objR.rev.Calificacion%>"> <%
                    end if
                  end if
                  pesoCell objS, objR, idx, allowEdit
                  resultadoCell objS, objR, idx, allowEvaluate, false
                else
                  pesoCell objS, objR, idx, false
                end if
              end if
            end if
          END IF %>
          </tr> <%
          IF objS.TipoSeccionK = EDK_TIPOLOG AND objS.lblModoCalculo<>"" THEN
            if objS.lblEsperado<>"" then %>
          <tr class="<%=estilo%>">
            <td colspan="<%=numColumnas%>" class="tsmall">
              <div style="font-weight:bold;margin-top:-3;margin-bottom:-3;"><%=objS.lblEsperado%>:</div> <%
                if allowEdit then %>
              <textarea id="rubroEsp<%=idx%>" name="rubroEsp<%=idx%>" rows="<%=(textrows-1)%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="width:100%;font-size:8pt;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.Esperado)%></textarea> <%
                else
                  response.write replace( serverHTMLencode(objR.Esperado), vbCRLF, "<br />" )
                  if idx >= 0 then %>
              <input type="hidden" id="rubroEsp<%=idx%>" name="rubroEsp<%=idx%>" value="<%=serverHTMLencode(objR.Esperado)%>"> <%
                  end if
                end if %>
            </td>
          </tr><%
            end if %>
          <tr class="<%=estilo%>">
            <td colspan="<%=numColumnas%>" class="tsmall">
              <div style="font-weight:bold;margin-top:-3;margin-bottom:-3;"><%=objS.lblModoCalculo%>:</div> <%
                  if allowEdit then %>
              <textarea id="rubroMod<%=idx%>" name="rubroMod<%=idx%>" rows="<%=(textrows-1)%>" maxlength="4000" onChange="tChanged(this,<%=idx%>,<%=bool2num(calRevision>0)%>)" style="width:100%;font-size:8pt;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%=serverHTMlencode(objR.ModoCalculo)%></textarea> <%
                  else
                    response.write replace( serverHTMLencode(objR.ModoCalculo), vbCRLF, "<br />" )
                    if idx >= 0 then %>
              <input type="hidden" id="rubroMod<%=idx%>" name="rubroMod<%=idx%>" value="<%=serverHTMLencode(objR.ModoCalculo)%>"> <%
                    end if
                  end if %>
            </td>
          </tr> <%
          END IF
  end sub
  
  '----------------------------------------
  
  function seccionTitleTooltip(lblName,ids,stage)
    'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
    dim stageEval : stageEval = (stage = -1) OR (stage > 0 AND edk_parciales_cuantitativas)
    dim objName : objName = lblName & "_" & ids & "_" & iif(stageEval,"E",stage)
    dim retval : retval = ""
    if eval(objName)<>"" then
      retval = " id=""" & objName & """ onmouseover=""edkToolTipGeneric('" & objName & "','" & strJS(eval(objName)) & "');"" onMouseOut=""hideToolTip();"""
    end if
    seccionTitleTooltip = retval
  end function

  function seccionTableTitles(stage,objS,raw,wdate)
    dim colTitles : colTitles = ""
    dim titlepast : titlepast = ""
    dim colspan, auxs, colw, xtraw, totw
    dim sFixed : sFixed = ((objS.TipoSeccionK = EDK_TIPOORG) OR (objS.TipoSeccionK = EDK_TIPOEST) OR objS.evalSinObservaciones)
    dim stageEval : stageEval = (stage = -1)
    IF objS.TipoSeccionK = EDK_TIPOLOG THEN
      columnaPeso = 2
      colSpan = 8
      colw = 32
      if lblEDK_logroMinimo="" then
        colSpan = colSpan - 1
        colw = colw + 10
      end if
      if lblEDK_logroMaximo="" then
        colSpan = colSpan - 1
        colw = colw + 10
      end if
      if hideSelfResults then
        colSpan = colSpan - 1
        colw = colw + 10
      end if
      if edk_hideresrubro OR hideSelfResults then
        colSpan = colSpan - 2
        colw = colw + 20
      end if
      coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRubro & "</td>"
      coltitles = coltitles & "<td width=""8%"" class=""tsmall"">" & lblEDK_Peso & "</td>"
      if lblEDK_logroMinimo<>"" then coltitles = coltitles & "<td width=""10%"">" & lblEDK_logroMinimo & "</td>"
      coltitles = coltitles & "<td width=""10%"">" & lblEDK_logroEsperado & "</td>"
      if lblEDK_logroMaximo<>"" then coltitles = coltitles & "<td width=""10%"">" & lblEDK_logroMaximo & "</td>"
      if not hideSelfResults then
        coltitles = coltitles & "<td width=""10%"">" & lblEDK_logro & "</td>"
      end if
      if not (edk_hideresrubro OR hideSelfResults) then                       
        coltitles = coltitles & "<td width=""10%"">" & lblEDK_logroResultado & "</td>" & _
                                "<td width=""10%"">" & lblEDK_ResultadoPonderado & "</td>"
      end if
    ELSE
      'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
      stageEval = stageEval OR (stage > 0 AND edk_parciales_cuantitativas)
      if stageEval then '-- Revision cuantitativa
        colSpan = 4
        titlepast = titlePastInfo(stage,objS)
        if titlepast <> "" then colSpan = colSpan+1
        if objS.conPeso then
          xtraw = 26
          if hideSelfResults then
            colSpan = colSpan - 1
            xtraw = xtraw - 10
          else
            if NOT sFixed then
              colSpan = colSpan + 1
              if (habAutoRevision AND NOT (personaSesion() = persona.idpersona)) OR (edk_useObsAdicional AND NOT habAutoRevision) then
                colSpan = colSpan + 1
                xtraw = xtraw + 10
              end if
            end if
            if objS.allowEditResultado then
              colSpan = colSpan - 1
              xtraw = xtraw - 10
            elseif escalaK.desglosada then
              colSpan = colSpan - 1 + escalaK.colDetalle.count
              xtraw = xtraw - 10 + (5 * escalaK.colDetalle.count)
            end if
          end if
          if edk_hidepesos then
            colSpan = colSpan - 1
            xtraw = xtraw - 8
          end if
          if edk_hideresrubro OR hideSelfResults then
            colSpan = colSpan - 1
            xtraw = xtraw - 8
          end if
        else
          xtraw = 0
        end if
        colw = int((100 - xtraw) / (iif(titlepast <> "",3,2)-iif(hideSelfResults,1,0)))
        coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRubro & "</td>"
        if titlepast <> "" then coltitles = coltitles & "<td width=""" & colw & "%"">" & titlepast & "</td>"
        if objS.conPeso then
          if not hideSelfResults then
            if not sFixed then
              if habAutoRevision then
                coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_RevisionFinalObsAuto & "</td>"
                if personaSesion() <> persona.idpersona then
                  coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_RevisionFinalObsJefe & "</td>"
                end if
              else
                if edk_useObsAdicional then
                  coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_ObservacionAdicional & "</td>"
                end if
                coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_Observaciones & "</td>"
              end if
            end if
            if objS.allowEditResultado=0 then
              if escalaK.desglosada then
                auxs = iif( khorPage_isPDF(), "", " onMouseOver=""escalaShowToolTip(" & objS.IdSeccionK & ",%%0);"" onMouseOut=""hideToolTip();""" )
                for i=1 to escalaK.colDetalle.count
                  set auxo = escalaK.colDetalle.obj(i)
                  coltitles = coltitles & "<td width=""5%"" align=""center"" " & strLang(auxs,i) & "><b>" & auxo.CalTitulo & "</b>" & iif(auxo.CalDescripcion<>"", "<div class=""tsmall"">" & auxo.CalDescripcion & "</div>", "" ) & "</td>"
                next
              else
                coltitles = coltitles & "<td width=""10%"">" & lblEDK_Calificacion & iif( escalaK.MaxResultado > 0, "<br />(" & escalaK.MinCalificacion & "-" & escalaK.maxEscalaCal() & ")", "" ) & "</td>"
              end if
            end if
          end if
          if not edk_hidepesos then
            coltitles = coltitles & "<td width=""8%"" class=""tsmall"">" & lblEDK_Peso & "</td>"
          end if
          if not (edk_hideresrubro OR hideSelfResults) then
            coltitles = coltitles & "<td width=""8%"" class=""tsmall"">" & lblEDK_Resultado & "</td>"
          end if
        end if
      elseif stage = 0 then 'Establecimiento de objetivos
        select case objS.TipoSeccionK
          case EDK_TIPOCOM, EDK_TIPOCOT
            totw = iif( edk_hidepesos, 100, 90 )
            colspan = 0.5
            if objS.lblRubroAdicional<>"" then colspan = colspan+1
            if objS.lblEjemplo<>"" then colspan = colspan+1
            if objS.lblEsperado<>"" then colspan = colspan+1
            colw = floor( totw / colspan)
            colspan = colspan + 0.5
            coltitles = coltitles & "<td width=""" & (totw-(colw * (colspan-1))) & "%"">" & objS.lblRubro & "</td>"
            if objS.lblRubroAdicional<>"" then  coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRubroAdicional & "</td>"
            if objS.lblEjemplo<>"" then         coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblEjemplo & "</td>"
            if objS.lblEsperado<>"" then        coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblEsperado & "</td>"
            if objS.conPeso AND not edk_hidepesos then
              coltitles = coltitles & "<td width=""10%"">" & lblEDK_Peso & "</td>"
              colspan = colspan+1
            end if
          case EDK_TIPOKPI
            totw = iif( edk_hidepesos, 100, 90 )
            xtraw = 0
            colspan = 1
            if objS.lblEsperado<>"" then colspan = colspan+1
            if objS.lblFechaCompromiso<>"" then colspan = colspan+1
            if objS.lblModoCalculo<>"" then colspan = colspan+1
            colw = floor( totw / colspan )
            if objS.lblFechaCompromiso<>"" then xtraw = floor((colw/2) / (colspan-1))
            coltitles = coltitles & "<td width=""" & (totw-(colw * (colspan-1))+xtraw) & "%""" & seccionTitleTooltip("lblRubro",objS.IdSeccionK,stage) & ">" & objS.lblRubro & "</td>"
            if objS.lblEsperado<>"" then        coltitles = coltitles & "<td width=""" & (colw+xtraw) & "%""" & seccionTitleTooltip("lblEsperado",objS.IdSeccionK,stage) & ">" & objS.lblEsperado & "</td>"
            if objS.lblFechaCompromiso<>"" then coltitles = coltitles & "<td width=""" & (colw/2) & "%""" & seccionTitleTooltip("lblFechaCompromiso",objS.IdSeccionK,stage) & ">" & objS.lblFechaCompromiso & "</td>"
            if objS.lblModoCalculo<>"" then     coltitles = coltitles & "<td width=""" & (colw+xtraw) & "%""" & seccionTitleTooltip("lblModoCalculo",objS.IdSeccionK,stage) & ">" & objS.lblModoCalculo & "</td>"
            if not edk_hidepesos then
              coltitles = coltitles & "<td width=""10%""" & seccionTitleTooltip("lblPeso",objS.IdSeccionK,stage) & ">" & lblEDK_Peso & "</td>"
              colspan = colspan+1
            end if
          case EDK_TIPOACT
            totw = iif( edk_hidepesos, 100, 90 )
            xtraw = 0
            colspan = 1
            if objS.lblFechaCompromiso<>"" then colspan = colspan+1
            if objS.lblEsperado<>"" then        colspan = colspan+1
            colw = floor( totw / colspan )
            if objS.lblFechaCompromiso<>"" then xtraw = floor((colw/2) / (colspan-1))
            coltitles = coltitles & "<td width=""" & (totw-(colw * (colspan-1))+xtraw) & "%"">" & objS.lblRubro & "</td>"
            if objS.lblEsperado<>"" then  coltitles = coltitles & "<td width=""" & (colw+xtraw) & "%"">" & objS.lblEsperado & "</td>"
            if objS.lblFechaCompromiso<>"" then coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblFechaCompromiso & "</td>"
            if objS.conPeso AND not edk_hidepesos then
              coltitles = coltitles & "<td width=""10%"">" & lblEDK_Peso & "</td>"
              colspan = colspan+1
            end if
          case EDK_TIPOEST, EDK_TIPOORG, EDK_TIPOOBJ
            totw = iif( edk_hidepesos, 100, 90 )
            xtraw = 0
            colspan = 1
            if objS.lblEsperado<>"" then        colspan = colspan+1
            colw = floor( totw / colspan )
            coltitles = coltitles & "<td width=""" & (totw-(colw * (colspan-1))+xtraw) & "%"">" & objS.lblRubro & "</td>"
            if objS.lblEsperado<>"" then  coltitles = coltitles & "<td width=""" & (colw+xtraw) & "%"">" & objS.lblEsperado & "</td>"
            if objS.conPeso AND not edk_hidepesos then
              coltitles = coltitles & "<td width=""10%"">" & lblEDK_Peso & "</td>"
              colspan = colspan+1
            end if
        end select
      else 'Revisiones tradicionales (no cuantitativas)
        colspan = 2
        titlepast = titlePastInfo(stage,objS)
        if titlepast <> "" then
          colw = 30
          colSpan = colSpan+1
        else
          colw = 45
        end if
        if (objS.lblRevision <> "") AND NOT sFixed then
          colSpan = colSpan + 1
          if (habAutoRevision AND NOT (personaSesion() = persona.idpersona)) OR (edk_useObsAdicional AND NOT habAutoRevision) then
            colSpan = colSpan + 1
            colw = colw - 15
          end if
        end if
        if edk_hidepesos then
          colSpan = colSpan - 1
          colw = colw + floor(10/colSpan)
        end if
        coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRubro & "</td>"
        if titlepast <> "" then coltitles = coltitles & "<td width=""" & colw & "%"">" & titlepast & "</td>"
        if objS.conPeso then
          if not edk_hidepesos then
            coltitles = coltitles & "<td width=""10%"" class=""tsmall"">" & lblEDK_Peso & "</td>"
          end if
          if (objS.lblRevision <> "") AND NOT sFixed then
            if habAutoRevision then
              coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_RevisionParcialObsAuto & "</td>"
              if not (personaSesion() = persona.idpersona) then
                coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRevision & "</td>"
              end if
            else
              if edk_useObsAdicional then
                coltitles = coltitles & "<td width=""" & colw & "%"">" & lblEDK_ObservacionAdicional & "</td>"
              end if
              coltitles = coltitles & "<td width=""" & colw & "%"">" & objS.lblRevision & "</td>"
            end if
          end if
        end if
      end if
    END IF
    if objS.conTipoRubro then
      coltitles = "<td width=""10%"">" & objS.lblTipoRubro & "</td>" & coltitles
      colspan = colspan + 1
      columnaPeso = columnaPeso + 1
    end if
    if objS.conObjetivoEstrategico then
      coltitles = "<td width=""10%"">" & objS.lblObjetivoEstrategico & "</td>" & coltitles
      colspan = colspan + 1
      columnaPeso = columnaPeso + 1
    end if
    if wdate then
      coltitles = "<td width=""10%"">" & lblFRS_Fecha & "</td>" & coltitles
      colspan = colspan + 1
      columnaPeso = columnaPeso + 1
    end if
    if objS.conPeso AND not raw then
      dim directions : directions = ""
      if not sFixed then
        select case stage
          case -1 'Revision final
            directions = objS.lblInstruccion3
          case 0  'Establecimiento de objetivos
            directions = strAdd( objS.lblInstruccion1, " ", seccionLeyendaRubros(objS) )
            directions = strAdd( directions, " ", seccionLeyendaPesos(objS) )
          case else 'Revisiones
            directions = objS.lblInstruccion2
        end select
      end if %>
          <tr class="celdaTit">
            <td colspan="<%=colspan%>" align="center"><%=objS.SeccionK%></td>
          </tr> <%
          if directions<>"" then %>
          <tr>
            <td class="tsmall" colspan="<%=colspan%>"><%=directions%></td>
          </tr> <%
          end if
    end if %>
          <tr class="celdaTit"><%=coltitles%></tr> <%
    numColumnas = colSpan
    seccionTableTitles = colSpan
  end function
  
  '----------------------------------------

  private sub tdPaintPastInfo(stage,objS,objR)
    if titlePastInfo(stage,objS) <> "" then
      dim sep : sep = " <hr width=""50%""> "
      dim retval : retval = replace( serverHTMLencode(objR.Esperado), vbCRLF, "<br />" )
      if (objS.TipoSeccionK <> EDK_TIPOCOM) AND (objS.TipoSeccionK <> EDK_TIPOCOT) AND NOT isnull(objR.FechaCompromiso) then
        retval = strAdd( formatDateDisp(objR.FechaCompromiso,false), ":", retval )
      end if
      select case objS.TipoSeccionK
        case EDK_TIPOCOM, EDK_TIPOCOT
          if not objS.showInitialData then
            set auxo = objS.colCompetencia.objByKey( iif(objS.TipoSeccionK = EDK_TIPOCOT, objR.IdCompetencia, objR.IdSeccionCompetencia) )
            retval = strAdd( replace(auxo.DescripcionAdicional,vbCRLF,"<br />"), sep, retval )
            retval = strAdd( replace(auxo.DescripcionCompetencia,vbCRLF,"<br />"), sep, retval )
            set auxo = nothing
          end if
        case EDK_TIPOKPI
          retval = strAdd( retval, sep, replace( serverHTMLencode(objR.ModoCalculo), vbCRLF, "<br />" ) )
      end select
      response.write "<td class=""tsmall"">" & retval & "</td>"
    end if
  end sub

  private function titlePastInfo(stage,objS)
    dim sep : sep = " / "
    dim retval : retval = ""
    if (stage <> 0) AND ((khorConfigValue(541,true) <> 0) OR objS.showInitialData OR NOT objS.conPeso) AND (objS.TipoSeccionK <> EDK_TIPOLOG) then
      retval = objS.lblEsperado
      if (objS.TipoSeccionK <> EDK_TIPOCOM) AND (objS.TipoSeccionK <> EDK_TIPOCOT) then
        retval = strAdd( objS.lblFechaCompromiso, sep, retval )
      end if
      select case objS.TipoSeccionK
        case EDK_TIPOCOM, EDK_TIPOCOT
          if not objS.showInitialData then
            retval = strAdd( objS.lblRubroAdicional, sep, retval )
            retval = strAdd( objS.lblEjemplo, sep, retval )
          end if
        case EDK_TIPOKPI
          retval = strAdd( retval, sep, objS.lblModoCalculo )
      end select
    end if
    titlePastInfo = retval
  end function
  
end class  

'========================================
%>