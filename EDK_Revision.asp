<!--#include file="./khorClass.asp"-->
<!--#include file="./ed_class.asp"-->
<!--#include file="./edk_auxiliar.asp"-->
<!--#include file="./plazaClass.asp"-->
<%Modificacion al producto
  thispage = "EDK_Revision.asp"
  thispageid = "FRMedkRevision"
  '-- Validacion de acceso: modulo habilitado y sesion activa
  validaEntrada khorEDKhabilitada(), "", thispageid
  dim mov : mov = ucase(reqplain("mov"))
  if mov<>"PDF" then
    childwin = (reqplain("childwin")=thispageid)
    if getSesionDesc()="" then sesionFromRequest
  else
    sesionFromRequest
    childwin = true
  end if
  checaSesion ses_userid&","&ses_super&","&ses_adminid, "", ""

  '-- Inicializacion
  dim modo : modo = ""
  dim filtroPersonaSelect : filtroPersonaSelect = ""
  dim IdPersonaUser : IdPersonaUser = personaSesion()
  dim upperManager : upperManager = false
  dim IdPeriodo : IdPeriodo = reqn("IdPeriodo")
  dim IdPersona : IdPersona = reqn("IdPersona")
  '-- Procesa parametro de IdEvaluacion
  dim IdEvaluacion : IdEvaluacion = reqn("idevaluacion")
  sq = "SELECT IdEvaluado, IdPeriodo FROM vED_GrupoPersona WHERE IdEvaluacion=" & IdEvaluacion
  set rs = getrs(conn,sq)
  if not rs.eof then
    IdPersona = rsNum(rs,"IdEvaluado")
    IdPeriodo = rsNum(rs,"IdPeriodo")
  else
    IdEvaluacion = 0
  end if
  rs.close
  set rs = nothing
  dim vgpFiltro : vgpFiltro = "(IdEvaluacionK>0) AND " & periodoSQLrubro(tipoCR_EDK,"vED_GrupoPersona.IdPeriodo")
  '-- Carga los periodos validos en una coleccion
  dim colPeriodos : set colPeriodos = new frsCollection
  if IdPersonaUser > 0 then
    if edHidePeriodosInactivos() then vgpFiltro = vgpFiltro & " AND (PeriodoActivo<>0)"
    vgpFiltroPeriodo = vgpFiltro
    if IdEvaluacion=0 then
      vgpFiltroPeriodo = vgpFiltroPeriodo & " AND ((IdEvaluado=" & IdPersonaUser & ") OR (IdEvaluador=" & IdPersonaUser & "))"
    else
      vgpFiltroPeriodo = vgpFiltroPeriodo & " AND (IdEvaluado=" & IdPersona & ")"
    end if
  else
    vgpFiltroPeriodo = strAdd( vgpFiltro, " AND ", khorCondicionUsuario("IdEvaluado") )
  end if
  sq = "SELECT DISTINCT IdPeriodo, Periodo, FechaFin, FechaIni FROM vED_GrupoPersona WHERE " & vgpFiltroPeriodo & " ORDER BY FechaFin DESC, FechaIni DESC"
  colPeriodos.keyDescFromQuery "IdPeriodo", "Periodo", "", sq
  '-- Valida parametro de periodo
  if colPeriodos.keyIndex(IdPeriodo) = 0 then 
    if colPeriodos.count = 0 then
      IdPeriodo = 0
    else
      IdPeriodo = colPeriodos.obj(1).key
    end if
    IdEvaluacion = 0
  end if
  if IdPeriodo > 0 then vgpFiltro = vgpFiltro & " AND vED_GrupoPersona.IdPeriodo = " & IdPeriodo
  '-- Valida acceso de persona/administrador a la persona-parametro, y crea filtros de acceso a personas
  if IdPersonaUser > 0 then
    sq = "SELECT IdEvaluado FROM vED_GrupoPersona WHERE " & vgpFiltro & " AND ((IdEvaluado=" & IdPersonaUser & ") OR (IdEvaluador=" & IdPersonaUser & "))"
    dim listaPersonas : listaPersonas = getBDlist("IdEvaluado",sq,false)
    if (IdEvaluacion>0) AND (inCSV(listapersonas,IdPersona) < 0) then
      '-- Evaluacion explcita pero la persona no es evaluado directo del usuario en contexto, valida tramo de control
      ok = false
      dim auxPlaza : set auxPlaza = new clsPlaza
      '-- Obtiene linea de reporte (plazas jefe a cualquier nivel) de todas las plazas que ocupa el evaluado
      dim parents : parents = ""
      sq = "SELECT IdPlaza, IdPlazaReporta FROM Plaza WHERE Status=" & PLAZA_ACTIVA & " AND IdPersona=" & IdPersona
      set rs = getrs(conn,sq)
      while not rs.eof
        auxPlaza.IdPlaza = rsNum(rs,"IdPlaza")
        auxPlaza.IdPlazaReporta = rsNum(rs,"IdPlazaReporta")
        parents = addCSV( parents, auxPlaza.getParents() )
        rs.movenext
      wend
      rs.close
      set rs = nothing
      '-- Es valido si el usuario actual esta en alguna de sus plazas jefe
      if parents<>"" then
        if bdExists("SELECT IdPlaza FROM Plaza WHERE IdPersona=" & IdPersonaUser & " AND IdPlaza IN (" & parents & ")") then
          upperManager = true
          listaPersonas = IdPersona
          ok = true
        end if
      end if
      set auxPlaza = nothing
      if not ok then IdPersona = 0
    end if
    if listaPersonas="" then
      ok = false
    else
      if (inCSV(listapersonas,IdPersona) < 0) then
        IdPersona = iif(inCSV(listapersonas,IdPersonaUser) < 0, 0, IdPersonaUser)
        IdEvaluacion = 0
      end if
      filtroPersonaSelect = "IdPersonal IN (" & listaPersonas & ")"
      ok = true
    end if
    modo = iif(IdPersona = IdPersonaUser, "PER", "SUP" )
  elseif khorPermisoModulo(Modulo_EDK,khorModulosActivos()) then
    modo = "ADM"
    filtroPersonaSelect = "EXISTS(SELECT * FROM vED_GrupoPersona WHERE vED_GrupoPersona.IdEvaluado=Personal.IdPersonal AND " & vgpFiltro & ")"
    sq = "SELECT IdPersonal FROM Personal WHERE " & strAdd( khorCondicionUsuario("IdPersonal"), " AND ", filtroPersonaSelect )
    if IdPersona > 0 then
      if not bdExists(sq & " AND IdPersonal=" & IdPersona) then
        IdPersona = getBDnum( "IdPersonal", sq & " ORDER BY Nombre" )
      end if
    end if
    filtroEvaluadosAux = strAdd( khorCondicionUsuario("IdPersonal"), " AND ", "EXISTS(SELECT * FROM ED_GrupoPersona INNER JOIN ED_Grupo ON ED_GrupoPersona.IdGrupo = ED_Grupo.IdGrupo WHERE ED_GrupoPersona.IdEvaluado=Personal.IdPersonal AND (IdEvaluacionK>0) AND ED_GrupoPersona.IdEvaluador=%%0 AND ED_Grupo.IdPeriodo=%%1)" )
    ok = true
  else
    ok = false
  end if
  validaEntrada ok, "", thispageid

'================================================================================'

  '-- Inicializacion de objetos
  dim edk : set edk = new EDK_Interfaz
  dim allowView : allowView = (IdPersona > 0) AND (IdPeriodo > 0)
  dim allowEdit, allowEvaluate, adminEventos
  dim stage : stage = 0
  dim stageEval : stageEval = stage
  dim edkrep : edkrep = (reqn("edkrep") <> 0)
  if allowView then
    edk.getInfoPeriodo IdPeriodo
    edk.getInfoEvaluacion IdPersona
    if edk.regGP.IdEvaluacionK = 0 then
      allowView = false
      errmsg = strAdd( errmsg, "<br>", strLang( lblFRS_X_NoEsDatoValido, lblEDK_TipoDeEvaluacion ) )
    elseif edk.regEvaK.conObjetivoEstrategico and edk.regGP.IdUnidadAdministrativa_EDP=0 then
      allowView = false
      errmsg = strAdd( errmsg, "<br>", strLang( lblFRS_X_NoEsDatoValido, lblKHOR_UnidadAdministrativa ) )
    else
      verificaPermisos
    end if
  end if
  
  sub verificaPermisos()
    if reqplain("idxRevision") = "" or (reqn("idxRevision")=0 and reqplain("IdEvaluacion") <> "" and edkrep) then
      edk.idxRevision = edk.calRevision
    else
      edk.idxRevision = reqn("idxRevision")
    end if
    if modo="ADM" then
      allowEdit = supAllowEdit(edk.idxRevision)
      allowEvaluate = supAllowEvaluate(edk.idxRevision)
    else
      allowEvaluate = supAllowEvaluate(edk.idxRevision)
      allowEdit = supAllowEdit(edk.idxRevision) _
                  OR ((modo="PER") AND  autoAllowEdit(edk.idxRevision) AND NOT edk.regGP.edkLocked)
    end if
    stage = iif( edk.idxRevision=0, 0, iif( edk.idxRevision=edk.periodo.NumRevisiones, -1, 1 ) )
    'edk_parciales_cuantitativas = true  '-- Registra revisiones parciales con resultado cuantitativo como la final
    stageEval = (stage = -1) or (stage > 0 AND edk_parciales_cuantitativas)
    completaPermisos
  end sub
  
  sub completaPermisos()
    dim basicAllow : basicAllow = (mov <> "PDF") AND (NOT edk.regEvaK.Inactiva) AND (edk.regGP.statusPer = 0) AND NOT edkrep
    allowEvaluate = allowEvaluate AND basicAllow AND NOT upperManager
    allowEdit = allowEdit AND basicAllow AND NOT upperManager
    adminEventos = (modo="ADM") AND basicAllow
  end sub
  
  function supAllowEvaluate(numrev)
    supAllowEvaluate = (edk.calRevision = numrev) AND ((modo="ADM") OR ((modo="SUP") AND edk.enFechasRevision(numrev)))
  end function
  
  function supAllowEdit(numrev)
    'edk_disableEditObjInRev = true '-- Deshabilita edición de objetivos en revisiones.
    'edk_allowEditNulls = true  '-- Permite editar objetivos en revisión cuando no se han establecido.
    edk_allowEditNulls = (edk_allowEditNulls = true)
    supAllowEdit = (numrev=0) AND ((modo="ADM") OR (modo="SUP" AND ((edk.enFechasRevision(edk.calRevision) AND (edk.calRevision=0 OR NOT edk_disableEditObjInRev)) OR (isnull(edk.regGP.Fecha_EDK) AND (edk.enFechasRevision(0) OR (edk_allowEditNulls))))))
  end function
  
  function autoAllowEdit(numrev)
    autoAllowEdit = (edk.periodo.edpAutoObj AND (numrev=0) AND edk.enFechasRevision(numrev))
  end function
  
'================================================================================'

  if reloaded=0 then
    if mov="COMMIT" then
      edk.processRequest stage
      if modo<>"PER" AND numRev=0 AND autoAllowEdit(edk.idxRevision) then
        edk.regGP.edkLocked = (reqn("edkLocked") <> 0)
        conn.execute "UPDATE ED_GrupoPersona SET edkLocked=" & bool2num(edk.regGP.edkLocked) & " WHERE IdEvaluacion=" & edk.regGP.IdEvaluacion
      end if
    end if
  end if
  if mov<>"PDF" then mov = ""

'================================================================================'

  if allowView then
    edk.getRubrosK
  end if
    
  thispageurl = thispage & "?childwin=" & reqplain("childwin") & "&idpersona=" & edk.persona.IdPersona & "&idperiodo=" & edk.periodo.IdPeriodo
  titulo = lblED_EDK_titulo
  tit1 = ""
  tit2 = ""
  
'================================================================================'

  sub tablaRevisiones(estilo)
    dim i, auxrev, strDesc, strStat, conLink
    for i=0 to edk.colRevision.count
      if i=0 then
        set auxrev = new ED_PeriodoRevision
        auxrev.FechaIni = edk.periodo.edpObjFechaIni
        auxrev.FechaFin = edk.periodo.edpObjFechaFin
        strDesc = lblEDK_RegistroObjetivos
      else
        set auxrev = edk.colRevision.obj(i)
        strDesc = iif( i<edk.colRevision.count, lblFRS_Revision & " " & i, lblED_RevisionFinal )
      end if
      conLink = true
      if (auxrev.FechaIni > Date) then
        strStat = lblKHOR_Pendiente
        conLink = false
      elseif i=edk.idxRevision then
        strStat = "<span style=""color:#FF0000;"">" & iif(allowEdit OR allowEvaluate,lblFRS_Editando,lblFRS_Consultando) & "</span>"
        conLink = false
      elseif ((modo="ADM") OR (modo="SUP") AND edk.enFechasRevision(i)) then
        strStat = iif( isnull(edk.regGP.Fecha_EDK) OR (i > edk.regGP.Status_EDK), lblFRS_Registrar,  lblFRS_Modificar )
      else
        strStat = iif( (modo="ADM") OR (i=0 AND (supAllowEdit(i) OR autoAllowEdit(i))), lblFRS_Modificar, lblFRS_Consultar )
      end if
      conlink = conlink AND (mov<>"PDF")
      estilo = switchEstilo(estilo) %>
                <TR class="<%=estilo%>">
                  <TD class="tsmall"><%=strDesc%></TD>
                  <TD class="tsmall"><%=formatDateDisp(auxrev.FechaIni,false) & " - " & formatDateDisp(auxrev.FechaFin,false)%></TD>
                  <TD class="noshowimp"> <%
                    if conLink then %>
                    [<a href="#" onClick="return verRevision(<%=i%>);"><%=strStat%></a>] <%
                    elseif mov<>"PDF" then
                      response.write strStat
                    end if %>
                  </TD>
                </TR> <%
    next
  end sub
  
  sub evaluacionResumen(stage,adminEventos)
    edk.evaluacionResumen stage, adminEventos
  end sub

'================================================================================'
  layoutHeadStart khorAppName() & " - " & titulo
%>
<STYLE>
<!--
#edRK {
  border-collapse:collapse;
  width: 100%;
}
#edRK td {
  BORDER: 1px solid #AAAAAA;
  padding: 1px;
}
.edRKborder {
  BORDER: 1px solid #AAAAAA;
}
//-->
</STYLE>
<%IF mov<>"PDF" THEN%>
<% includeJS %>
<script language="JavaScript">
<!--
  var dirty = false;
  function setDirty() {
    dirty = true;
    setFlashObj('extraBtn_0');
    return false;
  }
  function verCompromisos() {
    if ( !dirty || confirm("<%=strJS(lblFRS_abandonarCambios)%>") ) {
      abreCompromisosDesarrollo(<%=edk.regGP.IdEvaluacion%>,null,null,null,'<%=thispageurl%>');
    }
    return false;
  }
  function verRevision(idr) {
    if ( !dirty || confirm("<%=strJS(lblFRS_abandonarCambios)%>") ) {
      sendval('','mov','','idxRevision',idr);
    }
    return false;
  }
  function myRegresar() {
    if ( getValor('dirty') !=0 && !confirm("<%=strJS(lblFRS_abandonarCambios)%>") ) return false;
    return true;
  }
  <%IF allowEdit OR allowEvaluate THEN%>
  function guardar() {
    if (!verificaRubros()) return false;
    sendval('','mov','commit');
  }
  <%END IF%>
  function seleccionaPersona() {
    abreSeleccion('PERSONA',false,'','<%=lblKHOR_Persona%>',null,null,null,'<%=setSelectionFilter("PERSONA",filtroPersonaSelect)%>');
  }
  function setSeleccion(tipo,lista) {
    if (tipo == 'PERSONA') {
      if (dirty && !confirm('<%=strJS(lblFRS_abandonarCambios)%>')) {
        var obj = MM_findObj('IdPersona');
        if ( obj.type == 'select-one' ) setValor('IdPersona',<%=edk.persona.IdPersona%>);
      } else {
        sendval('','mov','','IdPersona',lista);
      }
    }
    return false;
  }
  function setPeriodo() {
    if (dirty && !confirm('<%=strJS(lblFRS_abandonarCambios)%>')) {
      setValor('IdPeriodo',<%=IdPeriodo%>);
    } else {
      sendval('','mov','');
    }
  }
  <%IF pdf_enabled() THEN%>
  function myPrintPage() { <%
    pdfkey = initPDFurl( thispageid & "_" & IdPersona & "_" & IdPeriodo, _
                          pdf_URL() & thispage & "?mov=pdf&IdPersona=" & IdPersona & "&IdPeriodo=" & IdPeriodo & "&idxRevision=" & edk.idxRevision ) %>
    openPDFjob('<%=pdfkey%>');
  }
  <%END IF%>
  function edkToolTipGeneric(objId,msg) {
    var obj = document.getElementById(objId);
    if ( msg != "" && obj ) {
      resizeToolTip(obj.style.width,obj.style.height);
      showToolTip(msg,false);
    }
  }
  <%IF edk_usetooltips THEN%>
  <% edk_jsDeclare %>
  function edkToolTip(idr,idsecom) {
    edkToolTipGeneric( 'tdRubro'+idr, edkToolTipText(idsecom) );
  }
  <%END IF%>
//-->
</script>
<% edk.jsEdicionEDK stage %>
<script type="text/javascript" src="dropdown.js"></script>
<%END IF%>
<%
  layoutHeadEnd
'================================================================================'
  layoutStart titulo, tit1, tit2, errmsg, khorWinWidth(), iif(mov<>"PDF" AND edk.persona.idpersona=0 AND modo="ADM", " onload=""seleccionaPersona();""", "")
  defaultFormStart thispage, "onSubmit=""return false;""", true
  
        conEvaluados = false
        if modo="ADM" AND edk.persona.IdPersona>0 AND IdPeriodo>0 then
          sq = "SELECT IdPersonal, (Nombre + (CASE WHEN StatusPer<>0 THEN ' (" & lblFRS_Inactivo & ")' ELSE '' END)) AS NombreStatus" & _
              " FROM Personal WHERE " & strLang(filtroEvaluadosAux,edk.persona.IdPersona &"||"& IdPeriodo) & _
              " ORDER BY Nombre"
          dim colSub : set colSub = new frsCollection
          colSub.keyDescFromQuery "IdPersonal", "NombreStatus", "", sq %>
        <div id="mnuSub" name="mnuSub" class="fillGris" style="padding:5px;position:absolute;visibility:hidden;" nowrap> <%
          response.write "<b>" & lblED_PersonasAEvaluar & "</b>"
          conEvaluados = true
          if colSub.count > 0 then
            for i=1 to colSub.count
              set auxo = colSub.obj(i)
              response.write "<br/>[<a href=""#"" onClick=""return setSeleccion('PERSONA'," & auxo.key & ");"">" & auxo.desc & "</a>]"
            next
          else
            response.write "<br/>" & lblFRS_Ninguno
          end if %>
        </div> <%
          colSub.clean
          set colSub = nothing
        end if
      %>
        <table border="0" cellpadding="1" cellspacing="1" width="100%">
          <tr>
            <td valign="top" align="left">
              <table border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td><b><%=lblKHOR_Evaluado%>:</b></td> <%
                  IF modo="ADM" then %>
                  <input type="hidden" id="IdPersona" name="IdPersona" value="<%=edk.persona.IdPersona%>">
                  <td style="<%=iif(edk.persona.statusper=0,"","text-decoration:line-through")%>"> <%
                    if conEvaluados then %>
                    <div id="mnuSub0"><a href="#" onClick="return false;"><%=edk.persona.nombre%></a></div>
                    <script type="text/javascript">
                    at_attach("mnuSub0", "mnuSub", "hover", "y", "pointer");
                    </script> <%
                    else
                      response.write edk.persona.nombre
                    end if %>
                  </td>
                  <td> <%
                    IF mov<>"PDF" THEN %>
                     <INPUT type=button value="<%=lblFRS_Seleccionar%>" onclick="seleccionaPersona();" class=whitebtn onblur=inBlur(this); onmouseover=inOver(this); onfocus=inFocus(this); onmouseout=inOut(this);> <%
                    END IF %>
                  </td> <%
                  ELSE %>
                  <td colspan="2"> <%
                    IF mov<>"PDF" THEN %>
                    <select id="IdPersona" name="IdPersona" onChange="setSeleccion('PERSONA',this.value);" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
                    <%
                      sq = "SELECT IdPersonal, (Nombre + (CASE WHEN StatusPer<>0 THEN ' (" & lblFRS_Inactivo & ")' ELSE '' END)) AS NombreStatus FROM Personal  WHERE %%0 ORDER BY Nombre"
                      optionFromQuery "IdPersonal", "NombreStatus", edk.persona.IdPersona, strLang( sq, filtroPersonaSelect )
                    %>
                    </select> <%
                    ELSE
                      response.write edk.persona.nombre
                    END IF %>
                  </td> <%
                  END IF %>
                </tr> <%
                dim colData : set colData = new frsCollection
                colData.addKeyDesc "puesto",  lblKHOR_Perfil,               edk.regGP.puestoEvaluado
                colData.addKeyDesc "clavek",  lblFRS_Clave,                 descFromTable("Personal","IdPersonal","Clave",edk.regGP.IdEvaluado)
                if edk.regEvaK.conObjetivoEstrategico then
                  colData.addKeyDesc "uaedp", lblKHOR_UnidadAdministrativa, edk.regGP.UnidadAdministrativa_EDP
                end if
                colData.addKeyDesc "nevldr",  lblKHOR_Evaluador,            edk.regGP.nombreEvaluador
                colData.addKeyDesc "pevldr",  lblKHOR_Perfil,               edk.regGP.puestoEvaluador
                for i=1 to colData.count
                  set auxo = colData.obj(i)
                  if inCSV(edk_data2exclude,auxo.key) < 0 then %>
                <tr>
                  <td><b><%=auxo.desc%>:</b></td> <td colspan="2"><%=auxo.aux%></td>
                </tr> <%
                  end if
                next
                colData.clean
                set colData = nothing %>
              </table>
            </td>
            <td valign="top" align="right">
              <TABLE cellSpacing="1" cellPadding="1" border="0">
                <tr>
                  <td><b><%=lblED_PeriodoDeEvaluacion%>:</b></td>
                  <td colspan="2"> <%
                    IF mov<>"PDF" THEN %>
                    <select id="IdPeriodo" name="IdPeriodo" onChange="setPeriodo();" class=whiteblur onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
                      <% optionFromMenuFijo colPeriodos.menuFijoFromKeyDesc(), IdPeriodo %>
                    </select> <%
                    ELSE
                      response.write descFromCat( CAT_PERIODO, IdPeriodo )
                    END IF %>
                  </td>
                </tr> <%
                IF edk_showGrupo THEN %>
                <tr>
                  <td><b><%=lblKHOR_Grupo%>:</b></td>  <td colspan="2"><%=edk.regGP.Grupo%></td>
                </tr> <%
                END IF
                IF NOT edk_hideTipoEvaluacion THEN %>
                <tr>
                  <td><b><%=lblEDK_TipoDeEvaluacion%>:</b></td>  <td colspan="2"><%=edk.regGP.EvaluacionK%></td>
                </tr> <%
                END IF
                IF allowView THEN %>
                <INPUT type="hidden" id="idxRevision" name="idxRevision" value="<%=edk.idxRevision%>"> <%
                  tablaRevisiones estilo
                END IF
                IF stage=-1 AND edkCompromisos() AND not edk.hideSelfResults AND mov<>"PDF" THEN %>
                <tr class="noshowimp">
                  <td colspan="3" align="center">
                    [<a href="#" onClick="return verCompromisos();"><%=lblED_CompromisosDeDesarrollo%></a>]
                  </td>
                </tr> <%
                END IF %>
              </TABLE>
            </td>
          </tr>
        </table>
        <div id="toolTip" style="display:none;width:100px;height:20px;z-index:43;" class="tooltip">&nbsp;</div>
        <input type="hidden" id="childwin" name="childwin" value="<%=serverHTMLencode(reqplain("childwin"))%>">
        <input type="hidden" id="mov" name="mov" value="">
        <input type="hidden" id="edkrep" name="edkrep" value="<%=bool2num(edkrep)%>">
      <%
        IF allowView THEN
          dim colspan, estilo, s, objS, r, objR, numRseccion
          dim numR : numR = 0
          for s=1 to edk.regEvaK.colSeccion.count
            set objS = edk.regEvaK.colSeccion.obj(s)
            tstyle = iif( (objS.conPeso or (mov="PDF") or (stage=0)) and (objS.lblRubro<>""), "", "display:none;" )
            objS.sumPeso = 0
            objS.sumResultado = 0 %>
        <br />
        <table border="0" cellspacing="0" cellpadding="1" align="center" id="edRK" style="page-break-inside:avoid;<%=tstyle%>"> <%
            colspan = edk.seccionTableTitles( stage, objS, false, false )
            numRseccion = 0
            '-- rubros
            for r=1 to edk.colRubroK.count
              set objr = edk.colRubroK.obj(r)
              if objr.IdSeccionK = objS.IdSeccionK then
                if mov<>"PDF" then objr.getHistory
                estilo = switchEstilo(estilo)
                edk.seccionTableRow stage, objS, objR, numR, allowEvaluate, allowEdit, estilo
                objS.sumPeso = objS.sumPeso + objr.PesoRubro
                objS.sumResultado = objS.sumResultado + objr.rev.Resultado
                numRseccion = numRseccion + 1
                numR = numR + 1
              end if
            next
            '-- completa la seccion con rubros vacios hasta el maximo de rubros permitido
            if allowEdit AND (objs.TipoSeccionK <> EDK_TIPOCOM) AND (objs.TipoSeccionK <> EDK_TIPOORG) AND (objs.TipoSeccionK <> EDK_TIPOEST) AND _
              NOT (objs.TipoSeccionK = EDK_TIPOCOT AND objs.MaxRubros > 0 AND numRseccion >= objs.MaxRubros) then
              for r=numRseccion+1 to objs.MaxRubros
                set objr = new EDK_Rubro
                objr.IdSeccionK = objS.IDSeccionK
                objr.PesoRubro = objS.PesoFijoRubro
                estilo = switchEstilo(estilo)
                edk.seccionTableRow stage, objS, objR, numR, allowEvaluate, allowEdit, estilo
                objS.sumPeso = objS.sumPeso + objr.PesoRubro
                numRseccion = numRseccion + 1
                numR = numR + 1
              next
            end if
            '-- totales
            edk.seccionTotalRow stage, objS, colspan, allowEvaluate, allowEdit %>
        </table> <%
          next
          '--- Resumen de resultados y comentarios
          evaluacionResumen stage, adminEventos
          edk.evaluacionComments stage
          edPrint_Firmas edk.regGP.nombreEvaluador, edk.regGP.nombreEvaluado
          '--- Historial
          if mov<>"PDF" then
            edk.evaluacionHistory
          end if
          '--- Objetivos Estrategicos
          edk.evaluacionSelOE stage, allowEdit
        END IF
        IF mov<>"PDF" THEN
          if allowEdit OR allowEvaluate then %>
          <IFRAME name="notimeoutASP" border="0" width="0" height="0" src="./khorKeepAlive.asp?<%=sesion2queryString()%>"></IFRAME>
          <input type="hidden" id="<%=sesionReqKey%>" name="<%=sesionReqKey%>" value="<%=sesionEncrypted()%>">
          <input type="hidden" id="IdUA" name="IdUA" value="<%=edk.regGP.IdUnidadAdministrativa_EDP%>"> <%
            if modo<>"PER" AND numRev=0 AND autoAllowEdit(edk.idxRevision) AND not edk.regEvaK.allFixed then %>
          <div align="center">
            <input type="checkbox" name="edkLocked" value="1"<%=checkedIf(edk.regGP.edkLocked)%>> <%=lblEDP_ImpedirCambiosObjetivos%>
          </div> <%
            end if
            if not (edk.regEvaK.allFixed AND stage=0) then
              stepOptions = lblFRS_Guardar & "||guardar()"
            end if
          end if
          defaultFormEnd stepOptions, "", true
        ELSE
          defaultFormEnd "", "", false
        END IF
        
  set edk = nothing
  set colPeriodos = nothing
  layoutEnd
'================================================================================'
%>
