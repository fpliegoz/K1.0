<!--#include file="./khorClass.asp"-->
<!--#include file="./UVMCortesEvaluacionClass.asp"-->
<%
  thispage = "UVMCortesEvaluacionGlobalEdit.asp"
  checaSesion ses_super&","&ses_adminid, "", ""
  modulosActivos = khorModulosActivos()
  
  validaEntrada khorPermisoModulo(Modulo_CatalogosMantenimiento,modulosActivos) or khorPermisoModulo(Modulo_CatalogosBajas,modulosActivos) or khorPermisoModulo(Modulo_CatalogosConsulta,modulosActivos), "", ""
  
  mov = reqplain("mov")
  dirty = (reqn("dirty")<>0)
  gMov = ucase(request("gMov"))
  nt = 0
 
  cat = CAT_UVM_CORTEEVAGLOBAL
  
  set oCat = new clsCatalogo
  oCat.getFromDB conn, cat

  '-- Inicializa corte de evaluación
  dim colrev : set colrev = new frsCollection
  dim reg : set reg = new corteEvaluacion
  reg.getFromDB conn, reqn("IdCorteEvaluacion")
  
  if dirty then
    reg.CorteEvaluacion = reqs("CorteEva")
    reg.IdPeriodo = reqn("IdPeriodo")
    reg.Activo = 1
    reg.FechaIni = getDateFromDMYstr(request("txtPerIni"))
    reg.FechaFin = getDateFromDMYstr(request("txtPerFin"))
  end if
  
  if gMov="DOIT" then
    errmsg = ""
    'if reloaded=0 then
      if mov="B" then
        sqCorte = "DELETE FROM UVM_CorteEvaluacion WHERE IdPeriodo = "&reg.IdPeriodo&" AND Modo IN (1,3)"
        conn.execute sqCorte
      else
        sqCorte = "SELECT * FROM UVM_CorteEvaluacion WHERE IdPeriodo = "&reg.IdPeriodo&" AND Modo = 1"
        set rsCorte = getrs (conn,sqCorte)
        IdCorte = reg.IdCorteEvaluacion
        if rsCorte.EOF or reg.IdCorteEvaluacion <> 0 then
          set rsNiv = getrs(conn,"SELECT * FROM UVM_CatNivelesCap")
          sqFac = "SELECT DISTINCT Dimension, ids = STUFF((SELECT ' , ' + CAST (IdDimension AS VARCHAR) "&_
                      "     FROM UVM_Dimension b WHERE b.Dimension = a.Dimension FOR XML PATH('')), 1, 2, ''), "&_
                      " sec = STUFF((SELECT TOP 1 ' , ' + CAST (Seccion AS VARCHAR) FROM UVM_Dimension c "&_
                      "     JOIN UVM_CatSeccCritEva e ON c.IdSeccion = e.IdSeccion "&_
                      "     WHERE c.Dimension = a.Dimension FOR XML PATH('')), 1, 2, '') "&_
                      "FROM UVM_Dimension a GROUP BY Dimension ORDER BY sec;"
          while not rsNiv.EOF
            set rsDimensiones = getrs(conn,sqFac)
            reg.CorteEvaluacion = reqs("CorteEva") & " - " & rsStr(rsNiv,"Nivel")
            while not rsDimensiones.EOF 
              reg.IdNivel = rsNum(rsNiv,"IdNivelCap")
              
              reg.Tipo  = Right(rsStr(rsDimensiones,"ids"), Len(rsStr(rsDimensiones,"ids"))-1)
              reg.Modo = 3
              if not reg.update(conn) then
                errmsg = "Ya existe un periodo con el ciclo, nivel y factor seleccionados."
              end if
              logAcceso LOG_ALTA, lblUVM_CortesEvaluacion, "Cortes Evaluacion IdCorte (" & reg.IdCorteEvaluacion & ")"
              reg.IdCorteEvaluacion = 0
              rsDimensiones.movenext
            wend  
            rsNiv.movenext
          wend
          rsDimensiones.close
          set rsDimensiones = nothing
          rsNiv.close
          set rsNiv = nothing
          reg.IdCorteEvaluacion = IdCorte
          reg.CorteEvaluacion = reqs("CorteEva")
          reg.IdPeriodo = reqn("IdPeriodo")
          reg.IdNivel = 0
          reg.Tipo  = 0
          reg.Activo = 1
          reg.Modo  = 1
          reg.FechaIni = getDateFromDMYstr(request("txtPerIni"))
          reg.FechaFin = getDateFromDMYstr(request("txtPerFin"))
          reg.update(conn)
          logAcceso LOG_ALTA, lblUVM_CortesEvaluacion, "Cortes Evaluacion IdCorte (" & reg.IdCorteEvaluacion & ")"
        else  
          errmsg = "Ya existe un periodo global con el periodo seleccionado."
        end if
        rsCorte.close
        set rsCorte = nothing
      end if
    'end if
    if errmsg = "" then
      redirect "UVMCortesEvaluacionGlobal.asp?fromcat="&request("fromcat")
    end if
    
  end if
  
  ctrlState = ""
  if mov="A" then
    tit1 = lblFRS_Agregando
    btnAccion = iif( oCat.habMantto, lblFRS_Guardar, "" )
  else
    if mov="B" then
      tit1 = lblFRS_Eliminando
      btnAccion = iif(oCat.catBorrable(reg.IdCorteEvaluacion), lblFRS_Eliminar, "" )
      ctrlState = " disabled"
    else
      tit1 = lblFRS_Modificando
      btnAccion = iif( oCat.habMantto, lblFRS_Guardar, "" )
    end if
  end if
  if btnAccion="" then
    tit1 = lblFRS_Consultando
    ctrlState = " disabled"
  end if
  tit1 = tit1 & " " & lblUVM_CortesEvaluacion
  tit = lblFRS_CatalogosDelSistema
  'tit2 = lblUVM_CortesEvaDescrEdit
  set oCat = nothing

'================================================================================
layoutHeadStart khorAppName() & " - " & tit & " - " & tit1
includeJS
%>
<SCRIPT LANGUAGE=javascript>
<!--
  var tabcount;
  var tab;
  function tabSel(t) {
    if (t!=tab) {
      for (var i=0; i<tabcount; i++) {
        setVisible('tabC'+i,i==t);
        ot=MM_findObj('tab'+i);
        ot.className=(i==t)?'tabOpen':'tabClose';
      }
      tab=t;
      setValor('curtab',tab);
    }
  }

  function setDirty() {
    setValor('dirty',1);
  }
  function regresar(){
    if ( getValor('dirty') !=0 ) {
      if ( !confirm("<%=strJS(lblFRS_abandonarCambios)%>") ) return;
    }
    document.TrueForm.action="UVMCortesEvaluacionGlobal.asp";
    sendval('','mov','');
  }
  function getDateObjFromTxtObj(fieldId){
    var fieldObj = $("#"+ fieldId);
    return getDateObjFromStr(fieldObj.val());
  }
  function getDateObjFromStr(dateStr){
    if( dateStr=="" ) return "";
    var arrFecha = dateStr.split("/");
    return new Date(arrFecha[2], parseInt(arrFecha[1])-1, arrFecha[0]);
  }
  function aceptar() {
  <%IF mov="B" THEN%>
    if (!confirm('<%=strJS(lblFRS_confirmaEliminar)%>')) return;
  <%ELSE%>
    var obj = MM_findObj('CorteEva');
    var IdPeriodo  = $("#IdPeriodo").val();
    if ( !validaStr( obj,true ) ) {
      alert('<%=strJS(lblFRS_DebeIngresarUnNombreDe_ & lblUVM_CorteEvaluacion)%>');
      obj.focus();
      obj.select();
      return;
    }else if(IdPeriodo == -1){
      alert('<%=strJS("Se debe seleccionar un Ciclo.")%>');
      return;
    }
    
    var date     = new Date(Date.now());
    var currentTime = date.getDate() + '/' + (date.getMonth() + 1) + '/' +  date.getFullYear();
    var fIni = getDateObjFromTxtObj('txtPerIni');
    var fFinC = getDateObjFromTxtObj('txtPerFin');
    
    if (fIni == "" || fFinC == ""){
      alert('<%=strJS(lblUVM_FechaFaltante)%>');
      return;
    }else if(fIni > fFinC ){
      alert('<%=strJS(lblUVM_FechaIniMenor)%>');
      return;
    }
  <%END IF %>
    sendval('','gMov','DOIT');
  }

  function setCalendar() {
    setDirty();    
  }
  
//-->
</SCRIPT>
<%
'----------------------------------------
layoutHeadEnd
layoutStart tit, tit1, tit2, "", khorWinWidth(), ""
defaultFormStart thispage, "", true
'----------------------------------------
%>
       <%IF errmsg<>"" THEN%><div class="alerta"><%=errmsg%></div><%END IF%>
        <TABLE cellSpacing="1" cellPadding="1" border="0" align="center">
          <TR class="celdaDark">
            <TD>Nombre del Periodo:</TD>
            <TD colspan="2">
              <INPUT style="WIDTH:200px" id="CorteEva" name="CorteEva" value="<%=serverHTMLencode(reg.CorteEvaluacion)%>"<%=ctrlState%> onChange="setDirty();" class="whiteblur" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this);" onmouseout="inOut(this);">
            </TD>
          </TR>
          <TR class="celdaLight">
            <TD>Ciclo:</TD>
            <td>
              <select name="IdPeriodo" id="IdPeriodo" style="font-size:10px;" class="whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);"><%
              sqCiclo = "SELECT IdPeriodo,(CAST (IdPeriodo AS VARCHAR)+'-'+CAST(Anio AS VARCHAR)+'-'+CAST (ciclo AS VARCHAR) ) AS Ciclo FROM UVM_CatCiclos ORDER BY IdPeriodo DESC"
               
              set rsCiclo = getrs(conn,sqCiclo)
              if rsCiclo.EOF then%>
                <option value="-1">No hay ciclos registrados</option><%
              else%>
                <option value="-1">Seleccionar Ciclo</option><%
                while not rsCiclo.EOF%>
                <option <%=iif(reg.IdPeriodo = rsNum(rsCiclo,"IdPeriodo"),"selected='selected'","")%> value="<%=rsNum(rsCiclo,"IdPeriodo")%>" ><%=rsStr(rsCiclo,"Ciclo")%></option><%
                  rsCiclo.movenext
                wend
              end if
              rsCiclo.close
              set rsCiclo = nothing%>
              </select>
            </td>
          </TR>
          <TR class="celdaDark">
            <TD><%=lblUVM_FechaInicio%>:</TD>
            <TD>
              <INPUT id=txtPerIni name=txtPerIni value="<%=formatDateDMAnull(reg.FechaIni)%>" class="whiteblur" style="WIDTH:80px;" readonly="true"> <%
              if ctrlState="" then %>
              <A href="#" onclick="khorCalendar('txtPerIni'); return false;"><IMG src="khorImg/ico_calendar.gif" align="middle" border="0" height="16" width="16" alt="<%=lblFRS_ClickParaSeleccionarFecha%>"></A> <%
              end if %>
            </TD>
          </TR>
          <TR class="celdaLight">
            <TD>Fecha de T&eacute;rmino:</TD>
            <TD>
              <INPUT id=txtPerFin name=txtPerFin value="<%=formatDateDMAnull(reg.FechaFin)%>" class="whiteblur" style="WIDTH:80px;" readonly="true"> <%
              if ctrlState="" then %>
              <A href="#" onclick="khorCalendar('txtPerFin'); return false;"><IMG src="khorImg/ico_calendar.gif" align="middle" border="0" height="16" width="16" alt="<%=lblFRS_ClickParaSeleccionarFecha%>"></A> <%
              end if %>
            </TD>
          </TR>
        </TABLE>

        <script language="JavaScript">
          tabcount = <%=nt%>;
          tabSel(<%=reqn("curtab")%>);
        </script>
        <input type="hidden" name="fromcat" value="<%=request("fromcat")%>">
        <INPUT type=hidden id="gMov" name="gMov" value="">
        <INPUT type=hidden id="dirty" name="dirty" value="<%=bool2num(dirty)%>">
        <INPUT type=hidden id="mov" name="mov" value="<%=mov%>">
        <INPUT type=hidden id="IdCorteEvaluacion" name="IdCorteEvaluacion" value="<%=reg.IdCorteEvaluacion%>">
        <INPUT type=hidden id="curtab" name="curtab" value="<%=reqn("curtab")%>">
<%
  extraBtns = iif( btnAccion="", "", btnAccion & "||aceptar()" )
  extraBtns = strAdd( extraBtns, "@@", lblFRS_Regresar & "||regresar()" )
  colRev.clean
  set colRev = nothing
  set permod = nothing
  set reg = nothing
'----------------------------------------
defaultFormEnd extraBtns, "", false
layoutEnd
'----------------------------------------
%>
