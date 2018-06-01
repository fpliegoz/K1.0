<!--#include file="./frsSHA256.asp"-->
<!--#include file="./khorClass.asp"-->
<!--#include file="./khorBD.asp"-->
<!--#include file="./coppelLoginAux.asp"-->
<%
  Modificacion unica de coppel files
  thispage = "coppelLogin.asp"
  loginConfigURL = khorLoginURL()
  'loginForceConfigURL: variable dirty para forzar a que se use pagina especial de login, nunca esta
  if loginForceConfigURL AND loginConfigURL<>"" AND ucase(loginConfigURL) <> ucase(thispage) then
    redirect loginConfigURL
  end if
  '--------------------------------------------------------'
  ' falta agregar esta url como parametro de configuración
  urlPostulacion = "./pbtPostulacionOferta.asp"
  '--------------------------------------------------------'
  omiteMenuLateral = true

  dim escolar             : escolar         = khorEscolar()
  dim embedded            : embedded        = (reqn("embedded")<>0)
  dim fromReq             : fromReq         = (reqn("fromReq")<>0)
  dim destination
  dim modo
  dim tipoLogin           : tipoLogin       = khorTipoLogin()
  dim multiSucursal       : multiSucursal   = khorMultiSucursal()
  dim encriptaPwd         : encriptaPwd     = khorEncriptaPasswords()
  dim tipoPassword
  dim conAutoregistro
  dim conRecuperacionPwd  
  dim loginSucursal
  dim loginTipoPer
  dim loginUsr
  dim loginPwd
  dim loginPwdField
  dim tipoMantto          : tipoMantto      = khorModoMantto
  dim enMantto            : enMantto        = (tipoMantto <> NO_MANTTO)
  dim IdPrefijoForFB      : IdPrefijoForFB  = CInt(khorConfigValue(549,true))
  dim FB_AppId            : FB_AppId        = khorConfigValue(550,false)
  dim sharePrefGrales     : sharePrefGrales = iif(multiSucursal, CInt(khorConfigValue(69,true))=1, true)
  'COPPEL ONLY
  dim perType             : perType = iif(reqn("perType")=0,2,reqn("perType"))
  
  canUseFacebook  = (IdPrefijoForFB<>0 AND tipoLogin=0 AND FB_AppId<>"" AND sharePrefGrales)
  LnkdIn_AppId    = khorConfigValue(554,false)
  IdPrefijoForLI  = CInt(khorConfigValue(555,true))
  canUseLnkdIn    = (IdPrefijoForLI<>0 AND tipoLogin=0 AND LnkdIn_AppId<>"" AND sharePrefGrales)
  
  idOferta = reqn("IdOferta")
  myReturnPath = thispage & "?IdOferta=" & idoferta

  IF superSesion() OR adminSesion()<>0 OR personaSesion()<>0 THEN
    destination = getSesionMenu()
    errmsg = ""
  ELSE
    '-- Titulos y mensajes
    tit1 = lblFRS_IdentificacionDeUsuario
    tit2 = ""
    '-- Proceso
    modo = procesaModo
    if escolar and modo <> "admin" then
      modo = "user"
    end if
    cssContext = modo
    portal = reqn("Portal")
    if portal = 0 then
      portal = iif( CStr(session("portal"))="", 0, session("portal") )
    end if
    if portal = 0 then
      session("portal") = ""
      session("accesoDesdePortal") = ""
    else
      session("portal") = 1
      session("accesoDesdePortal") = 1
    end if
    errmsg = iif((modo="user" AND tipoMantto <> NO_MANTTO) OR (modo="admin" AND tipoMantto = MANTTO_BD),lblKHOR_EnMantenimiento,lblFRS_DebeEstarRegistrado)
    if modo<>"" then
      loginUsr = cleanString(trim(reqsF("usr")))
      mov = reqs("mov")
      if modo="admin" then
        tit2 = lblKHOR_Administrador
        loginSucursal = 0
        loginTipoPer = 0
        loginPwdField = "pwd"
        lblUsr = lblKHOR_ClaveDeUsuario_Admin
      else
        if escolar and not fromReq then
          redirect "./khorEscolar/khorEscolar.asp"
        end if
        modo = "user"
        tit2 = lblKHOR_Persona
        if tipoLogin=1 then
          loginSucursal = reqn("loginSucursal")
          loginTipoPer = reqn("loginTipoPer")
        else
          loginTipoPer = 0
        end if
        lblUsr = khorLoginLabel(tipoLogin, khorDefaultPais())
        tipoPassword = khorTipoPassword()
        if tipoPassword=2 then
          loginPwdField = reqs("pwdsel")
          if mov="submit" and loginPwdField="" then
            loginPwdField = passwordInit(loginUsr,loginTipoPer)
            mov=""
          end if
        else
          loginPwdField = "pwd"
        end if
        conAutoregistro = (khorConfigValue(39,true)=0) AND NOT enMantto
        conRecuperacionPwd = ((tipoPassword<>1) and (khorConfigValue(397,true)<>1) and (mail_Remitente()<>"") and (not encriptaPwd))
        lblPwd = passwordLabel(loginPwdField)
      end if
      if mov="submit" then
        errmsg=lblFRS_badLogin
        loginPwd = reqPwd("pwd")
        limpiaSesion
        if loginPwd<>"" or ((modo<>"admin") AND (tipoPassword=1)) then
          '-- Coppel
          if perType=2 then
            loginResult = autentifica(modo,tipoLogin,loginUsr,loginPwd,loginTipoPer,loginPwdField)
          else
            loginResult = autentifica(modo,0,loginUsr,loginPwd,reqn("loginTipoPer"),loginPwdField)
          end if
          if loginResult then
            errmsg = ""
          end if
        end if
      end if
    end if
    IF fromReq THEN
      if reqs("dest") <> "" then
        newDestination = reqs("dest")
      end if
      if reqs("modoOverride") <> "" then
        modo = reqs("modoOverride")
      end if
      if autentifica(modo,tipoLogin,reqs("usr"),reqs("pwd"),loginTipoPer,loginPwdField) then
        errmsg = ""
      else
        if newDestination <> "" then
          redirect newDestination & "?err=" & errmsg
        end if
      end if
      if newDestination <> "" then
        destination = newDestination
      end if
    END IF
  END IF
  
  if errmsg="" then
    if embedded then
      redirect "khorLoginE.asp?destination=" & destination
    else
      if personaSesion()<>0 and IdOferta > 0 then
        destination = urlPostulacion & "?idOferta=" & idOferta
      end if
      if requiresPrivacyCheck() then
        redirect "frsPrivacidad.asp?returnpath="&destination
      else
        if superSesion() OR adminSesion()<>0 then
          res = asynchronousCall("./cleanUploads.ashx", "POST", "")
        end if
        redirect destination
      end if
    end if
  end if

  '----------------------------------------

  function requiresPrivacyCheck()
    if modo="user" AND privacyActive() then
      requiresPrivacyCheck = (getBDnum("x","SELECT (CASE WHEN AceptaTerminos IS NULL THEN 0 ELSE AceptaTerminos END) AS x FROM Personal WHERE IdPersonal="&personaSesion()) <> 1)
    else
      requiresPrivacyCheck = false
    end if
  end function

  function procesaModo()
    dim modo : modo = lcase(reqs("modo"))
    'Configuracion "dirty", definir en khorLabelEspecial.asp
    'loginNoCookie = true  'no guardar cookie, forza a que siempre inicie con seleccion de admin/persona
    if modo="select" then
      modo = ""
    elseif modo="" then
      modo = iif(loginNoCookie, "", request.Cookies("modoLogin") )
    elseif modo <> "user" and modo <> "admin" then
      modo = ""
    end if
    if not loginNoCookie then response.Cookies("modoLogin") = modo
    procesaModo = modo
  end function

  function passwordInit(usr,tipoper)
    dim auxs, auxa, i
    dim qrysel, qrytab
    dim sq, rs
    '-- Campos configurados
    auxs = khorConfigValue(245,false)
    if auxs<>"" then
      '-- Forma query para obtener todos estos campos
      auxa = split(auxs,",")
      auxs = ""
      qrysel = ""
      qrytab = ""
      for i=lbound(auxa) to ubound(auxa)
        auxa(i) = trim(auxa(i))
        select case auxa(i)
          case "amm"
            qrysel = qrysel & ", madre.ApellidoMaterno AS " & auxa(i)
            qrytab = qrytab & " LEFT JOIN PersonalFamilia madre ON repPersonal.IdPersonal=madre.IdPersonal AND madre.IdTipoFamilia=2"
          case "amp"
            qrysel = qrysel & ", padre.ApellidoMaterno AS " & auxa(i)
            qrytab = qrytab & " LEFT JOIN PersonalFamilia padre ON repPersonal.IdPersonal=padre.IdPersonal AND padre.IdTipoFamilia=1"
          case "rfc"
            qrysel = qrysel & ", repPersonal.RFC AS " & auxa(i)
          case "curp"
            qrysel = qrysel & ", repPersonal.CURP AS " & auxa(i)
          case "enac"
            qrysel = qrysel & ", repPersonal.EstadoNac AS " & auxa(i)
          case "cnac"
            qrysel = qrysel & ", repPersonal.PoblacionNac AS " & auxa(i)
        end select
      next
      if qrysel<>"" then
        '-- Con el query, forma de nuevo la lista con aquellos donde la persona tiene datos
        sq = "SELECT repPersonal.IdPersonal" & qrysel & " FROM repPersonal" & qrytab
        if tipoLogin=1 then
          sq = sq & " WHERE repPersonal.IdPrefijo="&tipoper&" AND repPersonal.Folio="&cdbl(getVal(usr,true))
        else
          sq = sq & " WHERE repPersonal." & khorLoginField(tipoLogin) & "='"&sqsf(usr)&"'"
        end if
        set rs = getrs(conn,sq)
        if not rs.eof then
          for i=lbound(auxa) to ubound(auxa)
            if trim(rs(auxa(i)))<>"" then
              auxs = strAdd(auxs,",",auxa(i))
            end if
          next
        end if
        rs.close
        set rs = nothing
        '-- Ahora si, selecciona aleatoriamente uno de ellos
        if auxs<>"" then
          randomize second(now)
          auxa = split(auxs,",")
          auxr = Rnd
          i = Round(cdbl(ubound(auxa)) * auxr)
          auxs = auxa(i)
        end if
      end if
    end if
    passwordInit = auxs
  end function

  function autentifica(modo,tipoLogin,usr,pwd,tipoper,pwdfield)
    dim sq, rs, auxId, retval
    dim campoPassword, joinPassword
    dim tabla, campoId
    dim p_tipo, s_tipo, s_id
    dim attempts, maxattempts, bloqueado
    dim forcepwdchg : forcepwdchg = ""
    retval = false
    if modo="user" and sqsf(usr)<>"" and tipoMantto = NO_MANTTO then
      joinPassword = ""
      select case pwdfield
        case "amm"
          campoPassword = "PersonalFamilia.ApellidoMaterno"
          joinPassword = " INNER JOIN PersonalFamilia ON Personal.IdPersonal=PersonalFamilia.IdPersonal AND PersonalFamilia.IdTipoFamilia=2"
        case "amp"
          campoPassword = "PersonalFamilia.ApellidoMaterno"
          joinPassword = " INNER JOIN PersonalFamilia ON Personal.IdPersonal=PersonalFamilia.IdPersonal AND PersonalFamilia.IdTipoFamilia=1"
        case "rfc"
          campoPassword = "Personal.RFC"
        case "curp"
          campoPassword = "Personal.CURP"
        case "enac"
          campoPassword = "geoEstado.Estado"
          joinPassword = " LEFT JOIN geoEstado ON Personal.IdEstadoNac=geoEstado.IdEstado"
        case "cnac"
          campoPassword = "geoPoblacion.Poblacion"
          joinPassword = " LEFT JOIN geoPoblacion ON Personal.IdPoblacionNac=geoPoblacion.IdPoblacion"
        case else
          campoPassword = "Personal.Password"
      end select
      sq = "SELECT Personal.IdPersonal, " & campoPassword & " AS pwdToCheck, Bloqueado, coalesce(FechaBloqueo, getdate()) FechaBloqueo, IntentosDeIngreso FROM Personal" & joinPassword & " WHERE Personal.StatusPer=0"
      IF multiSucursal THEN sq = sq & " AND Personal.IdSucursal NOT IN (Select IdSucursal from CatSucursal WHERE CatSucursal.Activa=0 OR CatSucursal.FechaVigencia < " & formatDateSQL(Date,false) & ")"
      if tipoLogin=1 then
        sq = sq & " AND Personal.IdPrefijo="&tipoper&" AND Personal.Folio="&cdbl(getVal(usr,true))
      else
        sq = sq & " AND Personal." & khorLoginField(tipoLogin) & "='"&sqsf(usr)&"'"
      end if
      p_tipo = tipoPassword
      s_tipo = ses_userid
      tabla = "Personal"
      campoId = "IdPersonal"
    elseif modo="admin" then
      if usr="" AND ucase(pwd)=bdpwd() then
        iniciaSesion ses_super, pwd
        destination = getSesionMenu()
        retval = true
      elseif tipoMantto = MANTTO_BD_ADMIN or tipoMantto = NO_MANTTO then
        sq = "SELECT IdUsuario, Password AS pwdToCheck, Bloqueado, coalesce(FechaBloqueo, getdate()) FechaBloqueo, IntentosDeIngreso FROM DBUsuarios WHERE (Activo<>0) AND (FechaVigencia IS NULL OR FechaVigencia >= " & formatDateSQL(now, false) & ") AND Usuario='"&sqsf(usr)&"'"
        IF multiSucursal THEN sq = strAdd( sq, " AND " , "IdUsuario NOT IN (Select IdUsuario from DBUsuarioSucursal LEFT JOIN CatSucursal ON DBUsuarioSucursal.IdSucursal=CatSucursal.IdSucursal WHERE CatSucursal.Activa=0 OR CatSucursal.FechaVigencia < " & formatDateSQL(Date,false) & ")" )
        p_tipo = 0
        s_tipo = ses_adminid
        tabla = "DBUsuarios"
        campoId = "IdUsuario"
      end if
    end if
    if sq<>"" and not retval then
      if (modo="user" and not enMantto) or (modo="admin" and tipoMantto = MANTTO_BD_ADMIN or tipoMantto = NO_MANTTO) then
        s_id = 0
        set rs = getrs(conn,sq)
        if not rs.eof then
          if encriptaPwd then
            passEnc = ucase(sha256(pwd))
            KHOR_DEFAULT_ADMIN_PWD = sha256(KHOR_DEFAULT_ADMIN_PWD)
          else
            passEnc = ucase(pwd)
          end if
          s_id = rsNum(rs,campoId)
          attempts = rsNum(rs,"IntentosDeIngreso")
          bloqueado = rsNum(rs,"Bloqueado")
          if bloqueado = 1 and dateadd("n", Khor_MinutosBloqueo, cdate(rs("FechaBloqueo"))) > now  then
            errmsg = lblFRS_EstaCuentaSeEncuentraBloqueada_ContacteAlAdministrador
            pwd = ""
          end if
          pwdToCheck = ucase(rs("pwdToCheck"))
          if ((p_tipo<>1) AND (pwd<>"") AND pwdToCheck=passEnc) OR ((p_tipo=1) AND pwd="") then
            'Limpia Failed Attempts
            conn.execute "UPDATE " & tabla & " SET IntentosDeIngreso=0, Bloqueado = 0 WHERE " & campoId & "=" & s_id
            retval = true
          end if
          if retval AND modo="admin" then
            forcepwdchg = iif( ucase(usr)=ucase(KHOR_DEFAULT_ADMIN_USR) AND pwdToCheck=ucase(KHOR_DEFAULT_ADMIN_PWD), "reset", "" )
          end if
        end if
        rs.close
        set rs = nothing
        if retval then
          if khorSesionUnica() then
            if not isLiveSessionQry(s_tipo, s_id) then
              iniciaSesion s_tipo, s_id
              destination = getSesionMenu()
              necesitaCambioPwd s_id, modo, forcepwdchg
              retval = true
            else
              retval = false
              errmsg = lblFRS_NoFuePosibleAccederALaInformacion & "<br>" & lblFRS_SessionAlreadyOpen
            end if
          else
            iniciaSesion s_tipo, s_id
            destination = getSesionMenu()
            necesitaCambioPwd s_id, modo, forcepwdchg
            retval = true
          end if
        elseif s_id > 0 then
          '-- Verifica si hay limite de intentos
          maxattempts = khorConfigValue(279,true)
          if maxattempts > 0 then
            if bloqueado = 0 then
              conn.execute "UPDATE " & tabla & " SET IntentosDeIngreso=IntentosDeIngreso+1 WHERE " & campoId & "=" & s_id
              attempts = attempts + 1
              if cint(attempts) >= cint(maxattempts) then
                conn.execute "UPDATE " & tabla & " SET Bloqueado=1, FechaBloqueo = getdate() WHERE " & campoId & "=" & s_id
              end if
            end if
          end if
          '-- Registra login fallido
          logAcceso LOG_GENERICO, "Login Failed", campoId & ": " & s_id
        end if
      else
        errmsg = lblKHOR_EnMantenimiento
      end if
    end if
    autentifica = retval
  end function

  function necesitaCambioPwd(id,tipo,forced)
    dim tabla,idf,sq,rs,retval
    dim f_password, razon
    if forced <> "" then
      retval = forced
    else
      necesitaCambioPwd = false
      if tipo = "user" then
        tabla = "Personal"
        idf = "IdPersonal"
      elseif tipo = "admin" then
        tabla = "DBUsuarios"
        idf = "IdUsuario"
      else
        exit function
      end if
      retval = ""
      if ansiSQL="yes" then
        sq = "SELECT (" & formatDateSQL(Now,true) & "-FechaUltimoCambioPwd) AS Dias"
      else
        sq = "SELECT DATEDIFF(day,FechaUltimoCambioPwd," & formatDateSQL(Now,true) & ") AS Dias"
      end if
      sq = sq & " FROM " & tabla & " WHERE " & idf & " = " & id & " AND FechaUltimoCambioPwd IS NOT NULL"
      set rs = getrs(conn,sq)
      if not rs.EOF then
        dim diasparacambio : diasparacambio = cint(khorConfigValue(282,true))
        retval = iif( (rsNum(rs,"Dias") >= diasparacambio) and (diasparacambio>0), "expired", "" )
      elseif (khorConfigValue(280,true) <> 0) then
        retval = "first"
      end if
      rs.close
      set rs = nothing
      if retval = "" then
        f_password = khorStrResetPassword()
        if f_password<>"" then
          if encriptaPwd then
            f_password = ucase(sha256(f_password))
          end if
          retval = iif( bdExists( "SELECT " & idf & " FROM " & tabla & " WHERE " & idf & "=" & id & " AND Password='" & f_password & "'" ), "reset", "" )
        end if
      end if
    end if
    if retval <> "" then
      session("mustChangePwd") = true
      if tipo = "user" then
        destination = khorCambioPasswordURL()&"?f=" & retval
      elseif tipo = "admin" then
        destination = "usuarioSetPwd.asp?f=" & retval
      end if
      session("pwdChangeURL") = destination
    end if
    necesitaCambioPwd = (retval <> "")
  end function
  
  sub paintSocialLogin()
    IF canUseFacebook OR canUseLnkdIn THEN %>
    <script>
      /* Configuracion de los parametros de función Ajax sólo para login */
      ajaxParams.error = function() { killBlocker(); alert("<%=strJS(lblFRS_ErrorAlObtenerInfo)%>"); };
      ajaxParams.success = function(data) {
        if( isNaN(data) ){
          killBlocker();
          alert(<%=iif(conAutoregistro, "'"& strJS(lblFRS_Error) &": '+ data", "'"& strJS(lblFRS_badLoginBySocialNetwrk) &"'")%>);
        }else
          window.location.href = "./khorFBredirect.asp?IdUsuario="+data+"&Tipo=Per&IdOferta="+<%=IdOferta%>;
      };
    </script>
    <div id="loginDIVsocial" class="loginDIV" align="center">
      <%=iif(canUseFacebook AND canUseLnkdIn AND lblFRS_UsasFBoLnkdIn<>"","<br/>" & lblFRS_UsasFBoLnkdIn & "<br/>","")%>
      <table align="center" border="0" cellpadding="0" cellspacing="0" style="text-align:center;">
        <tr><%
      IF canUseFacebook THEN%>
          <td style="vertical-align:middle;">
            <div id="fb-root"></div>
            <script>
              window.fbAsyncInit = function() { FB.init({ appId:'<%=FB_AppId%>', status:true, xfbml:true, version:'v2.4' }); };
              (function(d, s, id){
                 var js, fjs = d.getElementsByTagName(s)[0];
                 if (d.getElementById(id)) {return;}
                 js = d.createElement(s); js.id = id;
                 js.src = "//connect.facebook.net/<%=lblFRS_FB_locale%>/all.js";
                 fjs.parentNode.insertBefore(js, fjs);
               }(document, 'script', 'facebook-jssdk'));
            </script>
            <%=strAdd( iif(NOT canUseLnkdIn,lblFRS_FB_TienesCuentaFB,""), "<br/>", lblFRS_FB_IniciaSesionConFB & "<br/>" )%>
            <div class="fb-login-button" data-max-rows="1" data-size="medium" data-show-faces="false" data-auto-logout-link="false" data-scope="email,user_birthday,user_work_history" onlogin="logInMngmt"></div>
          </td><%
      END IF%>
      <%IF canUseFacebook AND canUseLnkdIn THEN%><td style="width:20px; vertical-align:top;">&nbsp;</td><%END IF%><%
      IF canUseLnkdIn THEN%>
          <td style="vertical-align:middle;">
            <%=strAdd( iif(NOT canUseFacebook,lblFRS_LI_TienesCuentaLI,""), "<br/>", lblFRS_LI_IniciaSesionConLI & "<br/>" )%>
            <script type="IN/Login" data-onAuth="onLinkedInAuth"></script>
          </td><%
      END IF%>
        </tr>
      </table>
    </div><%
    END IF
  end sub

'========================================
Response.AddHeader "P3P","CP=""NCAO PSA OUR"""
layoutHeadStart khorAppName()
includeJS
'========================================
%>
<script language="JavaScript">
<!--
  function validaEntrada() {
    var perType = $("#perType").val();
    <% IF modo="admin" THEN %>
    var ousr=$("#usr")[0];
    var opwd=$("#pwd")[0];
    <% ELSE %>
    var ousr=(perType=="1")? $("#loginPostulante #usr")[0] : $("#loginEmpleado #usr")[0];
    var opwd=(perType=="1")? $("#loginPostulante #pwd")[0] : $("#loginEmpleado #pwd")[0];
    <% END IF %>
    var xusr=stripCharsInBag(ousr.value,';\'"%');
    var xpwd=stripCharsInBag(opwd.value,'<%=strJS(charsForbiddenInPassword)%>');
    if (xusr!=ousr.value || xpwd!=opwd.value) {
      <% IF modo="admin" THEN %>
      alert('<%=strJS(strLang(lblFRS_InvalidCharsIn_X_oPassword,lblUsr))%>');
      <% ELSE %>
      alert(('<%=strJS(strLang(lblFRS_InvalidCharsIn_X_oPassword,"<<userType>>"))%>').replace("<<userType>>",(perType=="2")? "<%=lblUsr%>":"<%=khorLoginLabel(0, khorDefaultPais())%>"));
      <% END IF %>
      return false;
    }
    return true;
  }
  function navLogin(destination) { <%
    IF embedded THEN %>
    navigateToURL( "khorLoginE.asp", "destination",escape(destination) );<%
    ELSE %>
    parseAndNavToURL( destination ); <%
    END IF %>
    return false;
  } <%
IF modo="user" THEN
  IF conAutoregistro THEN %>
  function registro() {
    return navLogin( '<%=khorNewPersonaURL()%>?IdOferta=<%=IdOferta%>&returnpath=<%=myReturnPath%>' );
  } <%
  END IF
  IF conRecuperacionPwd THEN %>
  function recuperacion() {
    return navLogin( './personalGetPwd.asp?modo=user&returnpath=<%=myReturnPath%>' );
  } <%
  END IF
  IF tipoLogin=2 OR tipologin=3 THEN %>
  function abreClaveHelper() {
    navigateToURL( "personalClaveHelper.asp", "_target","claveHelper", "_winparams","toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=450,height=<%=iif(tipoLogin=2,300,350)%>" );
    return false;
  }
  function setClave(v) {
    setValor('usr',v);
  } <%
  END IF %>
  function AplicaIns() {
    return navLogin( './insAplica.asp?returnpath='+escape('<%=thispage%>') );
  }
  function AplicaCO() {
    return navLogin( './comoAplica.asp?returnpath='+escape('<%=thispage%>') );
  }
  function AplicaCapEva() {
    return navLogin( './encuestaPersonaCurso.asp?returnpath='+escape('<%=thispage%>') );
  } <%
END IF
IF modo="admin" THEN
  'conRecuperacionPwdAdmin variable para que en la pantalla de login de los administradores aparezca el boton para recuperación de contraseña
  IF conRecuperacionPwdAdmin THEN %>
  function recuperacionAdmin() {
    return navLogin( './personalGetPwd.asp?modo=admin&returnpath=<%=myReturnPath%>' );
  } <%
  END IF
END IF %>
//-->
</script><%
IF (canUseLnkdIn OR canUseFacebook) AND modo="user" THEN%>
<script type="text/javascript" src="khorSocialFunctions.js.asp"></script><%
  IF canUseLnkdIn THEN%>
<script type="text/javascript" src="<%=iif(withSSL, "https://", "http://")%>platform.linkedin.com/in.js">
  api_key: <%=LnkdIn_AppId%>
  lang:  <%=iif(lblFRS_FB_locale="es_LA","es_ES",lblFRS_FB_locale)%>
</script><%
  END IF
END IF%>
<STYLE type="text/css">
<!--
.loginMSG {
  padding: 5px;
  text-align: center;
  font-weight: bold;
}
.loginDIV {
  margin-top: 10px;
  padding: 5px;
  text-align: center;
}
//-->
</STYLE>
<%
'========================================
layoutHeadEnd
'========================================
  IF embedded THEN
    childwin = true
    response.write "<BODY>"
    IF errmsg<>"" THEN %><div class="alerta"><%=errmsg%></div><% END IF
  ELSE
    'Configuracion "dirty", definir en khorLabelEspecial.asp
    'loginFormNoTitle = true  'mostrar forma de login sin titulo, mostrarlo en el titulo de pantalla como antes
    layoutStart khorAppName(), tit1, iif(loginFormNoTitle,tit2,""), errmsg, khorWinWidth(), "" 
      if not useKhorStyles() then
        contentPaintBeginTag
      end if
        
        txtLogin = trim(khorConfigTxt(iif(modo="admin",3,4)))
        if txtLogin="" then txtLogin = trim(khorConfigTxt(1))
        txtLoginLoc = khorConfigValueWithDefault(494, false, "LEFT")
        if txtLoginLoc = "TOP" or txtLoginLoc = "BOTTOM" then
          fullTxtLogin = "<tr><td colspan='2' width='100%'><div id='loginDIVtxt' class='dashazul'>" & txtLogin & "</div></td></tr>"
        else
          fullTxtLogin = "<td width='50%'><div id='loginDIVtxt' class='dashazul'>" & txtLogin & "</div></td>"
        end if
        %>
        <table border="0" cellspacing="2" cellpadding="2" align="center" width="100%">
          <%
          IF txtLogin<>"" AND txtLoginLoc = "TOP" THEN
            response.write fullTxtLogin
          END IF
          %>
          <tr> 
            <%
            IF txtLogin<>"" AND txtLoginLoc = "LEFT" THEN
              response.write fullTxtLogin
            END IF
            %>
            <td id='loginDIVcell'> <%
  END IF
            defaultFormStart thispage, "onSubmit=""return validaEntrada()""", false %>
              <% IF modo="user" THEN %>
                <%
                  IF (canUseFacebook OR canUseLnkdIn) AND NOT socialLoginDown THEN
                    '-- Titulo excluido en drawLoginForm
                    IF NOT loginFormNoTitle THEN %>
                    <div class="headTitle" style="text-align:center;"><%=lblKHOR_Persona%></div> <%
                    END IF                   
                    paintSocialLogin
                    response.write strAdd( lblFRS_FBoLnkdInSeparadorForma, "<br/>", "" )
                    loginFormNoTitle = true '-- Para excluir el titulo en drawLoginForm
                  END IF%>
                  <style>
                    h1 {
                      color: #fff;
                      text-align: center;
                      font-weight: 300;
                    }

                    #slider {
                      position: relative;
                      overflow: hidden;
                      border-radius: 4px;
                    }

                    #slider ul {
                      position: relative;
                      margin: 0;
                      padding: 0;
                      height: 200px;
                      list-style: none;
                    }

                    #slider ul li {
                      position: relative;
                      display: block;
                      float: left;
                      margin: 0;
                      padding: 0;
                      width: 500px;
                      height: 196px;
                      text-align: center;
                    }

                    a.control_prev, a.control_next {
                      position: absolute;
                      top: 40%;
                      z-index: 999;
                      display: block;
                      padding: 4% 3%;
                      width: auto;
                      height: auto;
                      background: #2a2a2a;
                      color: #fff;
                      text-decoration: none;
                      font-weight: 600;
                      font-size: 18px;
                      opacity: 0.8;
                      cursor: pointer;
                    }

                    a.control_prev:hover, a.control_next:hover {
                      opacity: 1;
                      -webkit-transition: all 0.2s ease;
                    }

                    a.control_prev {
                      border-radius: 0 2px 2px 0;
                    }

                    a.control_next {
                      right: 0;
                      border-radius: 2px 0 0 2px;
                    }

                    .slider_option {
                      position: relative;
                      margin: 10px auto;
                      width: 160px;
                      font-size: 18px;
                    }
                  </style>
                  <script>
                    var slideCount; 
                    var slideWidth; 
                    var slideHeight;
                    var sliderUlWidth;
                    jQuery(document).ready(function ($) {
                      
                      slideCount = $('#slider ul li').length;
                      slideWidth = $('#slider ul li').width();
                      slideHeight = $('#slider ul li').height();
                      sliderUlWidth = slideCount * slideWidth;
                      
                      $('#slider').css({ width: slideWidth, height: slideHeight });
                      $('#slider ul').css({ width: sliderUlWidth, marginLeft: - slideWidth });
                      $('#slider ul li:last-child').prependTo('#slider ul');

                    });
                    function moveRight(perType) {
                      $("#perType").val(perType);
                      $('#slider ul').animate({
                        left: - slideWidth
                      }, 200, function () {
                        $('#slider ul li:first-child').appendTo('#slider ul');
                        $('#slider ul').css('left', '');
                        if(perType==1){
                          $("#loginPostulante #usr").removeAttr("disabled");
                          $("#loginPostulante #pwd").removeAttr("disabled");
                          $("#loginEmpleado #loginTipoPer").attr("disabled","disabled");
                          $("#loginEmpleado #usr").val("").attr("disabled","disabled");
                          $("#loginEmpleado #pwd").val("").attr("disabled","disabled");
                          $("#loginPostulante #usr")[0].focus();
                          $("#loginPostulante #usr")[0].select();
                        } else {
                          $("#loginEmpleado #usr").removeAttr("disabled");
                          $("#loginEmpleado #pwd").removeAttr("disabled");
                          $("#loginEmpleado #loginTipoPer").removeAttr("disabled");
                          $("#loginPostulante #usr").val("").attr("disabled","disabled");
                          $("#loginPostulante #pwd").val("").attr("disabled","disabled");
                          $("#loginEmpleado #usr")[0].focus();
                          $("#loginEmpleado #usr")[0].select();
                        }
                      });
                    }
                  </script>
                  <div id="slider" style="width:700px; margin-left:auto; margin-right:auto;">
                    <ul>
                      <li>
                  <div id="loginPostulante" style="">
                  <% drawLoginForm 0, loginSucursal, loginTipoPer, khorLoginLabel(0, khorDefaultPais()), iif(perType=1,loginUsr,""), tipoPassword, loginPwdField, lblPwd, "Candidato", false, perType=2 %>
                  <% IF modo<>"" THEN%>
                    <div align="center" style="display:<%=displayStyle(modo<>"")%>;">
                      <%inputSubmit "Submit", lblFRS_Ingresar, "whitebtn", ""%>
                    </div><%
                  IF conAutoregistro THEN %>
                    <div id="loginDIVautoregistro" class="loginDIV" style="display:inline-block;"> <%
                    response.write "Registro de Postulantes" & iif(lblFRS_SiEsPrimeraVez="","","<br/>")
                    if loginUseLinks then
                      response.write "[<a href=""#"" onClick=""registro(); return false;"">" & lblFRS_Registrarse & "</a>]"
                    else
                      inputButton "", lblFRS_Registrarse, "registro(); return false;", "whitebtn", ""
                    end if %> 
                    </div> <%
                  END IF%>
                    <div id="loginDIVcolaborador" class="loginDIV" style="display:inline-block;"> <%
                    response.write "Eres Colaborador:<br/>"
                    if loginUseLinks then
                      response.write "[<a href=""#"" onClick=""moveRight(2); return false;"">" & "Click aquí" & "</a>]"
                    else
                      inputButton "", "Click aquí", "moveRight(2); return false;", "whitebtn", ""
                    end if %> 
                    </div>
                  </div>
                      </li>
                      <li>
                  <div id="loginEmpleado" style="">
                  <% drawLoginForm tipoLogin, loginSucursal, loginTipoPer, lblUsr, iif(perType=2,loginUsr,""), tipoPassword, loginPwdField, lblPwd, "Colaborador", true, perType=1 %>
                  <% IF modo<>"" THEN%>
                    <div align="center" style="display:<%=displayStyle(modo<>"")%>;">
                      <%inputSubmit "Submit", lblFRS_Ingresar, "whitebtn", ""%>
                    </div>
                  <% END IF%>
                    <div id="loginDIVcandidato" class="loginDIV" style="display:inline-block;"> <%
                    response.write "Eres Candidato:<br/>"
                    if loginUseLinks then
                      response.write "[<a href=""#"" onClick=""moveRight(1); return false;"">" & "Click aquí" & "</a>]"
                    else
                      inputButton "", "Click aquí", "moveRight(1); return false;", "whitebtn", ""
                    end if %> 
                    </div>
                  </div>
                      </li>
                    </ul> 
                  </div><%
                  IF perType=2 THEN%>
                  <script>
                    $('#slider ul li:first-child').appendTo('#slider ul');
                  </script><%
                  END IF%>
                  <input type="hidden" id="perType" name="perType" value="<%=perType%>"/>
                  <% END IF %>
              <% ELSEIF modo="admin" THEN %>
                <table border="0" cellspacing="2" cellpadding="2" align="center"> <%
                  if NOT loginFormNoTitle then %>
                  <tr>
                    <td colspan="2" class="headTitle" style="text-align:center;"><%=tit2%></td>
                  </tr> <%
                  end if %>
                  <tr>
                    <td align="right"><%=lblUsr%>:</td>
                    <td>
                      <%inputText "usr", loginUsr, "", "whiteblur", "", "width:200px;", ""%>
                    </td>
                  </tr>
                  <tr>
                    <td align="right"><%=lblFRS_ContrasenaPassword%>:</td>
                    <td>
                      <% inputPassword "pwd", "", "", "whiteblur", "", "width:200px;", "" %>
                    </td>
                  </tr>
                </table>
                <div align="center" style="display:<%=displayStyle(modo<>"")%>;">
                  <%inputSubmit "Submit", lblFRS_Ingresar, "whitebtn", ""%>
                </div>
              <% ELSE %>
                <div class="loginDIV">
                  <%inputButton "", lblKHOR_Persona, "sendval('','modo','user');", "whitebtn", ""%>
                </div>
                <div class="loginDIV">
                  <%inputButton "", lblKHOR_Administrador, "sendval('','modo','admin');", "whitebtn", ""%>                  
                </div>
              <% END IF %>
              <% IF modo<>"" THEN %>
              <input type="hidden" name="mov" value="submit">
              <input type="hidden" name="idOferta" value="<%=IdOferta%>">
              <% if khorConfigValue(225,true)<>0 then %>
              <div class="loginDIV">
                [<a href="#" onClick="sendval('','modo','<%=iif(modo="user","admin","user")%>');"><%=iif(modo="user",lblKHOR_Administrador,lblKHOR_Persona)%></a>]
              </div>
              <% end if %>
              <% END IF %>
              <input type="hidden" name="modo" value="<%=modo%>">
              <input type="hidden" name="embedded" value="<%=bool2num(embedded)%>">
              <%
                IF modo="user" THEN
                  'Configuración "dirty": definir en khorLabelEspecial.asp o similar
                  'Para mostrar ligas en lugar de botones:
                  'loginUseLinks = true
                  'Para mostrar en diferente renglón (como era antes):
                  'loginSeparateAutoRegRecover = true
                  IF conAutoregistro OR conRecuperacionPwd THEN %>
                  <br/>
                  <table border="0" cellspacing="1" cellpadding="1" align="center">
                    <tr><%
                    IF conRecuperacionPwd THEN %>
                      <div id="loginDIVrecuperapwd" class="loginDIV"> <%
                        response.write lblFRS_SiOlvidoPwd & iif(lblFRS_SiOlvidoPwd="","","<br/>")
                        if loginUseLinks then
                          response.write "[<a href=""#"" onClick=""recuperacion(); return false;"">" & lblFRS_RecuperarPassword & "</a>]"
                        else
                          inputButton "", lblFRS_RecuperarPassword, "recuperacion(); return false;", "whitebtn", ""
                        end if %>
                      </div><%
                    END IF %>
                    </tr>
                  </table> <%
                  END IF
                  IF socialLoginDown THEN
                    paintSocialLogin
                  END IF
                  IF insAnonimaAplicable() AND idoferta=0 AND NOT enMantto THEN %>
                  <div id="loginDIVinsaplica" class="loginDIV" align="center">[<a href="#" onClick="return AplicaIns();"><%=lblIns_AplicarInstrumentoAnonimo%></a>]</div> <%
                  END IF
                  IF coboAnonimaAplicable() AND idoferta=0 AND NOT enMantto THEN %>
                  <div id="loginDIVcoboaplica" class="loginDIV" align="center">[<a href="#" onClick="return AplicaCO();"><%=lblCOBO_Ingresar%></a>]</div> <%
                  END IF
                  IF khorEncuestaAnonima() AND idoferta=0 AND NOT enMantto THEN %>
                  <div id="loginDIVcapevaplica" class="loginDIV" align="center">[<a href="#" onClick="return AplicaCapEva();"><%=lblEnc_EvalEveAnonima%></a>]</div> <%
                  END IF
                ELSE 
                  IF modo = "admin" THEN
                    IF conRecuperacionPwdAdmin THEN %>
                      <div id="loginDIVrecuperapwd" class="loginDIV"> <%
                        response.write lblFRS_SiOlvidoPwd & iif(lblFRS_SiOlvidoPwd="","","<br/>")
                        if loginUseLinks then
                          response.write "[<a href=""#"" onClick=""recuperacionAdmin(); return false;"">" & lblFRS_RecuperarPassword & "</a>]"
                        else
                          inputButton "", lblFRS_RecuperarPassword, "recuperacionAdmin(); return false;", "whitebtn", ""
                        end if %>
                      </div><%
                    END IF
                  END IF  
                END IF %>
            </FORM>
<%
IF embedded THEN
    response.write "</BODY>"
ELSE %>
            </td>
            <%
            IF txtLogin<>"" AND txtLoginLoc = "RIGHT" THEN
              response.write fullTxtLogin
            END IF
            %>
          </tr>
          <%
          IF txtLogin<>"" AND txtLoginLoc = "BOTTOM" THEN
            response.write fullTxtLogin
          END IF
          %>
        </table>
        <div id="loginDIVmsg" class="loginMSG" align="center">
          <%=lblFRS_Requirements%>
        </div> <%
        imgLogin = khorConfigValue(101,false)
        if imglogin<>"" then %>
        <div id="loginDIVimg" align=center style="padding:0px;">
          <img border="0" alt="" src="<%=imglogin%>">
        </div><%
        end if
      if not useKhorStyles() then
        contentPaintEndTag 
      end if

    layoutEnd
    removeBorder
END IF
IF modo<>"" THEN %>
<script language="JavaScript">
<!--
  <% IF modo="admin" THEN %>
  $("#usr")[0].focus();
  $("#usr")[0].select();
  <% ELSE %>
  $("#<%=iif(perType=1,"loginPostulante","loginEmpleado")%> #usr")[0].focus();
  $("#<%=iif(perType=1,"loginPostulante","loginEmpleado")%> #usr")[0].select();
  <% END IF %>
//-->
</script> <%
END IF %>