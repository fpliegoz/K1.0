<!--#include file="./frsClass.asp"-->
<!--#include file="./khorComun.asp"-->
<!--#include file="./psiSemaforoClass.asp"-->
<!--#include file="./khorObject.asp"-->
<!--#include file="./khorCompatibilidadClass.asp"-->
<!--#include file="./khorMenuClass.asp"-->
<!--#include file="./khorUsuarioClass.asp"-->
<!--#include file="./khorCatalogoClass.asp"-->
<!--#include file="./personalConfigClass.asp"-->
<!--#include file="./buscadorClass.asp"-->
<!--#include file="./dpPerfilClass.asp"-->
<!--#include file="./e360comun.asp"-->
<!--#include file="./mciComun.asp"-->
<!--#include file="./edc_Comun.asp"-->
<!--#include file="./coboComun.asp"-->
<!--#include file="./LidtrComun.asp"-->
<!--#include file="./eiComun.asp"-->
<!--#include file="./mapPotComun.asp"-->
<!--#include file="./tctComun.asp"-->
<!--#include file="./rcComun.asp"-->
<!--#include file="./extraDatoClass.asp"-->
<!--#include file="./ecdClass.asp"-->
<!--#include file="./layout.asp"-->
<%

'' ----------------------------------------------------------------------------
'' ---------- GLOBALS ---------------------------------------------------------
'' ----------------------------------------------------------------------------

CONST SECUENCIAPRUEBAEXTERNA = 600
CONST GRUPOPRUEBAEXTERNA = 6
CONST SECUENCIAPRUEBAPERSONALIZADA = 800
CONST GRUPOPRUEBAPERSONALIZADA = 8
CONST EXT_ESC_PORCENTAJE = 1
CONST EXT_ESC_NIVEL_0_6 = 2
CONST EXT_ESC_NIVDOMINIO = 3
CONST EXT_ESC_SEMAFORO = 4

cargaCatalogos 0

'===== Estado del grupo
Const E360_INCOMPLETO = 0
Const E360_ABIERTO = 1
Const E360_CERRADO = 2

function desc360statusGrupo(st)
  desc360statusGrupo=descFromMenuFijo(menufijo360statusGrupo,st)
end function

'===== Estado de evaluacion
Const E360_NOGENERADA = -1
Const E360_PENDIENTE = 0
Const E360_PARCIAL = 1
Const E360_COMPLETA = 2

function desc360statusEval(st)
  desc360statusEval=descFromMenuFijo(menufijo360statusEval,st)
end function

Function e360StatusEvaluacionRS(recset)
  if rsNum(recset,"c")=0 then
    e360StatusEvaluacionRS = E360_NOGENERADA
  elseif rsNum(recset,"c") = rsNum(recset,"s") then
    e360StatusEvaluacionRS = E360_COMPLETA
  elseif rsNum(recset,"s") = 0 then
    e360StatusEvaluacionRS = E360_PENDIENTE
  else
    e360StatusEvaluacionRS = E360_PARCIAL
  end if
end function

'' ----------------------------------------------------------------------------
'' ---------- SUCURSAL -----------------------------------------------------
'' ----------------------------------------------------------------------------
class clsSucursal
  public IdSucursal
  public Sucursal
  public Clave
  public IdOrganizacion
  public IdEmpresaCap
  public logoHeader
  public Activa
  public FechaVigencia
  public AvisoPrivacidad
  '-- Auxiliares, no estan en BD
  public Licencias
  public Contador

  public function actualizaLicencias()
    dim olic : set olic = khorLicenciasObjectCreate()
    actualizaLicencias = olic.sucursalAsignaLicencias(clng(IdSucursal), clng(Licencias))
    set olic = nothing
    logAcceso LOG_CAMBIO, "Sucursal-Licencias", "[" & IdSucursal & "] " & Licencias
  end function

  public sub leeLicencias()
    dim olic : set olic = khorLicenciasObjectCreate()
    Licencias = olic.sucursalLicenciasAsignadas( clng(IdSucursal), true )
    Contador = olic.sucursalLicenciasUsadas( clng(IdSucursal) )
    set olic = nothing
  end sub

  sub update(conn)
    dim sq ,rs
    if IdSucursal=0 then
      sq = "INSERT INTO catSucursal (Sucursal,IdOrganizacion,Clave,IdEmpresaCap,logoHeader,Activa,FechaVigencia,AvisoPrivacidad)" & _
          " VALUES ('"& sqsf(Sucursal) &"'," & IdOrganizacion & ",'" & sqsf(Clave) & "'," & IdEmpresaCap & ",'" & sqsf(logoHeader) & "'," & Activa & "," & formatDateSQL(FechaVigencia, false) & ",'"& sqsf(AvisoPrivacidad) &"');"
      conn.execute sq
      sq = "SELECT MAX(IdSucursal) AS lastid FROM catSucursal WHERE Sucursal='"& sqsf(Sucursal) &"' AND IdOrganizacion="&IdOrganizacion&";"
      IdSucursal = getBDnum("lastid",sq)
      logAcceso LOG_ALTA, "Sucursal", "[" & IdSucursal & "] " & Sucursal
    else
      sq = "UPDATE catSucursal SET Sucursal='"& sqsf(Sucursal) &"', IdOrganizacion=" & IdOrganizacion & ", Clave='" & sqsf(Clave) & "', IdEmpresaCap=" & IdEmpresaCap & ", logoHeader='" & sqsf(logoHeader) & "', Activa=" & Activa & ", FechaVigencia=" & formatDateSQL(FechaVigencia, false) & _
          ", AvisoPrivacidad = '"& sqsf(AvisoPrivacidad) &"' WHERE IdSucursal=" & IdSucursal
      conn.execute sq
      logAcceso LOG_CAMBIO, "Sucursal", "[" & IdSucursal & "] " & Sucursal
    end if
  end sub

  public sub getFromRS(rs)
    IdSucursal = rsNum(rs,"IdSucursal")
    Sucursal = rsStr(rs,"Sucursal")
    Clave = rsStr(rs,"Clave")
    IdOrganizacion = rsNum(rs,"IdOrganizacion")
    IdEmpresaCap = rsNum(rs,"IdEmpresaCap")
    logoHeader = rsStr(rs,"logoHeader")
    Activa=rsNum(rs,"Activa")
    FechaVigencia=rs("FechaVigencia")
    AvisoPrivacidad = rsStr(rs,"AvisoPrivacidad")
    if (khorConfigValue(329,true)<>0) then
      leeLicencias
    end if
  end sub

  function getfromDB(conn,id)
    dim sq, rs
    getfromDB=false
    if id<>"" then
      sq = "SELECT * FROM CatSucursal WHERE IdSucursal=" & id
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  public sub clean()
    IdSucursal = 0
    Sucursal = ""
    Clave = ""
    IdOrganizacion = 0
    IdEmpresaCap = 0
    Licencias = 0
    Contador = 0
    logoHeader = ""
    Activa=1
    FechaVigencia = null
    AvisoPrivacidad = ""
  end sub
  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

function filtroSucursalesActivas(idSucursalActual)
  dim aux
  if khorMultiSucursal() then
    aux = "(Activa=1 AND (FechaVigencia is NULL OR FechaVigencia >= " & formatDateSQL(Date,false) & "))" 
    if CStr(idSucursalActual)<>"0" and CStr(idSucursalActual)<>"" then aux = strAdd(aux, " OR " , "IdSucursal IN (" & idSucursalActual & ")")
  else
    aux = ""
  end if
  filtroSucursalesActivas = aux
end function
'' ----------------------------------------------------------------------------
'' ---------- PREFIJO -----------------------------------------------------
'' ----------------------------------------------------------------------------
class clsPrefijo
  public IdPrefijo
  public Prefijo
  public Longitud
  public Autonumerico
  public Descripcion
  public IdEmpresa
  public bEmpleado
  public strAuxiliar  ''revisar
  
  'ide: idempresa (incluye comunes si es <>0)
  'tipoNumerico: 0=indistinto, 1=capturable, 2=automatico
  'tipoEmpleado: 0=indistinto, 1=empleado, 2=no-empleado
  function wherePrefijos(ide,tipoNumerico,tipoEmpleado)
    dim retval : retval = khorSQLprefijoEmpresa(ide)
    if tipoNumerico<>0 then retval = retval & " AND Automatico=" & iif(tipoNumerico=1,0,1)
    if tipoEmpleado<>0 then retval = retval & " AND bEmpleado=" & iif(tipoEmpleado=1,1,0)
    wherePrefijos = retval
  end function

  function numPrefijosX(ide,tipoNumerico,tipoEmpleado)
    numPrefijosX = getBDnum( "cuantos", "SELECT COUNT(*) AS cuantos FROM Prefijo WHERE " & wherePrefijos(ide,tipoNumerico,tipoEmpleado) )
  end function
  function numPrefijos(ide,soloauto)  'For backward compatibility
    numPrefijos = numPrefijosX(ide,iif(soloauto,2,0),0)
  end function

  sub optionFromX(ide,tipoNumerico,tipoEmpleado)
    dim sq : sq = "SELECT IdPrefijo, (Descripcion "&db_concat&" ' (' "&db_concat&" Prefijo "&db_concat&" ')') AS Nombre" & _
                  " FROM Prefijo WHERE " & wherePrefijos(ide,tipoNumerico,tipoEmpleado) & " ORDER BY Descripcion"
    optionFromQuery "IdPrefijo", "Nombre", idprefijo, sq
  end sub
  sub optionFrom(ide,soloauto,soloemp)   'For backward compatibility
    optionFromX ide, iif(soloauto,2,0), iif(soloemp,1,0)
  end sub

  function numPrefijosTotales(soloauto) 'For backward compatibility, apparently not used anymore
    numPrefijosTotales = getBDnum( "cuantos", "SELECT COUNT(*) AS cuantos FROM Prefijo" & iif(soloauto, " WHERE Automatico=1", "" ) )
  end function
  
  function listaPrefijosParaFolioUnico()
    dim lstAux : lstAux = ""
    if khorFolioComunPrefijos() AND bEmpleado then
      lstAux = getBDlist("IdPrefijo","SELECT IdPrefijo FROM prefijo WHERE bEmpleado<>0 AND " & khorSQLprefijoEmpresa(IdEmpresa),false)
    end if
    if lstAux = "" then lstAux = IdPrefijo
    listaPrefijosParaFolioUnico = lstAux
  end function
  
  function nextFolio()
    nextFolio = getBDnum("nextfolio","SELECT (MAX(Folio)+1) AS nextfolio FROM Personal WHERE IdPrefijo IN (" & listaPrefijosParaFolioUnico() & ")")
  end function

  function folioRepetido(idpersona,folio)
    dim sq : sq = "SELECT IdPersona FROM vPersona WHERE idprefijo IN (" & listaPrefijosParaFolioUnico() & ")" & _
                  " AND folio=" & folio & " AND idpersona<>" & IdPersona
    folioRepetido = bdExists(sq)
  end function

  function formaClave(elem)
    formaClave = Prefijo & string(Longitud - len(elem),"0") & elem
  end function

  function chkUname(conn)
    dim sq, rs
    sq = "SELECT IdPrefijo FROM Prefijo WHERE (Descripcion='" & sqsf(Descripcion) & "' OR Prefijo='" & sqsf(Prefijo) & "') AND IdEmpresa=" & IdEmpresa
    set rs = getrs(conn,sq)
    if rs.EOF then
      chkUname = true
    else
      chkUname = (rs("IdPrefijo")&""=IdPrefijo&"")
    end if
    rs.close
    set rs = nothing
  end function

  sub update(conn)
    dim sq
    sq = "UPDATE Prefijo SET Prefijo='" & sqsf(prefijo) & "', Numeros=" & Longitud & ", Automatico=" & Autonumerico & ", Descripcion='" & sqsf(Descripcion) & "', bEmpleado=" & bool2num(bEmpleado) & " WHERE IdPrefijo=" & IdPrefijo
    conn.execute (sq)
    logAcceso LOG_CAMBIO, lblKHOR_TipoDePersona, Descripcion
  end sub

  sub insert(conn)
    dim sq
    sq = "INSERT INTO Prefijo (Prefijo,Numeros,Automatico,Descripcion,IdEmpresa,bEmpleado)"
    sq = sq & " VALUES ('" & sqsf(Prefijo) & "'," & Longitud & "," & Autonumerico & ",'" & sqsf(Descripcion) & "'," & IdEmpresa & "," & bool2num(bEmpleado) & ")"
    conn.execute sq
    sq = "SELECT MAX(IdPrefijo) AS lastid FROM Prefijo WHERE IdEmpresa="&IdEmpresa&" AND Prefijo='"&sqsf(Prefijo)&"'"
    IdPrefijo=getBD("lastid",sq)
    logAcceso LOG_ALTA, lblKHOR_TipoDePersona, Descripcion
  end sub

  sub delete(conn)
    conn.execute "DELETE FROM Prefijo WHERE IdPrefijo="&IdPrefijo
    logAcceso LOG_BAJA, lblKHOR_TipoDePersona, Descripcion
  end sub
  
  sub getfromRS(rs)
    IdPrefijo = rsNum(rs,"IdPrefijo")
    Prefijo = rsStr(rs,"Prefijo")
    Longitud = rsNum(rs,"Numeros")
    Autonumerico = rsNum(rs,"automatico")
    Descripcion = rsStr(rs,"Descripcion")
    IdEmpresa = rsNum(rs,"IdEmpresa")
    bEmpleado = rsBool(rs,"bEmpleado")
  end sub

  function getfromDB(conn,id)
    dim sq, rs
    getfromDB=false
    if id<>"" then
      sq = "SELECT * FROM Prefijo WHERE IdPrefijo=" & id
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  function getfromDBfilter(conn,filter)
    dim sq, rs
    getfromDBfilter=false
    if filter<>"" then
      sq = "SELECT * FROM Prefijo WHERE " & filter
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDBfilter=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  sub copyTo(reg)
    reg.IdPrefijo = IdPrefijo
    reg.Prefijo = Prefijo
    reg.Longitud = Longitud
    reg.Autonumerico = Autonumerico
    reg.Descripcion = Descripcion
    reg.IdEmpresa = IdEmpresa
    reg.bEmpleado = bEmpleado
    reg.strAuxiliar = strAuxiliar
  end sub

  sub clean
    IdPrefijo = 0
    Prefijo = "EXT"
    Longitud = 7
    Autonumerico = 1
    Descripcion = "EXT"
    IdEmpresa = 0
    bEmpleado = false
    strAuxiliar = ""
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub

end class

'' ----------------------------------------------------------------------------
'' ---------- PERSONA -----------------------------------------------------
'' ----------------------------------------------------------------------------

function copyStatusPersonaPlaza(idper,statusper)
  if PLAZA_INACTIVA = "" then
    PLAZA_INACTIVA = 0
    PLAZA_ACTIVA = 1
  end if
  if statusper=1 then
    '== Inactivo: inactiva la plaza con la fecha de baja actual
    conn.Execute "UPDATE Plaza SET Status=" & PLAZA_INACTIVA & ", FechaBaja=" & formatDateSQL(Now,true) & " WHERE IdPersona = " & idper & " AND Status=" & PLAZA_ACTIVA
    copyStatusPersonaPlaza = true
  elseif not bdexists("SELECT * FROM Plaza WHERE IdPersona=" & idper & " AND Status=" & PLAZA_ACTIVA ) then
    '== Activo: activa la plaza si esta inactiva
    conn.Execute "UPDATE Plaza SET Status=" & PLAZA_ACTIVA & " WHERE IdPersona = " & idper & " AND Status=" & PLAZA_INACTIVA
    copyStatusPersonaPlaza = false
  end if
end function

function setPersonaUltimaActualizacion(idper)
  conn.execute "UPDATE Personal SET UltimaActualizacion=" & formatDateSQL(Now,true) & " WHERE IdPersonal=" & idper
end function

'' ----------------------------------------------------------------------------

class PersonalTokenPwdReset
  public IdPersonal 
  public email
  public FechaVigencia
  public Usado  '0=no usado, 1=usado, 2=invalido
  public Token	

  '-- Crea, registra, y regresa un token de password-reset para la pesona con <mail_address>, solo si ese correo es único
  function create(conn,mail_address)
    clean
    IdPersonal = getBDNum("IdPersonal", "SELECT CASE WHEN COUNT(DISTINCT IdPersonal) = 1 THEN IdPersonal ELSE 0 END IdPersonal FROM Personal WHERE email = " & nsqsf(mail_address) & " GROUP BY IdPersonal")
    if IdPersonal <> 0 then
      FechaVigencia = DateAdd("n",30,now)
      Token =  GetGuid()
      Token = replace(replace(replace(Token,"{",""),"}",""),"v","")
      Token = replace(Token,"-","")
      email = mail_address
      conn.execute "UPDATE PersonalTokenPwdReset SET Usado = 2 WHERE Usado = 0 AND email = " & nsqsf(email)
      conn.execute "INSERT INTO PersonalTokenPwdReset (IdPersonal, FechaVigencia, Usado, Token, email) VALUES (" & IdPersonal & ", " & formatDateSQL(FechaVigencia,true) & ", 0, " & nsqsf(Token) & ", " & nsqsf(email) & ");"
    end if
    create = Token
  end function

  '-- valida input_token, y regresa el idpersonal asociado, o cero y mensaje de error en errmsg
  function validate(conn,input_token,errmsg)
    dim retval : retval = 0
    if getfromDB( conn, input_token ) then
      if token = input_token and Usado = 0 and FechaVigencia >= now then
        retval = IdPersonal
        errmsg = ""
      else
        if Usado = 1 then
          errmsg = lblFRS_TokenUsed
        elseif token <> token or Usado = 2 then
          errmsg = lblFRS_TokenInvalid
        elseif FechaVigencia < now then
          errmsg = lblFRS_TokenExpired
        end if
      end if
    end if
    validate = retval
  end function

  sub marcaUsado(conn)
    conn.execute "UPDATE PersonalTokenPwdReset SET Usado=1 WHERE Token = " & nsqsf(Token)
  end sub

  sub delete(conn)
    conn.execute "DELETE FROM PersonalTokenPwdReset WHERE Token = " & nsqsf(Token)
  end sub
  
  sub getfromRS(rs)
    IdPersonal = clng(rsNum(rs,"IdPersonal"))
    Token = rsStr(rs,"Token")
    email = rsStr(rs,"email")
    Usado = cint(rsNum(rs,"Usado"))
    FechaVigencia = rs("FechaVigencia")
  end sub

  function getfromDB(conn,Token)
    dim sq, rs
    getfromDB = false
    if Token <> "" then
      sq = "SELECT IdPersonal, FechaVigencia, Usado, Token, email FROM PersonalTokenPwdReset WHERE Token = " & nsqsf(Token)
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB = true
      end if
      rs.close
      set rs = nothing
    end if
  end function
  
  sub clean()
    IdPersonal = 0
    FechaVigencia = cdate("1/1/1900")
    Usado = 0
    Token = ""
    email = ""
  end sub
  
  private sub class_initialize()
    clean
  end sub  
  private sub class_terminate()
  end sub  
end class


class clsPersona
  public IdPersona
  public Nombre
  public Clave
  public Password
  public Email
  public IdPrefijo
  public Folio
  public IdEmpresa
  public StatusPer
  public FechaRegistro
  public Bloqueado

  public function conCompetencias()
    dim sq, rs, ok
    sq = "SELECT DISTINCT idprueba FROM MatrizCompetencias"
    set rs = getrs(conn,sq)
    ok = true
    while ok and not rs.eof
      ok = getBD("cuantos","SELECT COUNT(*) AS cuantos FROM PersonaPruebaH WHERE IdPersonal="&IdPersona&" AND IdPrueba="&rsNum(rs,"idprueba"))>0
      rs.movenext
    wend
    rs.close
    set rs = nothing
    conCompetencias=ok
  end function

  public property get esEmpleado()
    esEmpleado = (getBDnum("esEmpleado","SELECT (CASE WHEN bEmpleado<>0 THEN 1 ELSE 0 END) AS esEmpleado FROM Prefijo WHERE IdPrefijo="&IdPrefijo)<>0)
  end property

  public property get IdSucursal()
    IdSucursal=IdEmpresa
  end property

  public property get IdPerfil()
    IdPerfil=khorGetPerfilPersona(IdPersona)
  end property

  sub setPassword(conn,newpwd)
    dim sql
    sql = "UPDATE Personal SET Password='"&sqsf(newpwd)&"', FechaUltimoCambioPwd=" & formatDateSQL(Now,true) & " WHERE IdPersonal="&IdPersona&";"
    conn.execute sql
    if cInt(khorNumHistoricoPwd())>0 then
      sql = "INSERT INTO PersonalPassword(IdPersonal,Fecha,Password) VALUES("&IdPersona&","&formatDateSQL(Now,true)&",'"&sqsf(newpwd)&"')"
      conn.execute sql
    end if
    logAcceso LOG_CAMBIO, lblKHOR_Persona&":"&lblFRS_CambioDePassword, Clave
    Password=newpwd
  end sub

  function setStatus(status)
    if (StatusPer <> status) then
      dim sq : sq = "UPDATE Personal SET StatusPer="&status
      if copyStatusPersonaPlaza(IdPersona,status) then
        sq = sq & ", FechaBaja=" & formatDateSQL(Now,true)
      end if
      sq = sq & " WHERE IdPersonal="&IdPersona
      conn.execute (sq)
      StatusPer = status
      setStatus = true
    else
      setStatus = false
    end if
  end function

  sub fixFolio(regpre)
    dim lstAux : lstAux = regpre.listaPrefijosParaFolioUnico()
    if regpre.autonumerico<>0 then
      do
        conn.execute "UPDATE Personal SET Folio=(SELECT (MAX(Folio)+1) AS Folio FROM Personal WHERE IdPrefijo IN (" & lstAux & ")) WHERE IdPersonal=" & IdPersona
        Folio = getBDnum("Folio","SELECT Folio FROM Personal WHERE IdPersonal=" & IdPersona)
      loop while bdExists("SELECT IdPersonal FROM Personal WHERE IdPrefijo IN (" & lstAux & ") AND Folio=" & Folio & " AND IdPersonal <> " & IdPersona)
    end if
    conn.execute "UPDATE Personal SET Clave='" & regpre.formaClave(Folio) & "' WHERE IdPersonal=" & IdPersona
  end sub
  
  function setClave(regpre,fol,idemp)
    dim sq, auxPue, ok, autofolio
    dim oldClave : oldClave = Clave
    dim oldEmpresa : oldEmpresa = IdEmpresa
    setClave = false
    '-- Verifica que el puesto sea de la empresa o general compartido
    sq="SELECT Personal.Puesto FROM Personal INNER JOIN Puestos ON Personal.Puesto=Puestos.IdPuesto WHERE IdPersonal="&IdPersona&" AND "&khorSQLperfilesEmpresa(idemp,0)
    auxPue = getBDnum("puesto",sq)
    '-- Verificacion de folio
    if (regpre.idprefijo<>IdPrefijo) OR (fol<>Folio) then
      if (regpre.autonumerico<>0) then
        ok = true
        autofolio = true
        if fol > 0 then
          if regpre.folioRepetido(idpersona,fol) then
            ok = false
          else
            autofolio = false
            Folio = fol '-- mas adelante se regresa
          end if
        end if
        if ok then
          if autofolio then
            fol = regpre.nextFolio()
          else
            fol = Folio
          end if
          auxCla = regpre.formaClave(fol)
        end if
      elseif fol > 0 then
        auxCla = regpre.formaClave(fol)
        ok = not regpre.folioRepetido(idpersona,fol)
      else
        ok = false
      end if
    else
      auxCla = Clave
      fol = Folio
      ok = true
    end if
    if ok then
      sq = "UPDATE Personal SET Clave='" & sqsf(auxCla) & "', Prefijo='" & sqsf(regpre.Prefijo) & "'," & _
          " Folio=" & fol & ", IdPrefijo=" & regpre.IdPrefijo & _
          ", IdSucursal=" & idemp & ", Puesto=" & auxPue & ", UltimaActualizacion=" & formatDateSQL(Now,true) & _
          " WHERE IdPersonal=" & IdPersona
      conn.execute (sq)
      if (regpre.autonumerico<>0) then '-- double check
        if regpre.folioRepetido(idpersona,fol) then fixFolio regpre
      end if
      setClave = getfromDB( conn, IdPersona )
      logAcceso LOG_CAMBIO, lblKHOR_Persona&":"&lblKHOR_DatosDeRegistro, Clave & " ("&IdPersona&") ["&oldEmpresa&":"&oldClave&"]"
      onPersonaSetClave IdPersona, oldClave, oldEmpresa
    end if
  end function

  function create(ap,am,nom,idpre,fol,login,pwd,idpue,idemp,pais)
    dim regpre, sq, auxCla, auxNom, ok, tipoLogin
    tipoLogin=khorTipoLogin()
    create = false
    set regpre = new clsPrefijo
    if regpre.getfromDB(conn, idpre) then
      if regpre.autonumerico<>0 then
        fol = regpre.nextFolio()
        ok = true
      else
        ok = (fol>0) and not regpre.folioRepetido(0,fol)
      end if
      auxCla = regpre.formaClave(fol)
      if ok then
        auxNom = left(strAdd(strAdd(ap," ",am)," ",nom),100)
        sq = "INSERT INTO Personal (ApellidoPaterno, ApellidoMaterno, Nombres, Nombre, Password" & _
                                  ", Clave, Prefijo, Folio, IdPrefijo" & _
                                  ", Puesto,  IdSucursal, IdPais, Sexo, EdoCivil" & _
                                  ", FechaRegistro, UltimaActualizacion"
        if tipoLogin<>1 then sq = sq & ", " & khorLoginField(tipoLogin)
        sq = sq & ") VALUES ('" & sqsf(ap) & "', '" & sqsf(am) & "', '" & sqsf(nom) & "', '" & sqsf(auxNom) & "', '" & sqsf(pwd) & "'" & _
                            ", '" & sqsf(auxCla) & "', '" & sqsf(regpre.Prefijo) & "', " & fol & ", " & regpre.IdPrefijo & _
                            ", " & idpue & ", " & idemp & ", " & pais & ", null, null"& _
                            ", " & formatDateSQL(Now,true) & ", " & formatDateSQL(Now,true)
        if tipoLogin<>1 then sq = sq & ", '" & sqsf(login) & "'"
        sq = sq & ")"
        conn.execute (sq)
        IdPersona = getBDnum("newId","SELECT MAX(IdPersonal) AS newId FROM Personal WHERE IdSucursal="&idemp&" AND IdPrefijo="&regpre.idPrefijo&" AND Nombre='" & sqsf(auxNom) & "'")
        if (regpre.autonumerico<>0) then '-- double check
          if regpre.folioRepetido(idpersona,fol) then fixFolio regpre
        end if
        create = getfromDB( conn, IdPersona )
        logAcceso LOG_ALTA, lblKHOR_Persona, Clave & " ("&IdPersona&")"
        onPersonaCreate IdPersona
      end if
    end if
    set regpre = nothing
  end function

  sub borraArchivosPersona()
    dim cs, co
    cs = Application("conn_string_files")
    if (khorConFoto() OR modoDocumentos()<>0) AND cs <> "" then
      Set co = Server.CreateObject("ADODB.Connection")
      co.Open cs, Application("conn_username"), Application("conn_password")
      co.execute "DELETE FROM UserFiles WHERE IdUser=" & IdPersona & ";"
      co.close
      set co = nothing
    end if
  end sub

  sub deleteDirecto(conn)
    dim sq
    dim objAux
    set objAux=new capPersonaCurso
    khorBorraMasivo conn, objAux, "SELECT * FROM capPersonaCurso WHERE IdPersona=" & IdPersona & ";"
    set objAux=nothing
    set objAux=new EvalPersona
    khorBorraMasivo conn, objAux, "SELECT * FROM evalPersona WHERE IdPersona=" & IdPersona & ";"
    set objAux=nothing
    set objAux=new pbtOfertaPersona
    khorBorraMasivo conn, objAux, "SELECT * FROM pbtOfertaPersona WHERE IdPersona=" & IdPersona & ";"
    set objAux=nothing
    set objAux=new ecEntrevista
    khorBorraMasivo conn, objAux, "SELECT * FROM ecEntrevista WHERE IdPersona=" & IdPersona & ";"
    set objAux=nothing
    khorBorraCompatibilidadPersona conn, idpersona
    conn.execute "DELETE FROM BMB1_PersonalDatos WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM BMB1_PersonalResultados WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM BMB1_ResVendedor WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM BMB2_PersonalResultados WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM co_Respuesta WHERE IdParticipante IN (SELECT IdParticipante FROM co_Participante WHERE IdPersonal=" & IdPersona & ");"
    conn.execute "DELETE FROM co_Participante WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM ConsultaPersona WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM Escolaridad WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM Idiomas WHERE IdPersonal=" & IdPersona & ";"
    '-- LMS Stuff
    conn.execute "DELETE FROM blgArticle WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    conn.execute "DELETE FROM blgComment WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    conn.execute "DELETE FROM chatMessage WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    conn.execute "DELETE FROM chatSession WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    conn.execute "DELETE FROM frmLastVisit WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM frmPost WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    conn.execute "DELETE FROM LMSItemAccess WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM LMSGroupStudent WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM LMSPersonaHomework WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PrivateMessage WHERE IdEntidadFrom=" & IdPersona & " AND TipoFrom=0;"
    conn.execute "DELETE FROM PrivateMessage WHERE IdEntidadTo=" & IdPersona & " AND TipoTo=0;"
    conn.execute "DELETE FROM ScormSCOTrack WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM WikiPage WHERE IdEntidad=" & IdPersona & " AND Tipo=0;"
    '-- LMS Stuff
    conn.execute "DELETE FROM PersonalExperiencia WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM pcsAccionCompetencia WHERE IdAccion IN (SELECT IdAccion FROM pcsAccion WHERE IdPersonal=" & IdPersona & ");"
    conn.execute "DELETE FROM pcsAccion WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM pcsDetalle WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM pcsPersona WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalAreaInteres WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalCompetencias WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalCV WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalExpFuncional WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalExtras WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalExtraTipoData WHERE IdPersonalExtraTipo IN (SELECT IdPersonalExtraTipo FROM PersonalExtraTipo WHERE IdPersonal=" & IdPersona & ");"
    conn.execute "DELETE FROM PersonalExtraTipo WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalTelefono WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalFamilia WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalHabilidad WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalHH WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalHijos WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalHO WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalIdioma WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalPassword WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalPruebaHO WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalPruebaPermiso WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalPruebasNotas WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonaNotas WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonaPruebaH WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "UPDATE Plaza SET IdPersona=0 WHERE IdPersona=" & IdPersona & ";"
    conn.execute "UPDATE PlazaH SET IdPersona=0 WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM PlazaP WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM res16FP WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resCLEAVER WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resEQ WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resGordon WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resHERRMANN WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM ResINTRAC WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resKostick WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resLidSit WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resLidTr WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resLIFO WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resLuscher WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resOffice WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resOrtografia WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM resPPV WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM respuestasPruebas WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resTHERMAN WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM ResultadosIngles WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM resVALORES WHERE IdPersona=" & IdPersona & ";"
    deleteReferences  'Referencias Laborales
    conn.execute "DELETE FROM ListaPersona WHERE IdPersona=" & IdPersona & ";"
    conn.execute "DELETE FROM PersonalTokenPwdReset WHERE IdPersonal=" & IdPersona & ";"
    conn.execute "DELETE FROM Personal WHERE IdPersonal=" & IdPersona & ";"
    logAcceso LOG_BAJA, lblKHOR_Persona, Nombre & " (" & Clave & " [" & IdPersona & "])"
    borraArchivosPersona
  end sub

  function delete(conn)
    dim ok2del, tipobaja
    tipoBaja = khorExpedientesEliminables()
    ok2del = (tipoBaja<>1)
    if tipoBaja=2 then
      ok2del = not bdExists( "SELECT * FROM PersonaPruebaH WHERE IdPersonal=" & IdPersona ) _
              and not bdExists( "SELECT * FROM EvalPersona WHERE IdPersona=" & IdPersona )
    end if
    if ok2del then
      deleteDirecto conn
    end if
    delete = ok2del
  end function

  sub deleteReferences()
    dim rs
    set rs = getrs(conn,"SELECT IdPersonalReferencia FROM PersonalReferencia WHERE IdPersonal = " & IdPersona)
    while not rs.eof
      conn.execute "DELETE FROM pbtPersonalReferenciasExtra WHERE IdPersonalReferencia=" & rsNum(rs,"IdPersonalReferencia") & ";"
    wend
    conn.execute "DELETE FROM PersonalReferencia WHERE IdPersonal = " & IdPersona & ";"
  end sub

  sub getfromRS(rs)
    IdPersona = rsNum(rs,"IdPersona")
    Nombre = rsStr(rs,"Nombre")
    Clave = rsStr(rs,"Clave")
    Password = rsStr(rs,"Password")
    Email = rsStr(rs,"Email")
    IdPrefijo = rsNum(rs,"IdPrefijo")
    Folio = rsNum(rs,"Folio")
    IdEmpresa = rsNum(rs,"IdEmpresa")
    StatusPer = rsNum(rs,"StatusPer")
    FechaRegistro = rs("FechaRegistro")
    Bloqueado = rsNum(rs,"Bloqueado")
  end sub

  sub getfromRSP(rs)
    IdPersona = rsNum(rs,"IdPersonal")
    Nombre = rsStr(rs,"Nombre")
    Clave = rsStr(rs,"Clave")
    Password = rsStr(rs,"Password")
    Email = rsStr(rs,"Email")
    IdPrefijo = rsNum(rs,"IdPrefijo")
    Folio = rsNum(rs,"Folio")
    IdEmpresa = rsNum(rs,"IdSucursal")
    StatusPer = rsNum(rs,"StatusPer")
    FechaRegistro = rs("FechaRegistro")
    Bloqueado = rsNum(rs,"Bloqueado")
  end sub

  function rsFields(tabla)
    dim retval
    retval = "@IdPersona,@Nombre,@Clave,@Password,@Email,@StatusPer,@FechaRegistro,@IdPrefijo,@Folio,@IdSucursal,@Bloqueado"
    dim auxtab
    auxtab = iif(tabla<>"",tabla&".","")
    rsFields = replace(retval,"@",auxtab)
  end function

  function rspFields(tabla)
    dim retval
    retval = "@IdPersonal,@Nombre,@Clave,@Password,@Email,@StatusPer,@FechaRegistro,@IdPrefijo,@Folio,@IdSucursal,@Bloqueado"
    dim auxtab
    auxtab = iif(tabla<>"",tabla&".","")
    rspFields = replace(retval,"@",auxtab)
  end function

  function getfromDB(conn,id)
    dim sq, rs
    getfromDB=false
    if id<>"" then
      sq = "SELECT * FROM vPersona WHERE IdPersona=" & id
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  function getfromDBfilter(conn,filter)
    dim sq, rs
    getfromDBfilter=false
    if filter<>"" then
      sq = "SELECT * FROM vPersona WHERE " & filter
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDBfilter=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  function getByKeyOrName(conn,data,cond)
    dim retval : retval = false
    dim usefolio : usefolio = false
    dim tipoacceso, filter
    dim auxp, auxf, auxfil
    dim sq, rs
    if trim(data)<>"" then
      tipoacceso = khorTipoLogin()
      if tipoacceso=1 then '-clave
        splitClave data, auxp, auxf
        auxfil = iif( auxp<>"", "Prefijo LIKE '" & sqsf(trim(auxp)) & "%'", "" )
        filter = "(" & strAdd( auxfil, " AND ", "Folio=" & getVal(auxf,true) ) & ")"
      else
        filter = khorLoginField(tipoacceso) & "='" & sqsf(trim(data)) & "'"
        usefolio = (cstr(data) = cstr(stripCharsNotInBag(data, "0123456789")))  '-- El dato es un numero
      end if
      filter = "(" & filter & " OR Nombre='" & sqsf(trim(data)) & "')"
      filter = strAdd( filter, " AND ", cond )
      sq = "SELECT * FROM Personal WHERE " & filter
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRSP(rs)
        retval = true
      end if
      rs.close
      set rs = nothing
      if usefolio and not retval then
        '-- Busca por folio entre los empleados. why? because it's common practice
        sq = "SELECT * FROM Personal WHERE IdPrefijo IN (SELECT IdPrefijo FROM Prefijo WHERE bEmpleado<>0) AND Folio=" & getVal(data,true)
        auxfil = getBDlist("IdPersonal", strAdd(sq, " AND ", cond), false)  '-- Una lista, por si hay mas de uno
        if auxfil<>"" AND lenCSV(auxfil) = 1 then
          retval = getfromDB(conn,auxfil)
        end if
      end if
    end if
    getByKeyOrName = retval
  end function

  sub clean
    IdPersona = 0
    Nombre = ""
    Clave = ""
    Password = ""
    Email = ""
    IdPrefijo = ""
    Folio = ""
    IdEmpresa = 0
    StatusPer = 0
    FechaRegistro = null
    Bloqueado = 0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub

end class

'' ----------------------------------------------------------------------------
'' ---------- PERSONA PRUEBA ----------------------------------------------
'' ----------------------------------------------------------------------------
class clsPersonaPrueba
  public IdPersona
  public IdPrueba
  public Fecha
  public Resultado
  public Duracion
  public IdPuesto 'NO EN BD, solo para calculos

  sub getfromRS(rs)
    IdPersona = rsNum(rs,"IdPersonal")
    IdPrueba = rsNum(rs,"IdPrueba")
    Fecha = rs("Fecha")
    Resultado = rsStr(rs,"Resultados")
    Duracion = rsNum(rs,"Duracion")
  end sub

  function getfromDB(conn,idper,idpru)
    dim sq, rs
    getfromDB=false
    if idper<>"" and idpru<>"" then
      sq = "SELECT * FROM PersonaPruebaH WHERE IdPersonal=" & idper & " AND IdPrueba=" & idpru & " ORDER BY Fecha DESC"
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  function getfromDBfilter(conn,filtro)
    dim sq, rs
    getfromDBfilter=false
    if filtro<>"" then
      sq = "SELECT * FROM PersonaPruebaH WHERE " & filtro & " ORDER BY Fecha DESC"
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDBfilter=true
      end if
      rs.close
      set rs = nothing
    end if
  end function
  
  function interpretacionRangos(resultado,divisionLimits,divisionLabels)
    '-- logica extraida de gauge.ashx
    dim valuesPerLimit : valuesPerLimit = split(divisionLimits,",")
    dim labelPerLimit : labelPerLimit = split(divisionLabels,",")
    dim topValue : topValue = cdbl(valuesPerLimit(ubound(valuesPerLimit)))
    dim idx : idx = -1
    if resultado <> "" then
      dim score : score = cdbl(resultado)
      if score >= topValue then
        idx = ubound(valuesPerLimit)
      else
        dim i
        for i=0 to ubound(valuesPerLimit)
          if score <= cdbl(valuesPerLimit(i)) then
            idx = i
            exit for
          end if
        next
      end if
    end if
    if idx >= 0 then
      interpretacionRangos = labelPerLimit(idx)
    else
      interpretacionRangos = ""
    end if
  end function

  function interpretacion(tipo,texto,corta)
    dim o, retval
    dim auxres, aux, i, tam, c, n, sq, rs
    dim tempRes, tmpInt, tmpAux
    dim tstyle : tstyle = "khorInt_Titulo1"
    if esPruebaExterna(idprueba) then
      retval = "<p class=""khorInt_Normal"">" & _
        getBD("Interpretacion","SELECT Interpretacion FROM PersonaPruebaHint WHERE IdPersonal=" & IdPersona & " AND IdPrueba=" & IdPrueba & " AND Fecha=" & formatDateSQL(Fecha,true)) & _
        "</p>"
    elseif esPruebaPersonalizada(idprueba) then
      sq = "SELECT IdFactor FROM PruebaCustomFactor WHERE IdPrueba = " & IdPrueba & " ORDER BY IdFactor"
      dim factNum : factNum = 0
      set rs = getrs(conn, sq)
      while not rs.eof
        tempRes = cint(mid(Resultado, (factNum * 3) + 1, 3))
        tmpInt = getBD("Interpretacion", "SELECT Interpretacion FROM PruebaInterpretacion WHERE IdPruebaFactor = " & rsNum(rs,"IdFactor") & " AND ValorMinimo <= " & tempRes & " AND ValorMaximo >= " & tempRes)
        if tmpInt <> "" then
          retval = retval & "<p class=""khorInt_Titulo1"">" & getDescripcion("PruebasFactores", "IdPruebaFactor", "PruebaFactor", rsNum(rs,"IdFactor")) & "</p>"
          retval = retval & "<p class=""khorInt_Normal"">" & tmpInt & "</p>"
        end if
        factNum = factNum + 1
        rs.movenext
      wend
      set rs = nothing
    elseif inCSV("46,47,48,49,50,51,68,72,73,74,75,76,77,78,79,81", "" & idprueba) <> -1 then 'en bd
      dim useRomanos : useRomanos = inCSV("46,47,68,81", "" & idprueba) <> -1
      dim discardInterpretation : discardInterpretation = false
      if (idprueba=48 or idprueba=49) AND NOT corta then
        retval = retval & "<div class=""khorInt_Titulo1"">" & lblGFla_Valores_Organizacionales & "</div>"
        tstyle = "khorInt_Titulo2"
        if p48useIndiceHonestidad() then
          discardInterpretation = p48tendenciaConstante(idprueba,resultado)
          if discardInterpretation then
            retval = retval & "<p class=""khorInt_Normal"">" & lbl_pp48_InterpretacionMismaTendencia & "</p>"
          end if
        end if
      end if
      if not discardInterpretation then
      idxFactor = 0
      sq = "SELECT IdPruebaFactor, PruebaFactor, posPersona FROM PruebasFactores"
      if idprueba=68 then
        sq = sq & " INNER JOIN LidTrFactor ON PruebasFactores.IdPruebaFactor = (6800 + LidTrFactor.IdFactor)"
      end if
      sq = sq & " WHERE IdPrueba = " & idprueba & " AND posPersona>0 AND IdPruebaFactor IN (SELECT IdPruebaFactor FROM PruebaInterpretacion)"
      if (idprueba = 48 OR idprueba = 49) then  '-- IVO
        '-- En interpretacion corta usa unicamente interpretaciones de los indices de valores y de antivalores
        '-- En la normal NO los usa
        sq = sq & " AND IdPruebaFactor " & iif(corta,""," NOT ") & " IN (4819,4820,4909,4910)"
      end if
      sq = sq & " ORDER BY " & iif( idprueba=48, "CampoVistaPersona", iif(idprueba=68,"ltfSecuencia, ","") & "IdPruebaFactor" )
      set rs = getrs(conn, sq)
      while not rs.eof
        idxFactor = idxFactor + 1
        tempRes = cint(mid(Resultado, rsNum(rs, "posPersona"), 3))
        if idprueba=68 then tempRes = round(tempRes / 100)
        tmpInt = getBD("Interpretacion", "SELECT Interpretacion FROM PruebaInterpretacion WHERE IdPruebaFactor = " & rsNum(rs,"IdPruebaFactor") & " AND ValorMinimo <= " & tempRes & " AND ValorMaximo >= " & tempRes)
        if (idprueba = 48 OR idprueba = 49) AND p48useIndiceHonestidad() AND NOT corta then
          tmpAux = cint(mid(Resultado, rsNum(rs, "posPersona") + iif(idprueba = 48,63,33), 3))
          if tmpAux<> 0 then tmpInt = "<i>" & lbl_pp48_InterpretacionNoCongruente & "</i>"
        end if
        if tmpInt <> "" then
          if (idprueba=48 and idxFactor=12) or (idprueba=49 and idxFactor=5) then
            retval = retval & "<div class=""khorInt_Titulo1"">" & lblGFla_Antivalores_Organizacionales & "</div>"
          end if
          retval = retval & "<p class=""" & tstyle & """>" & iif(useRomanos,numRomano(idxFactor) & ". ","") & rsStr(rs, "PruebaFactor") & "</p>"
          retval = retval & "<p class=""khorInt_Normal"">" & tmpInt & "</p>"
        end if
        rs.movenext
      wend
      rs.close
      set rs = nothing
      end if
    elseif idprueba=108 AND NOT corta then  'Ingles v2
      auxres = ""
      c = getBD( "Respuesta", "SELECT Respuesta FROM respuestasPruebas WHERE IdPrueba=108 AND IdSubPrueba=4 AND IdPersonal=" & IdPersona )
      if c<>"" then
        auxres = auxres & "<div class=""khorInt_Titulo1"">" & p108_ensayoEscritoPorEvaluado & "</div>" & _
                "<div class=""khorInt_Normal"">" & c & "</div>"
        c = getBD( "Respuesta", "SELECT Respuesta FROM respuestasPruebas WHERE IdPrueba=108 AND IdSubPrueba=-1 AND IdPersonal=" & IdPersona )
        if c<>"" then
          auxres = auxres & "<br><div class=""khorInt_Titulo1"">" & p108_comentariosRelativosAlEnsayo & "</div>" & _
                  "<div class=""khorInt_Normal"">" & c & "</div>"
        end if
      end if
      retval = auxres
    'Pruebas Escolar
    elseif idprueba=38 then 'CPS
      if not corta then retval = formatPrueba38(Resultado, false)
    elseif idprueba=39 then 'D2
      if not corta then retval = formatPrueba39(Resultado, false)
    elseif idprueba=40 then 'EDAH
      if not corta then retval = formatPrueba40(Resultado, false)
    elseif idprueba=41 then 'ENI
      if not corta then retval = formatPrueba41Int(Resultado, false)
    elseif idprueba=42 then 'strOOp
      if not corta then retval = formatPrueba42(Resultado, false)
    elseif idprueba=43 then 'WAIS-III
      if not corta then retval = formatPrueba43Int(Resultado)
    elseif idprueba=44 then 'WISC-IV
      if not corta then retval = formatPrueba44Int(Resultado)
    elseif idprueba=45 then 'Zoo
      if not corta then retval = formatPrueba45Int(Resultado, IdPersona)
    'Pruebas de habilidades
    elseif idprueba=8 then 'ingles
      retval = interpretacionRangos(Resultado,lblGFla_DivisionesIngles,lblGFla_EtiquetasDivisionesIngles)
    elseif idprueba=11 then 'ortografia
      retval = interpretacionRangos(mid(Resultado,1,3),lblGFla_DivisionesOrtografia,lblGFla_EtiquetasDivisionesOrtografia)
    elseif inCSV("21,23,24,28", "" & idprueba) <> -1 then  'Velocidad y exactitud,Razonamiento Verbal,Habilidad Numerica,Raven
      retval = interpretacionRangos(mid(Resultado,1,3),"1,2,3,4,5",lblGFla_Etiquetas_Niv5_Escala_Big)
    elseif inCSV("30,33", "" & idprueba) <> -1 then 'Aritmetica, Inspección
      retval = interpretacionRangos(mid(Resultado,1,3),"14,28,43,57,71,86,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=34 then ' OTIS
      retval = interpretacionRangos(mid(Resultado,1,3),"1,2,3,4,5",lblGFla_Etiquetas_Niv5_Escala_Big_2)
    elseif idprueba=26 then ' Comprensión Mecánica
      retval = interpretacionRangos(mid(Resultado,1,3),"11,16,26,31,36,51,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=27 then 'Comprensión y Discernimiento
      retval = interpretacionRangos(mid(Resultado,1,3),"38,46,55,63,71,80,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=29 then 'Expresión Idiomática
      retval = interpretacionRangos(mid(Resultado,1,3),"41,46,53,58,64,70,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=31 then 'Componentes
      retval = interpretacionRangos(mid(Resultado,1,3),"39,47,54,64,77,87,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=32 then 'Ensambles
      retval = interpretacionRangos(mid(Resultado,1,3),"24,39,49,59,69,79,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=35 then 'Tablas
      retval = interpretacionRangos(mid(Resultado,1,3),"34,42,49,55,59,66,100",lblGFla_Etiquetas_Niv7_Escala)
    elseif idprueba=69 then 'MCI
      set o = new clsCompetenciaNivel
      tam = len(mid(Resultado,13)) / 12
      auxres = ""
      if IdPuesto <> 0 and khorMCIPruebaCompleta() then
        listaCompetenciasPerfil = getBDList("IdCompetencia", "SELECT IdCompetencia FROM PuestoCompetencia WHERE IdPuesto = " & IdPuesto, false)
      else
        listaCompetenciasPerfil = getBDList("IdCompetencia", "SELECT DISTINCT(IdCompetencia) FROM mciReactivo", false)
      end if
      for i=1 to tam
        aux = 13 + (i-1)*12
        c = mid(Resultado,aux,6)
        n = round(mid(Resultado,aux+6,3)/100)
        if inCSV(listaCompetenciasPerfil, cint(c)) <> -1 then
          tot = tot + 1
          if corta then
            auxres = auxres & "<div class=""khorInt_Titulo1"">" & tot & ". " & getDescripcion("CatCompetencias360","IdCompetencia","Competencia",c) & ": " & lblFRS_Nivel & " " & n & "</div>"
          else
            auxres = auxres & "<div class=""khorInt_Titulo1"">" & tot & ". " & getDescripcion("CatCompetencias360","IdCompetencia","Competencia",c) & "</div>"
            if o.getfromDB( conn, c, n, 0 ) then
              auxres = auxres & "<div class=""khorInt_Normal""><B>" & lblFRS_Nivel & " " & n & ": " & htmlFormat(o.Titulo) & "</B><BR>" & htmlFormat(o.Descripcion) & "</div>"
            end if
            auxres = auxres & "<br>"
          end if
        end if
      next
      set o = nothing
      retval = auxres
    else
      if idprueba=7 then  'Comportamiento
        auxres=mid(cstr(resultado),(tipo-1)*12+4,12)
      elseif idprueba=60 then 'Eneagrama
        if resultado<>"" then auxres = p60tiposDominantes(resultado) 'shouldn't be..? & p60triadaDominante(resultado)
      else
        auxres=resultado
      end if
      retval = khorInterpretacionPrueba(idprueba,auxres,true,texto,corta)
    end if
    retval = replace( retval, "class=""khorInt_Titulo1""", "class=""khorInt_Titulo1"" style=""page-break-inside:avoid;""" )
    retval = replace( retval, "class=""khorInt_Titulo2""", "class=""khorInt_Titulo2"" style=""page-break-inside:avoid;""" )
    retval = replace( retval, "class=""khorInt_Normal""", "class=""khorInt_Normal"" style=""page-break-inside:avoid;""" )
    interpretacion = retval
  end function

  sub clean()
    IdPersona = 0
    IdPrueba = 0
    Fecha = Now
    Resultado = ""
    Duracion = 0
    IdPuesto = 0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub

end class

'' ----------------------------------------------------------------------------
'' ---------- PRUEBA PREFIX ---------------------------------------------------
'' ----------------------------------------------------------------------------

class clsPruebaAplicable
  public prueba ' as clsPrueba
  public estado
  public fechaproceso
  public aplicar
  
  private sub class_initialize()
    set prueba = new clsPrueba
    estado = ""
    fechaproceso = null
    aplicar = false
  end sub
  private sub class_terminate()
    set prueba = nothing
  end sub
end class

'' ------------------------------------

sub pruebasGetColAplicables(col,idpersona,idperfil)
  dim pruebas_modo : pruebas_modo = khorModoAutoservicio()
  dim pruebas_modo_parametro : pruebas_modo_parametro = khorValorAutoservicio()
  dim pruebas_restringidas : pruebas_restringidas = khorConfigValue(8,false)
  dim forzaSecuencia : forzaSecuencia = (khorConfigValue(292,true)<>0)
  dim listapuestos : listapuestos = ""
  dim listapruebas : listapruebas = ""
  dim sq
  '-- Inicializa listas de pruebas y/o puestos
  IF ucase(pruebas_modo)="NONE" THEN
    '-- Solo las que ya haya tomado o tenga alguna autorizacion.
    sq = "SELECT DISTINCT IdPrueba FROM Pruebas" & _
      " WHERE IdPrueba IN (SELECT IdPrueba FROM PersonaPruebaH WHERE IdPersonal=" & idpersona & ")" & _
      " OR IdPrueba IN (SELECT DISTINCT IdPrueba FROM PersonalPruebaPermiso WHERE IdPersonal=" & idpersona & ")"
    listapruebas = getBDlist("IdPrueba",sq,false)
  ELSE
    select case ucase(pruebas_modo)
      case "TODAS"
        listapruebas = getBDlist("IdPrueba","SELECT IdPrueba FROM Pruebas",false)
      case "FIJAS"
        listapruebas = pruebas_modo_parametro
      case "PERFIL"
        listapuestos = pruebas_modo_parametro
      case "USER"
        if idperfil>0 then
          listapuestos = idperfil
        end if
    end select
    'Agrega puestos relacionados con procesos de seleccion (ofertas) donde esta postulado
    listapuestos = strAdd( listapuestos , ",", pbtPuestosParaPruebas(IdPersona) )
    'Agrega pruebas relacionadas con los puestos
    if listapuestos<>"" then
      if psiComActiva() then
        dim listaNiveles : listaNiveles = getBDlist("IdNivel","SELECT DISTINCT IdNivel FROM Puestos WHERE IdPuesto IN (" & listapuestos & ")",false)
        if listaNiveles="" then listaNiveles = 0
        sq = "SELECT IdPrueba FROM vPrueba WHERE IdPrueba IN (SELECT IdPrueba FROM PuestosPruebas WHERE IdPuesto IN (" & listapuestos & ")) OR (IdPrueba IN (SELECT IdPrueba FROM vPsiCom WHERE IdNivel IN (" & listaNiveles & ")))"
      else
        sq = "SELECT DISTINCT IdPrueba FROM PuestosPruebas WHERE IdPuesto IN (" & listapuestos & ")"
      end if
      listapruebas = strAdd( listapruebas, ",", getBDlist( "IdPrueba", sq, false ) )
    end if
  END IF
  '-- Verifica cada prueba
  IF listapruebas <> "" THEN
    sq = "SELECT * FROM vPrueba WHERE idprueba IN (" & listapruebas & ") AND activa<>0 AND bAutoservicio<>0"
    if pruebas_restringidas <> "" then sq = sq & " AND idprueba NOT IN ("& pruebas_restringidas &")"
    sq = sq & " ORDER BY SecuenciaAplica, Prueba"
    dim rs : set rs = getrs(conn,sq)
    dim aplicaSiempre : aplicaSiempre = khorAplicaSiempre() AND NOT forzaSecuencia
    dim hayAplicar : hayAplicar = false
    dim rsreg : set rsreg = new clsPersonaPrueba
    dim auxo, tomado
    while not rs.EOF
      set auxo = new clsPruebaAplicable
      auxo.prueba.getFromRS rs
      tomado = rsreg.getfromDB(conn,idpersona,auxo.prueba.idprueba)
      auxo.aplicar = true
      if tomado and (auxo.prueba.mesesVigencia > 0) then
        tomado = (DateDiff("m", rsreg.fecha, Now) < auxo.prueba.mesesVigencia)
      end if
      if tomado then
        auxo.estado = descFromMenuFijo(menufijoStatusPrueba,2)
        auxo.fechaproceso = formatDateDisp(rsreg.fecha,true)
        auxo.aplicar = aplicaSiempre OR khorPersonaPruebaPermisoExists(conn,idpersona,auxo.prueba.idprueba,0)
        if auxo.prueba.idprueba=9 then
          if khorHOincompleta(IdPersona) then
            auxo.estado = descFromMenuFijo(menufijoStatusPrueba,1)
            auxo.aplicar = true
          end if
        end if
      else
        auxo.estado = descFromMenuFijo(menufijoStatusPrueba,0)
        auxo.fechaproceso = "&nbsp;"
      end if
      if auxo.prueba.idprueba=108 then
        if getBDNum("cuantos", "SELECT COUNT(*) AS cuantos FROM respuestasPruebas WHERE IdPrueba = 108 AND IdPersonal = " & IdPersona & " AND Status = 0") > 0 then
          auxo.estado = iif(khorConfigValueWithDefault(595, true, 0)<>0, descFromMenuFijo(menufijoStatusPrueba,2), lblKHOR_CalificacionPendiente)  'descFromMenuFijo(menufijoStatusPrueba,1)
          auxo.aplicar = false
        end if
      end if
      if forzaSecuencia and auxo.aplicar then
        auxo.aplicar = not hayAplicar
        hayAplicar = true
      end if
      col.add auxo, auxo.prueba.idprueba
      rs.movenext
    wend
    rs.close
    set rs = nothing
    set rsreg = nothing
  END IF
end sub

function pruebaAplicable(idpersona,idprueba)
  dim retval : retval = (superSesion() OR (adminSesion()>0))
  if not retval then
    dim colPruebas : set colPruebas = new frsCollection
    pruebasGetColAplicables colPruebas, idpersona, khorGetPerfilPersona(IdPersona)
    dim auxo : set auxo = colPruebas.objByKey(idprueba)
    if not auxo is nothing then
      retval = auxo.aplicar
    end if
    set auxo = nothing
    colPruebas.clean
    set colPruebas = nothing
  end if
  pruebaAplicable = retval
end function

'' ----------------------------------------------------------------------------
'' ---------- PRUEBA PREFIX ---------------------------------------------------
'' ----------------------------------------------------------------------------

class PruebaPrefix
  public name
  public kind
  public transform

  private sub class_initialize()
    name = ""
    kind = ""
    transform = ""
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- PRUEBA ------------------------------------------------------
'' ----------------------------------------------------------------------------

function setPruebasExternas()
  setPruebasActivas
  setPruebasExternas = setApplicationBDexists("hayPruebasExternas", "SELECT IdPrueba FROM Pruebas WHERE Activa<>0 AND IdPruebaGrupo=" & GRUPOPRUEBAEXTERNA)
end function
function hayPruebasExternas()
  hayPruebasExternas = getApplicationBDexists("hayPruebasExternas", "SELECT IdPrueba FROM Pruebas WHERE Activa<>0 AND IdPruebaGrupo=" & GRUPOPRUEBAEXTERNA)
end function

function esPruebaExterna(idp)
  esPruebaExterna = bdExists("SELECT * FROM Pruebas WHERE IdPrueba=" & idp & " AND IdPruebaGrupo="&GRUPOPRUEBAEXTERNA)
end function

function esPruebaPersonalizada(idp)
  esPruebaPersonalizada = bdExists("SELECT * FROM Pruebas WHERE IdPrueba=" & idp & " AND bPersonalizada=1 AND IdPruebaGrupo="&GRUPOPRUEBAPERSONALIZADA)
end function

function menuFijoPruebaEscala()
  dim retval
  dim aux, i
  dim arrElem, arrItem
  aux = ""
  arrElem = split(menufijoNivelEvaluacionAbrev,"|")
  for i=lbound(arrElem) to ubound(arrElem)
    if arrElem(i)<>"" then
      arrItem=split(arrElem(i),":")
      for j=2 to ubound(arrItem)
        arrItem(1) = strAdd( arrItem(1), ":", arrItem(j) )
      next
      aux=aux&iif(aux="","",",")&arrItem(1)
    end if
  next
  retval= EXT_ESC_PORCENTAJE & ":" & lblFRS_Porcentaje & "|" & EXT_ESC_NIVEL_0_6 & ":" & aux & "|" & EXT_ESC_NIVDOMINIO & ":" & lblKHOR_NivelDominioCompetencia & "|" & EXT_ESC_SEMAFORO & ":" & lblFRS_Semaforo & "|"
  menuFijoPruebaEscala = retval
end function

' ------------------------------------

class clsPruebaFactor
  public IdPruebaFactor
  public IdPrueba
  public PruebaFactor
  public auxEscala
  public CampoVistaPersona
  public CampoVistaPerfil
  public PruebaFactorAbr
  public PosPersona
  public PosPerfil
  public poloBajo
  public poloAlto
  public definicionBajo
  public definicionAlto
  '-- Auxiliar (not in bd)
  public colEscala
  public auxData
  public valPersona
  public valPerfil

  function descEscalaIdx(idx)
    dim obj
    set obj = colEscala.obj(idx)
    if obj is nothing then
      descEscalaIdx = ""
    else
      descEscalaIdx = obj.desc
    end if
  end function

  sub update(conn)
    dim sq
    if IdPruebaFactor=0 then
      sq = "INSERT INTO PruebasFactores (IdPrueba,PruebaFactor,auxEscala) VALUES (" & IdPrueba & ",'" & sqsf(PruebaFactor) & "','" & sqsf(auxEscala) & "')"
      conn.execute sq
      IdPruebaFactor = getBDnum("ultimo","SELECT MAX(IdPruebaFactor) AS ultimo FROM PruebasFactores WHERE IdPrueba=" & IdPrueba)
    else
      sq = "UPDATE PruebasFactores SET PruebaFactor='" & sqsf(PruebaFactor) & "', auxEscala='" & sqsf(auxEscala) & "' WHERE IdPruebaFactor=" & IdPruebaFactor
      conn.execute sq
    end if
  end sub

  sub getFromRS(rs)
    IdPruebaFactor = rsNum(rs,"IdPruebaFactor")
    IdPrueba = rsNum(rs,"IdPrueba")
    PruebaFactor = rsStr(rs,"PruebaFactor")
    auxEscala = rsStr(rs,"auxEscala")
    colEscala.keyDescFromMenuFijo auxEscala
    CampoVistaPersona = rsStr(rs,"CampoVistaPersona")
    CampoVistaPerfil = rsStr(rs,"CampoVistaPerfil")
    PruebaFactorAbr = rsStr(rs,"PruebaFactorAbr")
    PosPersona = rsNum(rs,"PosPersona")
    PosPerfil = rsNum(rs,"PosPerfil")
    poloBajo = rsStr(rs,"poloBajo")
    poloAlto = rsStr(rs,"poloAlto")
    definicionBajo = rsStr(rs,"definicionBajo")
    definicionAlto = rsStr(rs,"definicionAlto")
  end sub

  sub clean()
    IdPruebaFactor=0
    IdPrueba=0
    PruebaFactor=""
    auxEscala=""
    CampoVistaPersona = ""
    CampoVistaPerfil = ""
    PruebaFactorAbr = ""
    PosPersona = 0
    PosPerfil = 0
    poloBajo = ""
    poloAlto = ""
    definicionBajo = ""
    definicionAlto = ""
    colEscala.clean
    valPersona = 0
    valPerfil = 0
  end sub
  private sub class_initialize()
    set colEscala = new frsCollection
    clean
  end sub
  private sub class_terminate()
    colEscala.clean
    set colEscala = nothing
  end sub
end class

' ------------------------------------

class clsPrueba
  public IdPrueba
  public Prueba
  public bIncluye
  public bPondera
  public bPerfila
  public bInterpreta
  public Secuencia
  public Activa
  public bAutoservicio
  public Descripcion
  public IdPruebaGrupo
  public bProtocolo
  public CostoPrueba
  public mesesVigencia
  public SecuenciaAplica
  public DuracionAprox
  public bPersonalizada
  public IdEmpresa
  '--------------------
  public tipoEscala
  '--------------------
  public numFactores
  public Factor

  '--------------------

  function conPerfilacionAutomatica()
    conPerfilacionAutomatica = bPerfila AND (IdPruebaGrupo<>GRUPOPRUEBAEXTERNA) AND (IdPruebaGrupo<>GRUPOPRUEBAPERSONALIZADA) AND (idprueba<>18) AND (idprueba<>19) AND (idprueba<>20) AND (idprueba<>60) AND (idprueba<>91)
  end function

  '--------------------

  function resultadoFactores(strRes)
    dim i, res, aux
    if tipoEscala = EXT_ESC_SEMAFORO then
      res = 3
      for i = 1 to numFactores
        aux = cint( getVal( mid(strRes,(i-1)*3+1,3), true ) )
        if aux<res then res=aux
      next
    else  '-- Promedio
      res = 0
      for i = 1 to numFactores
        res = res + cint( getVal( mid(strRes,(i-1)*3+1,3), true ) )
      next
      res = div(res,numFactores)
    end if
    resultadoFactores = res
  end function

  function listaFactores(sep)
    dim retval
    retval = getBDlist("PruebaFactor","SELECT PruebaFactor FROM PruebasFactores WHERE IdPrueba=" & IdPrueba & " ORDER BY IdPruebaFactor",true)
    if sep<>"" then
      retval = replace(retval,"','",sep)
      if left(retval,1)="'" then retval = mid(retval,2)
      if right(retval,1)="'" then retval = mid(retval,1,len(retval)-1)
    end if
    listaFactores = retval
  end function

  sub addFactor(id,txt,txtesc)
    numFactores = numFactores + 1
    redim preserve Factor(numFactores)
    set Factor(numFactores) = new clsPruebaFactor
    Factor(numFactores).IdPruebaFactor=id
    Factor(numFactores).IdPrueba=IdPrueba
    Factor(numFactores).PruebaFactor=txt
    Factor(numFactores).auxEscala=txtesc
  end sub

  sub getFactoresFromBD()
    dim sq, rs
    cleanFactores
    sq = "SELECT * FROM PruebasFactores WHERE IdPrueba=" & IdPrueba & " ORDER BY IdPruebaFactor"  '-- NO cambiar ordenamiento!!!
    set rs = getrs(conn,sq)
    while not rs.eof
      numFactores = numFactores + 1
      redim preserve Factor(numFactores)
      set Factor(numFactores) = new clsPruebaFactor
      Factor(numFactores).getFromRS rs
      rs.movenext
    wend
    rs.close
    set rs = nothing
  end sub

  sub updateFactores(conn)
    dim lista, i
    lista = ""
    for i=1 to numFactores
      Factor(i).IdPrueba = IdPrueba
      Factor(i).update conn
      lista = strAdd(lista,",",Factor(i).IdPruebaFactor)
    next
    conn.execute "DELETE FROM PruebasFactores WHERE IdPrueba=" & IdPrueba & iif( lista<>"", " AND IdPruebaFactor NOT IN (" & lista & ")", "" )
  end sub

  '--------------------

  function escalaLen()
    select case tipoEscala
      case EXT_ESC_PORCENTAJE
        escalaLen = 10
      case EXT_ESC_NIVEL_0_6
        escalaLen = lenMenuFijo(menufijoNivelEvaluacionAbrev)-1
      case EXT_ESC_NIVDOMINIO
        escalaLen = khorMaxNivelCompetencia()
      case EXT_ESC_SEMAFORO
        escalaLen = 3
    end select
  end function

  function escalaDesc(valor)
    dim retval
    if tipoEscala=EXT_ESC_NIVEL_0_6 then
      retval = descFromMenuFijo(menufijoNivelEvaluacionAbrev,Valor)
      if retval="" then retval = descFromMenuFijo(menufijoNivelEvaluacionAbrev,0)
    else
      retval = valor
    end if
    escalaDesc=retval
  end function

  function escalaDescDivMenu()
    dim menu, i
    select case tipoEscala
      case EXT_ESC_PORCENTAJE
        menu = "0:-|"
        for i=1 to 10
          menu = menu & i & ":" & (i*10) & "|"
        next
      case EXT_ESC_NIVEL_0_6
        menu = menufijoNivelEvaluacionAbrev
      case EXT_ESC_NIVDOMINIO
        menu = "0:-|" & menuFijoNivelCompetencia()
    end select
    escalaDescDivMenu = menu
  end function

  function escalaDescDiv(valor)
    dim menu, retval, i
    menu = escalaDescDivMenu()
    retval = descFromMenuFijo(menu,Valor)
    if retval="" then retval = descFromMenuFijo(menu,0)
    escalaDescDiv=retval
  end function

  '--------------------

  function hayReferenciasPruebaExterna()
    hayReferenciasPruebaExterna = bdExistenReferencias(IdPrueba,"PersonaPruebaH,PuestosPruebas","IdPrueba")
  end function

  sub setDefaultsExterna()
    bIncluye = true
    Activa = true
    bAutoservicio = false
    IdPruebaGrupo = GRUPOPRUEBAEXTERNA
    bProtocolo = false
    CostoPrueba=0
    tipoEscala = EXT_ESC_PORCENTAJE
    bPondera = true
    bPerfila = true
    SecuenciaAplica = SECUENCIAPRUEBAEXTERNA
    DuracionAprox = 0
  end sub

  '-- Diseñado exclusivamente para externas y/o personalizadas
  sub update(conn)
    dim sq, curl, defl
    curl = CurrentLocale()
    defl = lblFRS_Locale
    if IdPrueba=0 then
      IdPrueba = getBDnum("newId", "SELECT MAX(IdPrueba) AS newId FROM Pruebas WHERE IdPruebaGrupo IN (" & GRUPOPRUEBAEXTERNA & "," & GRUPOPRUEBAPERSONALIZADA & ")")
      if IdPrueba=0 then IdPrueba = 1000   'Este numero se esta tomando como origen de las pruebas externas/personalizadas
      IdPrueba = IdPrueba + 1
      Secuencia = IdPrueba
      sq = "INSERT INTO pruebas (IdPrueba,Prueba,bIncluye,bPondera,bPerfila,bInterpreta,Secuencia,Activa,Descripcion,bAutoservicio,IdPruebaGrupo,bProtocolo,CostoPrueba,tipoEscala,SecuenciaAplica,DuracionAprox,bPersonalizada,IdEmpresa"
      if curl<>defl then sq = sq & ",Prueba" & curl & ",Descripcion" & curl
      sq = sq & ") VALUES (" & IdPrueba & ",'" & sqsf(Prueba) & "'," & bool2num(bIncluye) & "," & bool2num(bPondera) & "," & bool2num(bPerfila) & "," & bool2num(bInterpreta) & "," & Secuencia & "," & bool2num(Activa) & ",'" & sqsf(Descripcion) & "'," & bool2num(bAutoservicio) & "," & IdPruebaGrupo & "," & bool2num(bProtocolo) & "," & CostoPrueba & "," & tipoEscala & "," & SecuenciaAplica & "," & DuracionAprox & "," & bool2num(bPersonalizada) & "," & IdEmpresa
      if curl<>defl then sq = sq & ",'" & sqsf(Prueba) & "','" & sqsf(Descripcion) & "'"
      sq = sq & ")"
      conn.execute sq
      logAcceso LOG_ALTA, lblKHOR_Prueba, Prueba
    else
      sq = "UPDATE Pruebas SET" & _
          " Prueba" & iif(curL <>defl,curl,"") & "='" & sqsf(Prueba) & "'" & _
          ",bPondera=" & bool2num(bPondera) & _
          ",bPerfila=" & bool2num(bPerfila) & _
          ",bInterpreta=" & bool2num(bInterpreta) & _
          ",Activa=" & bool2num(Activa) & _
          ",Descripcion='" & sqsf(Descripcion) & "'" & _
          ",tipoEscala=" & tipoEscala & _
          ",SecuenciaAplica=" & SecuenciaAplica &_
          ",DuracionAprox=" & DuracionAprox &_
          ",bPersonalizada=" & bool2num(bPersonalizada) & _
          ",IdEmpresa=" & IdEmpresa &_
          " WHERE IdPrueba=" & IdPrueba
      conn.execute sq
      logAcceso LOG_CAMBIO, lblKHOR_Prueba, Prueba
    end if
    setPruebasExternas
  end sub

  sub delete(conn)
    conn.execute "DELETE FROM PruebasFactores WHERE IdPrueba=" & IdPrueba
    conn.execute "DELETE FROM PruebaCustomFactor WHERE IdPrueba=" & IdPrueba
    conn.execute "DELETE FROM pruebas WHERE IdPrueba=" & IdPrueba
    logAcceso LOG_BAJA, lblKHOR_Prueba, Prueba
    setPruebasExternas
  end sub

  function chkUname()
    dim sq
    sq = "SELECT IdPrueba FROM Pruebas WHERE Prueba='" & sqsf(Prueba) & "'"
    if IdPrueba<>0 then sq = sq & "AND IdPrueba<>" & IdPrueba
    chkUname = not bdExists(sq)
  end function

  sub setActivas(lista)
    dim sql
    sql = "UPDATE Pruebas SET Activa=0 WHERE IdPruebaGrupo<>" & GRUPOPRUEBAEXTERNA & iif(lista<>""," AND IdPrueba NOT IN (" & lista & ");","")
    conn.execute(sql)
    if (lista<>"") then
      sql = "UPDATE Pruebas SET Activa=1 WHERE IdPruebaGrupo<>" & GRUPOPRUEBAEXTERNA & " AND IdPrueba IN (" & lista & ");"
      conn.execute(sql)
    end if
    setPruebasActivas
  end sub

  sub getfromRS(rs)
    IdPrueba = rsNum(rs,"IdPrueba")
    bIncluye = rsBool(rs,"bIncluye")
    bPondera = rsBool(rs,"bPondera")
    bPerfila = rsBool(rs,"bPerfila")
    bInterpreta = rsBool(rs,"bInterpreta")
    Secuencia = rsNum(rs,"Secuencia")
    Activa = rsBool(rs,"Activa")
    bAutoservicio=rsBool(rs,"bAutoservicio")
    Descripcion=rsStrLocale(rs,"Descripcion")
    IdPruebaGrupo = rsNum(rs,"IdPruebaGrupo")
    if IdPruebaGrupo = GRUPOPRUEBAEXTERNA then
      Prueba = rsStrLocale(rs,"Prueba")
    else
      Prueba = khorNombrePruebaFromRS(rs)
    end if
    bProtocolo=rsBool(rs,"bProtocolo")
    CostoPrueba=rsNum(rs,"CostoPrueba")
    tipoEscala=rsNum(rs,"tipoEscala")
    mesesVigencia=rsNum(rs,"mesesVigencia")
    SecuenciaAplica=rsNum(rs,"SecuenciaAplica")
    DuracionAprox=rsNum(rs,"DuracionAprox")
    bPersonalizada=rsBool(rs,"bPersonalizada")
    IdEmpresa=rsNum(rs,"IdEmpresa")
  end sub

  function getfromDB(conn,id)
    dim sq, rs
    id= clng(id)
    getfromDB=false
    sq = "SELECT * FROM vPrueba WHERE IdPrueba=" & id
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub cleanFactores()
    dim i
    for i=1 to numFactores
      set Factor(i) = nothing
    next
    numc = 0
    redim Factor(numFactores)
  end sub
  sub clean()
    IdPrueba=0
    Prueba=""
    bIncluye=0
    bPondera=0
    bPerfila=0
    bInterpreta=0
    Secuencia=0
    Activa=0
    bAutoservicio=0
    Descripcion=""
    IdPruebaGrupo=0
    bProtocolo=0
    CostoPrueba=0
    tipoEscala=0
    mesesVigencia=0
    SecuenciaAplica=0
    DuracionAprox=0
    bPersonalizada=0
    IdEmpresa=0
    cleanFactores
  end sub
  private sub class_initialize()
    numFactores = 0
    clean
  end sub
  private sub class_terminate()
    cleanFactores
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- PERFIL-PRUEBA -----------------------------------------------
'' ----------------------------------------------------------------------------

class clsPerfilPrueba
  public IdPerfil
  public IdPrueba
  public Peso
  public Opciones
  public Resultado
  public Auxiliar
  public RangoInclusion
  private mDirecto

  public property get Directo()
    Directo=mDirecto
  end property

  public property get modo()
    if mDirecto then
      modo="avanzado"
    else
      modo="simple"
    end if
  end property

  function cambiaModoPerfilacion(conn)
    mDirecto = NOT mDirecto
    dim sq : sq = "UPDATE PuestosPruebas SET bDirecto=" & cint(mDirecto)
    if (NOT mDirecto) AND (len(stripCharsInBag(Opciones,"0")) = 0) then
      Resultado = ""
      sq = sq & ", Resultado=''"
    end if
    sq = sq & " WHERE idpuesto="& IdPerfil &" AND idprueba="& Idprueba
    conn.execute (sq)
    sq = getDescripcion("Puestos","IdPuesto","Puesto",idperfil)
    sq = sq & " - " & getDescripcion("Pruebas","IdPrueba","Prueba",idprueba)
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDePruebas, sq
    cambiaModoPerfilacion = modo
  end function

  sub calcResultados()
    dim o
    if not bdExists("SELECT * FROM pruebas WHERE IdPrueba=" & idprueba & " AND bProtocolo<>0") then
      resultado=opciones
    else
      set o = khorPerfilesObjectCreate()
      resultado = o.opciones2resultados(clng(idprueba),cstr(opciones))
      set o = nothing
    end if
  end sub

  sub calcOpciones()
    dim o
    set o = khorPerfilesObjectCreate()
    opciones = o.resultados2opciones(clng(idprueba),cstr(resultado))
    set o = nothing
  end sub

  Function InversoValorINTRAC(serie,resultado)
    dim res : res = clng(resultado)
    dim cMaximos(9)
    cMaximos(1) = 40 : cMaximos(2) = 20 : cMaximos(3) = 30 : cMaximos(4) = 14 : cMaximos(5) = 14 : cMaximos(6) = 20 : cMaximos(7) = 27 : cMaximos(8) = 18 : cMaximos(9) = 17
    dim minimos(9,7)
    minimos(1,0) = 0 : minimos(1,1) = 1 : minimos(1,2) = 9 : minimos(1,3) = 17 : minimos(1,4) = 24 : minimos(1,5) = 30 : minimos(1,6) = 35 : minimos(1,7) = 39
    minimos(2,0) = 0 : minimos(2,1) = 1 : minimos(2,2) = 7 : minimos(2,3) = 10 : minimos(2,4) = 13 : minimos(2,5) = 16 : minimos(2,6) = 18 : minimos(2,7) = 20
    minimos(3,0) = 0 : minimos(3,1) = 1 : minimos(3,2) = 5 : minimos(3,3) = 12 : minimos(3,4) = 19 : minimos(3,5) = 25 : minimos(3,6) = 28 : minimos(3,7) = 30
    minimos(4,0) = 0 : minimos(4,1) = 1 : minimos(4,2) = 5 : minimos(4,3) = 7 : minimos(4,4) = 9 : minimos(4,5) = 12 : minimos(4,6) = 13 : minimos(4,7) = 14
    minimos(5,0) = 0 : minimos(5,1) = 1 : minimos(5,2) = 4 : minimos(5,3) = 6 : minimos(5,4) = 9 : minimos(5,5) = 11 : minimos(5,6) = 13 : minimos(5,7) = 14
    minimos(6,0) = 0 : minimos(6,1) = 1 : minimos(6,2) = 7 : minimos(6,3) = 10 : minimos(6,4) = 13 : minimos(6,5) = 16 : minimos(6,6) = 18 : minimos(6,7) = 20
    minimos(7,0) = 0 : minimos(7,1) = 1 : minimos(7,2) = 7 : minimos(7,3) = 13 : minimos(7,4) = 18 : minimos(7,5) = 22 : minimos(7,6) = 25 : minimos(7,7) = 27
    minimos(8,0) = 0 : minimos(8,1) = 1 : minimos(8,2) = 3 : minimos(8,3) = 6 : minimos(8,4) = 9 : minimos(8,5) = 13 : minimos(8,6) = 15 : minimos(8,7) = 17
    minimos(9,0) = 0 : minimos(9,1) = 1 : minimos(9,2) = 3 : minimos(9,3) = 6 : minimos(9,4) = 9 : minimos(9,5) = 13 : minimos(9,6) = 15 : minimos(9,7) = 17
    for i=1 to 9
      for j=1 to 7
        minimos(i,j) = round( 100.0 * minimos(i,j) / cMaximos(i) )
      next
    next
    if serie>=1 and serie <=9 and res>=0 and res<=7 then
      InversoValorINTRAC = minimos(serie,res)
    else
      InversoValorINTRAC = 0
    end if
  end Function

  Function InversoValorEQ(serie,resultado)
    dim rval(21,4)
    rval(1,0) = 55	: rval(1,1) = 16	: rval(1,2) = 8	: rval(1,3) = 3	: rval(1,4) = 0
    rval(2,0) = 52	: rval(2,1) = 21	: rval(2,2) = 14	: rval(2,3) = 7	: rval(2,4) = 0
    rval(3,0) = 43	: rval(3,1) = 15	: rval(3,2) = 8	: rval(3,3) = 3	: rval(3,4) = 0
    rval(4,0) = 0	: rval(4,1) = 1	: rval(4,2) = 19	: rval(4,3) = 24	: rval(4,4) = 29
    rval(5,0) = 0	: rval(5,1) = 1	: rval(5,2) = 13	: rval(5,3) = 17	: rval(5,4) = 20
    rval(6,0) = 0	: rval(6,1) = 1	: rval(6,2) = 15	: rval(6,3) = 22	: rval(6,4) = 28
    rval(7,0) = 0	: rval(7,1) = 1	: rval(7,2) = 21	: rval(7,3) = 27	: rval(7,4) = 33
    rval(8,0) = 0	: rval(8,1) = 1	: rval(8,2) = 13	: rval(8,3) = 19	: rval(8,4) = 24
    rval(9,0) = 0	: rval(9,1) = 1	: rval(9,2) = 21	: rval(9,3) = 28	: rval(9,4) = 34
    rval(10,0) = 0	: rval(10,1) = 1	: rval(10,2) = 18	: rval(10,3) = 23	: rval(10,4) = 28
    rval(11,0) = 0	: rval(11,1) = 1	: rval(11,2) = 20	: rval(11,3) = 27	: rval(11,4) = 34
    rval(12,0) = 0	: rval(12,1) = 1	: rval(12,2) = 21	: rval(12,3) = 29	: rval(12,4) = 33
    rval(13,0) = 0	: rval(13,1) = 1	: rval(13,2) = 13	: rval(13,3) = 19	: rval(13,4) = 23
    rval(14,0) = 0	: rval(14,1) = 1	: rval(14,2) = 18	: rval(14,3) = 23	: rval(14,4) = 29
    rval(15,0) = 0	: rval(15,1) = 1	: rval(15,2) = 16	: rval(15,3) = 21	: rval(15,4) = 26
    rval(16,0) = 0	: rval(16,1) = 1	: rval(16,2) = 24	: rval(16,3) = 29	: rval(16,4) = 34
    rval(17,0) = 0	: rval(17,1) = 1	: rval(17,2) = 13	: rval(17,3) = 17	: rval(17,4) = 20
    rval(18,0) = 97	: rval(18,1) = 32	: rval(18,2) = 19	: rval(18,3) = 9	: rval(18,4) = 0
    rval(19,0) = 0	: rval(19,1) = 1	: rval(19,2) = 17	: rval(19,3) = 22	: rval(19,4) = 27
    rval(20,0) = 0	: rval(20,1) = 1	: rval(20,2) = 14	: rval(20,3) = 17	: rval(20,4) = 20
    rval(21,0) = 0	: rval(21,1) = 1	: rval(21,2) = 13	: rval(21,3) = 17	: rval(21,4) = 20
    InversoValorEQ = rval(clng(serie),clng(resultado))
  end Function
  
  function interpretacionExt(texto,corta)
    dim retval : retval = ""
    dim i
    if clng(idprueba)=68 then
      i = 0
      dim sq : sq = "SELECT Factor, Definicion FROM PruebasFactores" & _
                    " INNER JOIN LidTrFactor ON PruebasFactores.IdPruebaFactor = (6800 + LidTrFactor.IdFactor)" & _
                    " WHERE IdPrueba = 68 ORDER BY ltfSecuencia, IdPruebaFactor"
      dim rs : set rs = getrs(conn, sq)
      while not rs.eof
        i = i + 1
        retval = retval & "<p class=""khorInt_Titulo1"">" & numRomano(i) & ". " & rsStr(rs,"Factor") & "</p>"
        retval = retval & "<p class=""khorInt_Normal"">" & rsStr(rs,"Definicion") & "</p>"
        rs.movenext
      wend
      rs.close
      set rs = nothing
    else
      dim auxres : auxres = ""
      if clng(idprueba)=15 and resultado<>"" then
        for i=0 to 8
          auxres = auxres & getIntLen( InversoValorINTRAC( i+1, CInt(ValidaNumero(mid(resultado, 1 + i * 3, 3))) ), 3)
        next
      elseif clng(idprueba)=16 and len(resultado)>=63 then
        for i=0 to 20
          auxres = auxres & getIntLen( InversoValorEQ( i+1, CInt(ValidaNumero(mid(resultado, 1 + i * 3, 3))) ), 3)
        next
      else
        auxres = resultado
      end if
      retval = khorInterpretacionPrueba(idprueba,auxres,false,texto,corta)
    end if
    retval = replace( retval, "class=""khorInt_Titulo1""", "class=""khorInt_Titulo1"" style=""page-break-inside:avoid;""" )
    retval = replace( retval, "class=""khorInt_Titulo2""", "class=""khorInt_Titulo2"" style=""page-break-inside:avoid;""" )
    retval = replace( retval, "class=""khorInt_Normal""", "class=""khorInt_Normal"" style=""page-break-inside:avoid;""" )
    interpretacionExt = retval
  end function

  function interpretacion()
    interpretacion = interpretacionExt(false,false)
  end function

  sub updatePerfil(conn)
    dim sq
    sq="UPDATE PuestosPruebas SET aux=" & cint(Auxiliar) & ", Opciones='" & Opciones &"', Resultado='" & Resultado & "', Auxiliar='" & Auxiliar & "', RangoInclusion=" & cint(RangoInclusion)
    sq=sq&" WHERE idpuesto="& IdPerfil &" AND idprueba="& Idprueba
    conn.execute (sq)
    sq = getDescripcion("Puestos","IdPuesto","Puesto",idperfil)
    sq = sq & " - " & getDescripcion("Pruebas","IdPrueba","Prueba",idprueba)
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDePruebas, sq
    khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub

  sub delete(conn)
    conn.execute "DELETE FROM PuestosPruebas WHERE IdPuesto=" & IdPerfil & " AND IdPrueba=" & IdPrueba & ";"
  end sub

  sub getfromRS(rs)
    IdPerfil = rsNum(rs,"IdPuesto")
    IdPrueba = rsNum(rs,"IdPrueba")
    Peso = rsNum(rs,"Peso")
    aux = rsNum(rs,"aux")
    Opciones = rsStr(rs,"Opciones")
    Resultado = rsStr(rs,"Resultado")
    Auxiliar = rsStr(rs,"Auxiliar")
    RangoInclusion = rsNum(rs,"RangoInclusion")
    mDirecto = rsBool(rs,"bDirecto")
  end sub

  function getfromDB(conn,idper,idpru)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM PuestosPruebas WHERE IdPuesto=" & idper & " AND IdPrueba=" & idpru
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      if idprueba=91 then Resultado="007006007006005005005006"
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    IdPerfil=0
    IdPrueba=0
    Peso=0
    Opciones=""
    Resultado=""
    Auxiliar="000"
    RangoInclusion=0
    aux=0
    mDirecto=false
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- PERFIL ------------------------------------------------------
'' ----------------------------------------------------------------------------

class clsPerfil
  public IdPerfil
  public Perfil
  public Bateria
  public IdNivel
  public NotaPerfil
  public Clave
  public IdCuestionario
  public IdArea
  public vpEscala
  public vpPuntos
  public vpNivel
  public IdOcupacion
  public noAsignable
  public IdEvaluacionK
  public ClaveExterna

  private mIdEmpresa
  private mEmpresa

  public property get lockedExterno(okey)
    dim oneKey : oneKey = okey
    if isnumeric(okey) then oneKey = "PRU"
    'ToDo
    lockedExterno = (ClaveExterna <> "")
  end property

  public property get IdPerfilGenerico()
    IdPerfilGenerico = getBDnum("IdPuesto","SELECT IdPuesto FROM Puestos WHERE " & khorSQLperfilesGenericos( mIdEmpresa, IdNivel, IdArea ) & " ORDER BY IdEmpresa DESC")
  end property

  public property get lockedGenerico(okey)
    dim strM, oneKey
    if lockedExterno(okey) then
      lockedGenerico = true
    else
      if isnumeric(okey) then
        oneKey = "PRU"
      else
        oneKey = left( okey, 3 )
      end if
      strM = khorPerfilesGenericosModulos()
      lockedGenerico = (Bateria=0) AND khorPerfilesEditarSoloGenericos() AND (IdPerfilGenerico>0) AND (inStr(strM,oneKey)>0)
    end if
  end property
  
  public property get PerfilClave()
    PerfilClave = Perfil & iif(Clave<>""," ("&Clave&")","")
  end property

  public property get Empresa()
    Empresa=mEmpresa
  end property

  public property get IdEmpresa()
    IdEmpresa = mIdEmpresa
  end property

  public property let IdEmpresa(id)
    mIdEmpresa = id
    if id=0 then
      mEmpresa = ""
    else
      mEmpresa = getBD("Descripcion","SELECT Descripcion FROM vCatSucursal WHERE IdSucursal="&id)
    end if
  end property

  private function descEcualizador(tabla,incnivel)
    dim sq, auxs, campo
    if incnivel then
      if ansiSQL="yes" then
        campo="(Competencia"&db_concat&"'='"&db_concat&"TO_CHAR(Nivel)) AS nomCompetencia"
      else
        campo="(Competencia"&db_concat&"'='"&db_concat&"LTRIM(STR(Nivel))) AS nomCompetencia"
      end if
    else
      campo="Competencia AS nomCompetencia"
    end if
    sq="SELECT " & campo & " FROM "&tabla&" INNER JOIN vCompetencia ON "&tabla&".IdCompetencia=vCompetencia.IdCompetencia"
    sq=sq&" WHERE IdPuesto=" & IdPerfil & iif(ucase(tabla)="ECPUESTOCOMPETENCIANIVEL" AND ecEntrevistaPorArea(), ""," AND Nivel>0") & " ORDER BY Competencia"
    auxs=getBDlist("nomCompetencia",sq,false)
    descEcualizador=replace(auxs,",",", ")
  end function

  sub paintCopyPerfilOptions() %>
    <script language="JavaScript">
    <!--
      function copyPerfil() {
        if (getValor('idperfilcopy','int') <= 0) { alert('<%=strJS(lblFRS_DebeSeleccionar_ & lblKHOR_Perfil)%>'); return; }
        if (dirty && !confirm('<%=strJS(lblFRS_abandonarCambios)%>')) return;
        sendval('','mov','copy');
      }
    //-->
    </script>
    <TABLE BORDER="0" CELLSPACING="1" CELLPADDING="1" ALIGN="CENTER" class="noshowimp">
      <TR>
        <TD><%=lblKHOR_CopiarDelPerfil%>:</TD>
        <TD>
          <SELECT id="idperfilcopy" name="idperfilcopy" class="showSearch whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
            <option value="0"></option>
            <% optionFromCat "vPerfil","IdPerfil","perfilClave",false,khorSQLperfilesEmpresa(idempresa,iif(Bateria,1,-1)) & " AND IdPerfil<>" & idperfil,0 %>
          </SELECT>
        </TD>
        <TD>
          <input type="button" value="<%=lblFRS_Copiar%>" onClick="copyPerfil()" class="whitebtn" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
        </TD>
      </TR>
    </TABLE> <%
  end sub

  '---------- Perfil de Compatibilidad
  function descPerfil()
    dim auxs
    auxs = descRubroPrincipal(CR_RECYSEL)
    auxs = auxs & descRubroPrincipal(CR_TALENTO)
    descPerfil = auxs
  end function
  private function descRubroPrincipal(tipoRubro)
    dim auxs, sq, IdRubro, showDetail, rs, rsAux : auxs = ""
    dim campoRubro : campoRubro     = iif(tipoRubro=CR_RECYSEL,"rSel","tSel")
    dim campoEnRubro : campoEnRubro = iif(tipoRubro=CR_RECYSEL,"rEnabled","tEnabled")
    sq = "SELECT * FROM puestorubro INNER JOIN compatibilidadrubrocat ON puestorubro.idrubro=compatibilidadrubrocat.idrubro WHERE idrubrodep=0 AND "& campoEnRubro &"=1 AND "& campoRubro &"!=0 AND idpuesto="& IdPerfil
    set rs = getrs(conn,sq)
    if not rs.EOF then
      auxs = auxs &"<div style='float:left; min-width:150px; font-size:11px; margin-right:10px; white-space:nowrap;'>"
      auxs = auxs &"<span style='font-weight:bold;'>"& iif(tipoRubro=CR_RECYSEL,LBL_CR_R,LBL_CR_T) &"</span>"
      auxs = auxs &"<ul>"
      while not rs.EOF
        auxs = auxs &"<li>"& rsStr(rs,"Rubro")
        IdRubro = rsNum(rs,"IdRubro")
        showDetail = IdRubro<>tipoCR_PSI and IdRubro<>tipoCR_COM and IdRubro<>tipoCR_EXA and IdRubro<>tipoCR_ECD and IdRubro<>tipoCR_ENT
        if showDetail then
          if IdRubro=tipoCR_ESC then 'Caso especial: Escolaridad
            set rsAux = getrs(conn,"SELECT Escolaridad AS Entidad FROM PuestoEscolaridad PE INNER JOIN CatEscolaridad E ON PE.IdEscolaridad = E.IdEscolaridad WHERE PE.IdPuesto=" & IdPerfil)
          elseif IdRubro=tipoCR_EXP then 'Caso especial: Experiencia
            set rsAux = getrs(conn,"SELECT Experiencia AS Entidad FROM PuestoExperiencia PE INNER JOIN CatExperiencia E ON PE.IdExperiencia = E.IdExperiencia WHERE PE.IdPuesto=" & IdPerfil)
          elseif IdRubro=tipoCR_HAB then 'Caso especial: Habilidad
            set rsAux = getrs(conn,"SELECT Habilidad AS Entidad FROM PuestoHabilidad PH INNER JOIN CatHabilidad H ON PH.IdHabilidad = H.IdHabilidad WHERE PH.IdPuesto=" & IdPerfil)
          else
            set rsAux = getrs(conn,"SELECT Rubro AS Entidad FROM puestorubro INNER JOIN compatibilidadrubrocat ON puestorubro.idrubro=compatibilidadrubrocat.idrubro WHERE idrubrodep="& IdRubro &" AND "& campoEnRubro &"=1 AND "& campoRubro &"!=0 AND idpuesto=" & IdPerfil)
          end if
          if not rsAux.EOF then
            auxs = auxs &"<ul>"
            while not rsAux.EOF
              auxs = auxs &"<li>"& rsStr(rsAux,"Entidad") &"</li>"
              rsAux.movenext
            wend
            auxs = auxs &"</ul>"
          end if
          rsAux.close
          set rsAux = nothing
        end if
        auxs = auxs &"</li>"
        rs.movenext
      wend
      auxs = auxs &"</ul>"
      auxs = auxs &"</div>"
    end if
    rs.close
    set rs = nothing
    descRubroPrincipal = auxs
  end function
  
  sub deletePerfil()
    dim obj
    set obj=new clsPuestoRubro
    khorBorraMasivo conn, obj, "SELECT * FROM PuestoRubro WHERE IdPuesto=" & IdPerfil & ";"
    set obj=nothing
  end sub
  sub clonePerfilClean(idoriginal,clean)
    deletePerfil
    dim objpr
    set objpr = new clsPuestoRubro
    set rs = getrs(conn,"SELECT * FROM PuestoRubro WHERE IdPuesto=" & idoriginal)
    while not rs.eof
      objpr.IdPuesto = IdPerfil
      objpr.IdRubro = rsNum(rs,"IdRubro")
      objpr.clone IdOriginal
      rs.movenext
    wend
    set rs = nothing
    set objpr = nothing
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDe_ & lblKHOR_Compatibilidad, Perfil
    if clean then khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub
  sub clonePerfil(idoriginal)
    clonePerfilClean idoriginal, true
  end sub

  '---------- Pruebas de Psicometria
  function perfilado()
    dim sq, lval, aval, i, ok
    ok = false
    sq = "SELECT resultado FROM PuestosPruebas, vPrueba"
    sq = sq & " WHERE PuestosPruebas.IdPrueba=vPrueba.IdPrueba AND bPerfila<>0"
    sq = sq & " AND PuestosPruebas.IdPrueba NOT IN (18,91)"
    sq = sq & " AND IdPuesto=" & IdPerfil
    lval = getBDlist("resultado",sq,false)
    ok = (lval<>"") OR bdExists("SELECT * FROM PuestosPruebas WHERE IdPuesto=" & IdPerfil & " AND IdPrueba IN (18,91)")
    if ok then
      aval = split( lval, "," )
      for i=lbound(aval) to ubound(aval)
        ok = ok AND (trim(aval(i)) <> "")
      next
    end if
    perfilado = ok
  end function
  sub updatePruebas(conn, lpruebas,lponderaciones)
    dim sql,arrPru,arrPon
    sql = "SELECT * FROM PuestosPruebas WHERE IdPuesto=" & IdPerfil & _
          " AND NOT EXISTS(SELECT * FROM CompatibilidadRubroCat WHERE tipoRubro=" & tipoCR_PRU & " AND IdRelated=PuestosPruebas.IdPrueba)"
    if (lpruebas<>"") then sql = sql & " AND IdPrueba NOT IN (" & lpruebas & ")"
    dim objPP
    set objPP = new clsPerfilPrueba
    khorBorraMasivo conn, objPP, sql
    set objPP = nothing
    sql=""
    arrPru = split(lpruebas,",")
    arrPon = split(lponderaciones,",")
    for i = lbound(arrPru) to ubound(arrPru)
      if bdExists("SELECT * FROM PuestosPruebas WHERE idpuesto="& IdPerfil &" AND idprueba="& arrPru(i)) then
        sql = "UPDATE PuestosPruebas SET Peso=" & arrPon(i) & " WHERE idpuesto="& IdPerfil &" AND idprueba="& arrPru(i) &";"
      else
        sql = "INSERT INTO PuestosPruebas (idpuesto,IdPrueba,Peso) VALUES ("& IdPerfil &", "& arrPru(i) &"," & arrPon(i) & ");"
      end if
      conn.execute(sql)
    next
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDePruebas, Perfil
    khorBorraCompatibilidadPerfil conn, IdPerfil
    propagate2NoGenericos "PSI", false
  end sub
  function descPruebas()
    dim sq, auxs
    sq="SELECT Prueba FROM PuestosPruebas INNER JOIN vPrueba ON PuestosPruebas.IdPrueba=vPrueba.IdPrueba"
    sq=sq&" WHERE IdPuesto=" & IdPerfil & " AND Activa=1 ORDER BY secuenciagrupo, secuencia"
    auxs=getBDlist("Prueba",sq,false)
    descPruebas=replace(auxs,",",", ")
  end function
  sub deletePruebas(cond)
    conn.execute strAdd( "DELETE FROM PuestosPruebas WHERE idpuesto=" & IdPerfil, " AND ", cond ) & ";"
  end sub
  sub clonePruebasClean(idoriginal,cond,clean)
    deletePruebas cond
    insertInto conn, true, "PuestosPruebas", "IdPuesto,IdPrueba,Peso,Aux,Opciones,Resultado,Auxiliar,Observaciones,bDirecto", strAdd( "SELECT (" & IdPerfil & ") AS idpuesto,IdPrueba,Peso,Aux,Opciones,Resultado,Auxiliar,Observaciones,bDirecto FROM PuestosPruebas WHERE idpuesto=" & idoriginal, " AND ", cond )
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDePruebas, Perfil
    if clean then khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub
  sub clonePruebas(idoriginal)
    clonePruebasClean idoriginal, "", true
  end sub

  '---------- Competencias
  sub updateCompetencias(conn,listacom,listaniv,listapeso,comagregadas)
    dim sql, i, j
    dim arrId, arrVal
    sql="DELETE FROM PuestoCompetencia WHERE IdPuesto=" & IdPerfil & " AND " & khorSQLcompetencias(comagregadas)
    if listacom<>"" then
      sql=sql&" AND IdCompetencia NOT IN ("&listacom&")"
    end if
    conn.execute(sql)
    arrId = split(listacom,",")
    arrVal = split(listaniv,",")
    arrPeso = split(listapeso,",")
    j=0
    for i = lbound(arrId) to ubound(arrId)
      if bdExists("SELECT * FROM PuestoCompetencia WHERE idpuesto="& IdPerfil &" AND idcompetencia="& arrId(i)) then
        sql = "UPDATE PuestoCompetencia SET Nivel=" & arrVal(i) & ", Peso=" & arrPeso(i) & " WHERE idpuesto="& IdPerfil &" AND idcompetencia="& arrId(i) &";"
      else
        sql = "INSERT INTO PuestoCompetencia (IdPuesto,IdCompetencia,Nivel,Peso) VALUES ("& IdPerfil &", "& arrId(i) &"," & arrVal(i) & "," & arrPeso(i) & ");"
      end if
      conn.execute(sql)
    next
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDe_ & iif(comagregadas, lblKHOR_CompetenciasAgregadas, lblKHOR_CompetenciasProyectadas), perfil
    if not comagregadas then khorBorraCompatibilidadPerfil conn, IdPerfil
    propagate2NoGenericos iif(comagregadas,"COA","COM"), false
  end  sub
  function descCompetencias(incnivel,comagregadas)
    descCompetencias = descEcualizador( iif(comagregadas,"PuestoTecnicas","PuestoPsicometria"), incnivel )
  end function
  sub deleteCompetencias(comagregadas)
    conn.execute "DELETE FROM PuestoCompetencia WHERE idpuesto=" & IdPerfil & " AND " & khorSQLcompetencias(comagregadas) & ";"
  end sub
  sub cloneCompetenciasClean(idoriginal,clean,comagregadas)
    deleteCompetencias comagregadas
    insertInto conn, false, "PuestoCompetencia", "IdPuesto,IdCompetencia,Nivel,Peso", "SELECT ("&IdPerfil&") AS IdPuesto,IdCompetencia,Nivel,Peso FROM PuestoCompetencia WHERE IdPuesto=" & idoriginal & " AND " & khorSQLcompetencias(comagregadas)
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDe_ & iif(comagregadas, lblKHOR_CompetenciasAgregadas, lblKHOR_CompetenciasProyectadas), Perfil
    if clean AND not comagregadas then khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub
  sub cloneCompetencias(idoriginal,comagregadas)
    cloneCompetenciasClean idoriginal, true, comagregadas
  end sub

  '---------- Entrevista
  function numCompetenciasEC()
    numCompetenciasEC=getBDnum("cuantos","SELECT COUNT(*) AS cuantos FROM ECpuestoCompetenciaNivel WHERE IdPuesto="&IdPerfil)
  end function
  function numPreguntasEC()
    dim res
    if ecEntrevistaAleatoria() then
      numPreguntasEC=getBDnum("cuantos","SELECT SUM(npNivel+npInferior+npSuperior) AS cuantos FROM ECpuestoCompetenciaNivel WHERE IdPuesto="&IdPerfil)
    else
      numPreguntasEC=getBDnum("cuantos","SELECT COUNT(IdCompetenciaPregunta) AS cuantos FROM ECpuestoCompetenciaNivel INNER JOIN ECcompetenciaPregunta ON ECpuestoCompetenciaNivel.IdCompetencia=ECcompetenciaPregunta.IdCompetencia WHERE IdPuesto="&IdPerfil)
    end if
  end function
  function entrevistable()
    entrevistable = (idperfil<>0 AND numCompetenciasEC()>0 AND numPreguntasEC()>0)
  end function
  function descEntrevista(incnivel)
    descEntrevista = descEcualizador("ECPuestoCompetenciaNivel",incnivel)
  end function
  sub deleteEntrevista()
    dim obj
    set obj=new clsPerfilEntrevista
    khorBorraMasivo conn, obj, "SELECT * FROM ecPerfilCompetencia WHERE IdPerfil=" & IdPerfil & ";"
    set obj=nothing
  end sub
  sub cloneEntrevista(idoriginal)
    deleteEntrevista
    insertInto conn, false, "ECPuestoCompetenciaNivel", "IdPuesto,IdCompetencia,Nivel,Peso,npNivel,npInferior,npSuperior", "SELECT ("&IdPerfil&") AS IdPuesto,IdCompetencia,Nivel,Peso,npNivel,npInferior,npSuperior FROM ECPuestoCompetenciaNivel WHERE IdPuesto=" & idoriginal
    insertInto conn, true, "ECPuestoCompetenciaPregunta", "IdPuesto,IdCompetenciaPregunta", "SELECT ("&IdPerfil&") AS IdPuesto,IdCompetenciaPregunta FROM ECPuestoCompetenciaPregunta WHERE IdPuesto=" & idoriginal
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDeEntrevistaPorCompetencias, Perfil
  end sub

  '---------- Examenes
  sub updateExamenes(conn,leva,ltip,clean)
    dim sql, arrId, arrVal, i, j
    if clean then
      sql="DELETE FROM evalPuesto WHERE IdPuesto="&IdPerfil
      if leva<>"" then
        sql=sql&" AND IdEvaluacion NOT IN ("&leva&")"
      end if
      conn.execute(sql)
    end if
    sql=""
    arrId = split(leva,",")
    arrVal = split(ltip,",")
    for i = lbound(arrId) to ubound(arrId)
      if bdExists("SELECT * FROM evalPuesto WHERE idpuesto="& IdPerfil &" AND IdEvaluacion="& arrId(i)) then
        sql = "UPDATE evalPuesto SET Automatico=" & arrVal(i) & " WHERE IdPuesto="& IdPerfil &" AND IdEvaluacion="& arrId(i) &";"
      else
        sql = "INSERT INTO evalPuesto (IdPuesto,IdEvaluacion,Automatico) VALUES ("& IdPerfil &", "& arrId(i) &"," & arrVal(i) & ");"
      end if
      conn.execute(sql)
    next
    khorBorraCompatibilidadPerfil conn, IdPerfil
    logAcceso LOG_CAMBIO, lblEva_PerfilDeExamenes, Perfil
    propagate2NoGenericos "EXA", false
  end sub
  function descExamenes()
    dim sq, auxs
    sq="SELECT Evaluacion FROM evalPuesto INNER JOIN Evaluacion ON evalPuesto.IdEvaluacion=Evaluacion.IdEvaluacion"
    sq=sq&" WHERE IdPuesto=" & IdPerfil & " ORDER BY Evaluacion"
    auxs=getBDlist("Evaluacion",sq,false)
    descExamenes=replace(auxs,",",", ")
  end function
  sub deleteExamenes()
    conn.execute "DELETE FROM evalPuesto WHERE idpuesto=" & IdPerfil & ";"
  end sub
  sub cloneExamenesClean(idoriginal,clean)
    deleteExamenes
    insertInto conn, false, "evalPuesto", "IdPuesto,IdEvaluacion,Automatico","SELECT (" & IdPerfil & ") AS idpuesto,IdEvaluacion,Automatico FROM evalPuesto WHERE idpuesto=" & idoriginal
    logAcceso LOG_CAMBIO, lblEva_PerfilDeExamenes, Perfil
    if clean then khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub
  sub cloneExamenes(idoriginal)
    cloneExamenesClean idoriginal, true
  end sub
  
  '---------- ECD
  function descECD()
    descECD = descEcualizador("ECDPuestoCompetencia", false)
  end function
  sub deleteECD()
    conn.execute "DELETE FROM ECDPuestoCompetencia WHERE IdPuesto=" & IdPerfil & ";"
  end sub
  sub cloneECD(idoriginal)
    deleteEntrevista
    insertInto conn, false, "ECDPuestoCompetencia", "IdPuesto,IdCompetencia,Nivel", "SELECT ("&IdPerfil&") AS IdPuesto,IdCompetencia,Nivel FROM ECDPuestoCompetencia WHERE IdPuesto=" & idoriginal
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDeECD, Perfil
  end sub
  
  '---------- Quiz
  function descEncuestas()
    dim sq, auxs
    sq = "SELECT Name FROM QuizPuesto INNER JOIN Quiz ON QuizPuesto.IdQuiz = Quiz.IdQuiz"
    sq = sq & " WHERE IdPuesto=" & IdPerfil & " ORDER BY Name"
    auxs = getBDlist("Name", sq, false)
    descEncuestas = replace(auxs, ",", ", ")
  end function
  sub deleteEncuestas()
    conn.execute "DELETE FROM QuizPuesto WHERE IdPuesto=" & IdPerfil & ";"
  end sub
  sub cloneEncuestas(idoriginal)
    deleteEncuestas
    insertInto conn, false, "QuizPuesto", "IdQuiz,IdPuesto", "SELECT IdQuiz,(" & IdPerfil & ") AS IdPuesto FROM QuizPuesto WHERE IdPuesto=" & idoriginal
    logAcceso LOG_CAMBIO, lblECD_PerfilDeEncuestas, Perfil
  end sub

  '---------- 360
  function usadoEn360()
    Dim ng
    'Referencias en grupos con ese puesto
    ng = getBDnum("cuantos", "SELECT COUNT(*) AS cuantos FROM Grupo360 WHERE Perfil=" & IdPerfil)
    if ng=0 then
      'Referencias en grupos interfuncionales por competencias con evaluados de ese puesto
      ng = getBDnum("cuantos","SELECT COUNT(*) AS cuantos FROM Grupo360 WHERE Tipo=2 AND IdGrupo IN (SELECT DISTINCT IdGrupo FROM GrupoEntidad360 WHERE TipoEntidad=1 AND IdEntidad IN (SELECT IdPersonal FROM Personal WHERE Puesto=" & IdPerfil & "))")
    end if
    usadoEn360 = (ng > 0)
  end function
  function desc360()
    dim numcom : numcom = getBDnum("numC","SELECT COUNT(distinct idcompetencia) AS numC FROM PuestoCompetencia360 WHERE IdPuesto=" & IdPerfil)
    dim numcon : numcon = getBDnum("numC","SELECT COUNT(distinct idconducta) AS numC FROM PuestoConducta360 WHERE IdPuesto=" & IdPerfil)
    desc360 = strAdd( iif(numcom=0,"",numcom & " " & lblKHOR_Competencia_s), ", ", iif(numcon=0,"",numcon & " " & lblKHOR_Conducta_s)  )
  end function
  sub delete360()
    conn.execute "DELETE FROM PuestoConducta360 WHERE IdPuesto="&IdPerfil & ";"
    conn.execute "DELETE FROM PuestoCompetencia360 WHERE IdPuesto="&IdPerfil & ";"
    conn.execute "DELETE FROM PuestoActividad360 WHERE IdPuesto="&IdPerfil & ";"
    conn.execute "DELETE FROM PuestoObjetivo360 WHERE IdPuesto="&IdPerfil & ";"
  end sub
  sub clone360(idoriginal)
    dim sq
    dim rso, rsa, rscm
    dim ido, ida, idcm
    delete360
    '--Objetivos
    sq = "SELECT IdPuestoObjetivo AS id FROM PuestoObjetivo360 WHERE IdPuesto=" & idoriginal
    set rso = getrs(conn,sq)
    while not rso.eof
      insertInto conn, true, "PuestoObjetivo360", "IdPuesto,Nombre,Peso,Indicador,Objetivo", "SELECT ("&IdPerfil&") AS IdPuesto,Nombre,Peso,Indicador,Objetivo FROM PuestoObjetivo360 WHERE IdPuestoObjetivo=" & rsNum(rso,"id")
      ido = getBDnum("id","SELECT MAX(IdPuestoObjetivo) AS id FROM PuestoObjetivo360 WHERE IdPuesto="&IdPerfil)
      '--Actividades
      sq = "SELECT IdPuestoActividad AS id FROM PuestoActividad360 WHERE IdPuestoObjetivo=" & rsNum(rso,"id")
      set rsa = getrs(conn,sq)
      while not rsa.eof
        insertInto conn, true, "PuestoActividad360", "IdPuesto,Nombre,Peso,Actividad,IdPuestoObjetivo", "SELECT ("&IdPerfil&") AS IdPuesto,Nombre,Peso,Actividad,("&ido&") AS IdPuestoObjetivo FROM PuestoActividad360 WHERE IdPuestoActividad=" & rsNum(rsa,"id")
        ida = getBDnum("id","SELECT MAX(IdPuestoActividad) AS id FROM PuestoActividad360 WHERE IdPuesto="&IdPerfil)
        '--Competencias
        sq = "SELECT IdPuestoCompetencia AS id FROM PuestoCompetencia360 WHERE IdPuestoActividad=" & rsNum(rsa,"id")
        set rscm = getrs(conn,sq)
        while not rscm.eof
          insertInto conn, false, "PuestoCompetencia360", "IdPuesto,IdCompetencia,Peso,IdPuestoActividad", "SELECT ("&IdPerfil&") AS IdPuesto,IdCompetencia,Peso,("&ida&") AS IdPuestoActividad FROM PuestoCompetencia360 WHERE IdPuestoCompetencia=" & rsNum(rscm,"id")
          idcm = getBDnum("id","SELECT MAX(IdPuestoCompetencia) AS id FROM PuestoCompetencia360 WHERE IdPuesto="&IdPerfil)
          '--Conductas
          insertInto conn, false, "PuestoConducta360", "IdPuesto,IdConducta,Peso,IdPuestoCompetencia", "SELECT ("&IdPerfil&") AS IdPuesto,IdConducta,Peso,("&idcm&") AS IdPuestoCompetencia FROM PuestoConducta360 WHERE IdPuestoCompetencia=" & rsNum(rscm,"id")
          rscm.movenext
        wend
        rscm.close
        set rscm = nothing
        rsa.movenext
      wend
      rsa.close
      set rsa = nothing
      rso.movenext
    wend
    rso.close
    set rso = nothing
  end sub

  '---------- EDD
  function descEDD()
    dim numo : numo = getBDnum("numo","SELECT count(*) as numo FROM EDD_Objetivo WHERE IdPuesto=" & IdPerfil)
    descEDD = iif( numo=0, "", numo & " " & lblEDD_ObjetivosGralesPorPerfil )
  end function
  sub deleteEDD()
    conn.execute "DELETE FROM EDD_Objetivo WHERE IdPuesto="&IdPerfil & ";"
  end sub
  sub cloneEDD(idoriginal)
    deleteEDD
    insertInto conn, false, "EDD_Objetivo", "Por_Obj,Por_Peso,IdPuesto,IdClasificacion,Verbo,Cuanto,Que,Cuando,ParaQue,unidad","SELECT Por_Obj,Por_Peso,(" & IdPerfil & ") AS idpuesto,IdClasificacion,Verbo,Cuanto,Que,Cuando,ParaQue,unidad FROM EDD_Objetivo WHERE idpuesto=" & idoriginal
  end sub

  '---------- Valuacion
  function descValuacion()
    dim auxs : auxs = ""
    if bdExists("SELECT * FROM vpPuestoFactor WHERE IdPuesto=" & IdPerfil) then
      auxs = getBDnum("sump","SELECT SUM(puntos) as sump FROM vpPuestoFactor WHERE IdPuesto=" & IdPerfil)
    end if
    descValuacion = auxs
  end function
  sub updateValuacion()
    conn.execute "UPDATE Puestos SET vpEscala=" & vpEscala & ", vpPuntos=" & vpPuntos & ", vpNivel=" & vpNivel & " WHERE IdPuesto=" & IdPerfil & ";"
    khorBorraCompatibilidadPerfil conn, IdPerfil
    logAcceso LOG_CAMBIO, lblVP_Valuacion, Perfil
  end sub
  sub deleteValuacion()
    vpEscala = 0
    vpPuntos = 0
    vpNivel = 0
    updateValuacion
    conn.execute "DELETE FROM vpPuestoFactor WHERE idpuesto=" & IdPerfil & ";"
  end sub
  sub cloneValuacion(idoriginal)
    deleteValuacion
    insertInto conn, false, "vpPuestoFactor", "IdPuesto,IdFactor,GradoNum,Puntos","SELECT (" & IdPerfil & ") AS idpuesto,IdFactor,GradoNum,Puntos FROM vpPuestoFactor WHERE idpuesto=" & idoriginal
    dim auxo
    set auxo = new clsPerfil
    auxo.getFromDB conn, idoriginal
    vpEscala = auxo.vpEscala
    vpPuntos = auxo.vpPuntos
    vpNivel = auxo.vpNivel
    set auxo = nothing
    updateValuacion
  end sub

  '---------- Descripcion
  function descDescripcion()
    dim retval, obj
    set obj=new dpPuesto
    if obj.getfromDB(conn,IdPerfil) then
      retval = obj.describe
    end if
    set obj=nothing
    descDescripcion = retval
  end function
  sub deleteDescripcion()
    dim obj
    set obj=new dpPuesto
    if obj.getfromDB(conn,IdPerfil) then
      obj.delete conn
    end if
    set obj=nothing
  end sub
  sub cloneDescripcion(idoriginal,keepboss)
    dim obj
    set obj = new dpPuesto
    if not obj.getfromDB(conn, IdPerfil) then
      obj.IdPuesto = IdPerfil
    end if
    obj.clone idoriginal, keepboss
    set obj = nothing
  end sub

  '---------- Requisitos
  function descRequisitos()
    dim retval, obj
    set obj=new dpPerfil
    obj.IdPuesto = IdPerfil
    retval = obj.describe
    if retval = "" then
      retval = lblFRS_NoHayInformacion
    end if
    set obj=nothing
    descRequisitos = retval
  end function
  sub deleteRequisitos()
    dim obj
    set obj=new dpPerfil
    obj.IdPuesto = IdPerfil
    obj.delete conn
    set obj=nothing
  end sub
  sub cloneRequisitos(idoriginal)
    dim obj
    set obj = new dpPerfil
    obj.IdPuesto = IdPerfil
    obj.clone idoriginal
    set obj = nothing
  end sub

  '---------- Cursos
  sub addCurso(idcurso)
    conn.execute "INSERT INTO capPuestoCurso (IdPuesto,IdCurso) VALUES (" & idperfil & "," & idcurso & ");"
    logAcceso LOG_ALTA, lblCap_PerfilDeCursos, Perfil & " - " & getDescripcion("capCurso","IdCurso","Curso",idcurso)
    propagate2NoGenericos "CUR", false
  end sub
  sub delCurso(idcurso)
    conn.execute "DELETE FROM capPuestoCurso WHERE IdPuesto=" & idperfil & " AND IdCurso=" & selid & ";"
    logAcceso LOG_BAJA, lblCap_PerfilDeCursos, Perfil & " - " & getDescripcion("capCurso","IdCurso","Curso",idcurso)
    propagate2NoGenericos "CUR", false
  end sub
  function descCursos()
    dim sq, auxs
    sq = "SELECT Curso FROM capPuestoCurso INNER JOIN capCurso ON capPuestoCurso.IdCurso=capCurso.IdCurso" & _
        " WHERE IdPuesto=" & IdPerfil & " ORDER BY Curso"
    auxs=getBDlist("Curso",sq,false)
    descCursos=replace(auxs,",",", ")
  end function
  sub deleteCursos()
    conn.execute "DELETE FROM capPuestoCurso WHERE idpuesto=" & IdPerfil & ";"
  end sub
  sub cloneCursos(idoriginal)
    insertInto conn, false, "capPuestoCurso", "IdPuesto,IdCurso","SELECT (" & IdPerfil & ") AS idpuesto,IdCurso FROM capPuestoCurso WHERE idpuesto=" & idoriginal
    logAcceso LOG_CAMBIO, lblCap_PerfilDeCursos, Perfil
  end sub

  '---------- Rubro especifico de compatibilidad

  sub cloneRubro(idoriginal,IdRubro)
    dim objpr
    set objpr = new clsPuestoRubro
    objpr.IdPuesto = IdPerfil
    objpr.IdRubro = IdRubro
    objpr.clone IdOriginal
    set objpr = nothing
  end sub

  '========== MANTENIMIENTO

  sub cloneGenerico(idoriginal,clean)
    dim arrM, i
    arrM = split( khorPerfilesGenericosModulos(), "," )
    for i=lbound(arrM) to ubound(arrM)
      select case trim(arrM(i))
        case "PER"
          clonePerfilClean idoriginal, false
        case "PSI"
          clonePruebasClean idoriginal, "", false
        case "COM"
          cloneCompetenciasClean idoriginal, false, false
        case "COA"
          cloneCompetenciasClean idoriginal, false, true
        case "ECD"
          cloneECD idoriginal
        case "ENT"
          cloneEntrevista idoriginal
        case "EXA"
          cloneExamenesClean idoriginal, false
        case "CUR"
          cloneCursos idoriginal
      end select
    next
    if clean then khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub

  sub clone(conn)
    dim idoriginal
    idoriginal = IdPerfil
    insert conn
    cloneGenerico idoriginal, false
    clone360 idoriginal
    cloneEDD idoriginal
    cloneDescripcion idoriginal, false
    cloneRequisitos idoriginal
    cloneEncuestas idoriginal
    cloneECD idoriginal
    cloneValuacion idoriginal
    khorBorraCompatibilidadPerfil conn, IdPerfil
  end sub

  '-- Copia elementos del perfil actual a la lista de perfiles [listaP] la cual puede ser tambien un query que regrese IdPerfil'es
  '-- [listaKeys] es un csv de las siguientes claves:
  '     NIV : Nivel
  '     ARE : Area
  '     PER : Perfil de Compatibilidad
  '     PSI : Perfil de Pruebas (Psicometria)
  '     COM : Perfil de Competencias (Proyectadas de psicometria)
  '     COA : Perfil de Competencias (Agregadas, no proyectadas)
  '     ECD : Perfil de Evaluación de Competencias Directa (ECD)
  '     ENT : Perfil de Entrevista
  '     EXA : Perfil de Examenes
  '      QZ : Perfil de Encuestas
  '     360 : Perfil de 360
  '     EDL : Cuestionario default de EDL
  '     EDK : Evaluacion default de EDK
  '     EDD : Perfil de EDD
  '     DES : Descripción de Puesto. [paramAux] es un booleano que indica si se debe mantener el puesto jefe en la descripcion de puesto
  '     REQ : Requisitos 
  '     CUR : Perfil de Cursos
  '     PRU : Prueba especifica. IdPrueba=[paramAux]
  '     RUB : Rubro especifico de compatibilidad. IdRubro=[paramAux]
  sub propagate(listaP,listaKeys,paramAux)
    dim reg, sq, rs, savereg, clean
    dim arrOp, i
    if listaKeys<>"" then
      arrOp = split(listaKeys,",")
      for i=lbound(arrOp) to ubound(arrOp)
        arrOp(i) = ucase(trim(arrOp(i)))
      next
      set reg = new clsPerfil
      sq = "SELECT * FROM vPerfil WHERE IdPerfil IN (" & listaP & ")"
      set rs = getRS(conn,sq)
      while not rs.eof
        reg.getFromRS rs
        savereg = false
        clean = false
        for i=lbound(arrOp) to ubound(arrOp)
          select case arrOp(i)
            case "NIV"
              savereg = true
              reg.IdNivel = IdNivel
            case "ARE"
              savereg = true
              reg.IdArea = IdArea
            case "PER"
              reg.clonePerfilClean IdPerfil, false
              clean = true
            case "PSI"
              reg.clonePruebasClean IdPerfil, "", false
              clean = true
            case "COM"
              reg.cloneCompetenciasClean IdPerfil, false, false
              clean = true
            case "COA"
              reg.cloneCompetenciasClean IdPerfil, false, true
            case "ECD"
              reg.cloneECD IdPerfil
            case "ENT"
              reg.cloneEntrevista IdPerfil
            case "EXA"
              reg.cloneExamenesClean IdPerfil, false
              clean = true
            case "QZ"
              reg.cloneEncuestas IdPerfil
            case "360"
              reg.clone360 IdPerfil
            case "EDL"
              savereg = true
              reg.IdCuestionario = IdCuestionario
            case "EDK"
              savereg = true
              reg.IdEvaluacionK = IdEvaluacionK
            case "EDD"
              reg.cloneEDD IdPerfil
            case "DES"
              reg.cloneDescripcion IdPerfil, paramAux
              if khorDescripcionPuestoHabilitado() = 2 then
                reg.clone360 IdPerfil
              end if
              clean = true
            case "REQ"
              reg.cloneRequisitos IdPerfil
            case "CUR"
              reg.cloneCursos IdPerfil
            case "PRU"
              reg.clonePruebasClean IdPerfil, "IdPrueba=" & paramAux, false
              clean = true
            case "RUB"
              reg.cloneRubro IdPerfil, paramAux
              clean = true
          end select
        next
        if savereg then
          reg.update conn
        else
          logAcceso LOG_CAMBIO, lblKHOR_Perfil & " (" & lblFRS_Copia & ")", reg.Perfil
        end if
        if clean then khorBorraCompatibilidadPerfil conn, reg.IdPerfil
        '---
        rs.movenext
      wend
      rs.close
      set rs = nothing
    end if
  end sub

  sub propagate2NoGenericos(oneKey,paramAux)
    if khorPerfilesAutoHeredaGenericos() AND Bateria AND IdNivel>0 AND IdArea>0 AND inStr(khorPerfilesGenericosModulos(),oneKey)>0 then
      propagate "SELECT IdPerfil FROM vPerfil WHERE Bateria=0 AND IdNivel=" & IdNivel & " AND IdArea=" & IdArea, oneKey, paramAux
    end if
  end sub

  function catBorrable()
    dim refs
    refs = "ECPersonaPuesto,psiComPersonalCategoria,vEDD_PersonaRevision,Grupo360.Perfil,PlazaP,Plaza,PlazaH,TCTequipo,mciGrupo.IdPuestoEvaluado"
    catBorrable = NOT bdExistenReferencias(IdPerfil,refs,"IdPuesto")
  end function

  function delete(conn)
    khorBorraCompatibilidadPerfil conn, IdPerfil
    deleteCursos
    deleteValuacion
    deleteDescripcion
    deleteRequisitos
    deleteEDD
    delete360
    deleteExamenes
    deleteEncuestas
    deleteEntrevista
    deleteCompetencias true
    deleteCompetencias false
    deletePruebas ""
    deletePerfil
    conn.execute "UPDATE pbtOferta SET idperfil=0 WHERE idperfil=" & IdPerfil & ";"
    conn.execute "UPDATE Personal SET puesto=null WHERE puesto=" & IdPerfil & ";"
    conn.execute "DELETE FROM Puestos WHERE idpuesto=" & IdPerfil & ";"
    logAcceso LOG_BAJA, lblKHOR_Perfil, "(" & IdPerfil & ") " & Perfil
  end function

  sub psiComCheckPerfilPruebas()
    dim lpruebas, ap, i
    lpruebas = psiComPruebas(IdNivel)
    if lPruebas<>"" then
      ap = split( lpruebas, "," )
      for i=lbound(ap) to ubound(ap)
        if NOT bdExists("SELECT * FROM PuestosPruebas WHERE idpuesto="& IdPerfil &" AND idprueba="& ap(i)) then
          conn.execute "INSERT INTO PuestosPruebas (idpuesto,IdPrueba,Peso) VALUES ("& IdPerfil &", "& ap(i) &",0);"
        end if
      next
    end if
  end sub

  sub insert(conn)
    dim sq
    sq = "INSERT INTO Puestos (Puesto,IdEmpresa,Bateria,IdNivel,NotaPerfil,Clave,IdCuestionario,IdArea,IdOcupacion,noAsignable,IdEvaluacionK,ClaveExterna)" & _
          " VALUES ('" & sqsf(Perfil) & "'," & IdEmpresa & "," & bool2num(Bateria) & "," & IdNivel & ",'" & sqsf(NotaPerfil) & "','" & sqsf(Clave) & "'," & IdCuestionario & "," & IdArea & "," & IdOcupacion & "," & noAsignable & "," & IdEvaluacionK & ",'" & sqsf(ClaveExterna) & "');"
    conn.execute sq
    sq = "SELECT MAX(IdPuesto) AS lastid FROM Puestos WHERE IdEmpresa="&IdEmpresa&" AND Puesto='"&sqsf(Perfil)&"';"
    IdPerfil=getBD("lastid",sq)
    logAcceso LOG_ALTA, lblKHOR_Perfil, "(" & IdPerfil & ") " & Perfil
    psiComCheckPerfilPruebas
  end sub

  sub update(conn)
    dim sq, rs, auxniv, auxcue, auxeva
    set rs = getrs(conn,"SELECT IdNivel, IdCuestionario, IdEvaluacionK FROM Puestos WHERE IdPuesto="&idperfil)
    auxniv = rsNum(rs,"IdNivel")
    auxcue = rsNum(rs,"IdCuestionario")
    auxeva = rsNum(rs,"IdEvaluacionK")
    rs.close
    set rs = nothing
    sq = "UPDATE Puestos SET Puesto='" & sqsf(perfil) & "', IdNivel=" & IdNivel & ", NotaPerfil='" & sqsf(NotaPerfil) & "', Clave='" & sqsf(Clave) & "', IdCuestionario=" & IdCuestionario & ", IdArea=" & IdArea & ", IdOcupacion=" & IdOcupacion & ", noAsignable=" & noAsignable & ", IdEvaluacionK=" & IdEvaluacionK & _
        " WHERE IdPuesto=" & idperfil
    conn.execute (sq)
    logAcceso LOG_CAMBIO, lblKHOR_Perfil, "(" & IdPerfil & ") " & Perfil
    if auxniv<>IdNivel then psiComCheckPerfilPruebas
    if auxcue<>IdCuestionario and bdExists("SELECT * FROM PuestoRubro WHERE IdPuesto=" & IdPerfil & " AND IdRubro=" & tipoCR_EDL & " AND (rSel<>0 OR tSel<>0)") then
      khorBorraCompatibilidadPerfil conn, IdPerfil
    end if
    if auxeva<>IdEvaluacionK and bdExists("SELECT * FROM PuestoRubro WHERE IdPuesto=" & IdPerfil & " AND IdRubro=" & tipoCR_EDK & " AND (rSel<>0 OR tSel<>0)") then
      khorBorraCompatibilidadPerfil conn, IdPerfil
    end if
  end sub

  function chkUname(conn)
    chkUname = NOT bdExists("SELECT * FROM Puestos WHERE Puesto='" & sqsf(Perfil) & "' AND IdEmpresa=" & IdEmpresa & " AND IdPuesto<>" & IdPerfil)
  end function

  function chkUnameClave(conn)
    if Clave<>"" then
      chkUnameClave = NOT bdExists("SELECT * FROM Puestos WHERE (Puesto='" & sqsf(Perfil) & "' OR Clave='" & sqsf(Clave) & "') AND IdEmpresa=" & IdEmpresa & " AND IdPuesto<>" & IdPerfil)
    else
      chkUnameClave = chkUname(conn)
    end if
  end function

  sub getfromRS(rs)
    IdPerfil = rsNum(rs,"IdPerfil")
    Perfil = rsStr(rs,"Perfil")
    Bateria=rsBool(rs,"Bateria")
    IdEmpresa = rsNum(rs,"IdEmpresa")
    IdNivel = rsNum(rs,"IdNivel")
    NotaPerfil = rsStr(rs,"NotaPerfil")
    Clave = rsStr(rs,"Clave")
    IdCuestionario = rsNum(rs,"IdCuestionario")
    IdArea = rsNum(rs,"IdArea")
    vpEscala = rsNum(rs,"vpEscala")
    vpPuntos = rsNum(rs,"vpPuntos")
    vpNivel = rsNum(rs,"vpNivel")
    IdOcupacion = rsNum(rs,"IdOcupacion")
    noAsignable = rsNum(rs,"noAsignable")
    IdEvaluacionK = rsNum(rs,"IdEvaluacionK")
    ClaveExterna = rsStr(rs,"ClaveExterna")
  end sub

  function getfromDB(conn,id)
    dim sq, rs
    getfromDB=false
    if id<>"" then
      sq = "SELECT * FROM vPerfil WHERE IdPerfil=" & id
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDB=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  function getfromDBfilter(conn,filter)
    dim sq, rs
    getfromDBfilter=false
    if filter<>"" then
      sq = "SELECT * FROM vPerfil WHERE " & filter
      set rs = getrs(conn,sq)
      if not rs.EOF then
        getfromRS(rs)
        getfromDBfilter=true
      end if
      rs.close
      set rs = nothing
    end if
  end function

  sub clean()
    IdPerfil=0
    Perfil=""
    Bateria=false
    IdNivel=0
    mIdEmpresa=0
    mEmpresa=0
    Clave=""
    IdCuestionario = 0
    IdArea = 0
    vpEscala = 0
    vpPuntos = 0
    vpNivel = 0
    IdOcupacion = 0
    noAsignable = 0
    IdEvaluacionK = 0
    ClaveExterna = ""
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- COMPETENCIA NIVEL -------------------------------------------------
'' ----------------------------------------------------------------------------
class clsCompetenciaNivel
  public IdCompetencia
  public Nivel
  public IdNivelPerfil
  public Titulo
  public Descripcion

  sub update(conn)
    dim sql
    if bdExists("SELECT * FROM ECcompetenciaNivel WHERE IdCompetencia="& IdCompetencia &" AND Nivel="& Nivel &" AND IdNivelPerfil=" & IdNivelPerfil) then
      sql = "UPDATE ECcompetenciaNivel SET Titulo='" & sqsf(Titulo) & "', Descripcion='" & sqsf(Descripcion) & "' WHERE IdCompetencia="& IdCompetencia &" AND Nivel="& Nivel &" AND IdNivelPerfil=" & IdNivelPerfil & ";"
    else
      sql = "INSERT INTO ECcompetenciaNivel (IdCompetencia,Nivel,IdNivelPerfil,Titulo,Descripcion) VALUES ("& IdCompetencia &", "& Nivel &"," & IdNivelPerfil & ",'" & sqsf(Titulo) & "','" & sqsf(Descripcion) & "');"
    end if
    conn.execute sql
  end sub

  sub getfromRS(rs)
    IdCompetencia=rsNum(rs,"IdCompetencia")
    Nivel=rsNum(rs,"Nivel")
    IdNivelPerfil=rsNum(rs,"IdNivelPerfil")
    Titulo=trim(rsStr(rs,"Titulo"))
    Descripcion=rsStr(rs,"Descripcion")
  end sub

  function getfromDB(conn,idc,nivdom,nivper)
    dim sq, rs
    getfromDB=false
    clean
    idcompetencia = idc
    if khorInterpretacionNivelPerfil() then
      nivel = 0
      idnivelperfil = cint(nivper)
    else
      nivel = cint(nivdom)
      idnivelperfil = 0
    end if
    nivel = cint(nivdom)
    sq = "SELECT * FROM ECcompetenciaNivel WHERE IdCompetencia=" & IdCompetencia
    if idnivelperfil>0 then sq = sq & " AND IdNivelPerfil=" & idnivelperfil
    if nivel>0 or idnivelperfil=0 then sq = sq & " AND Nivel=" & Nivel
    sq = sq & " ORDER BY IdNivelPerfil,Nivel"
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    IdCompetencia=0
    Nivel=0
    IdNivelPerfil=0
    Titulo=""
    Descripcion=""
  end sub
  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class


'' ----------------------------------------------------------------------------
'' ---------- COMPETENCIA -------------------------------------------------
'' ----------------------------------------------------------------------------
class clsCompetencia
  public IdCompetencia
  public Competencia
  public Tipo
  public Definicion
  public Inactiva
  public Clave
  public IdCompetenciaGrupo
  private mIdEmpresa
  private mEmpresa
  '-- Coleccion auxiliar de nivel y nivel seleccionado
  public colNivel
  public selNivel

  public property get Predefinida()
    Predefinida = bdExists("SELECT * FROM MatrizCompetencias WHERE IdCompetencia="&IdCompetencia)
  end property

  public property get Empresa()
    Empresa=mEmpresa
  end property

  public property get IdEmpresa()
    IdEmpresa = mIdEmpresa
  end property

  public property let IdEmpresa(id)
    mIdEmpresa = id
    if id=0 then
      mEmpresa=""
    else
      mEmpresa = getBD("Descripcion","SELECT Descripcion FROM vCatSucursal WHERE IdSucursal="&id)
    end if
  end property

  function catBorrable()
    dim referencias
    referencias = "ECPersonaCompetencia,PersonalEvaluacion360,PuestoCompetencia360,GrupoCompetenciaEvaluador360,EvaCatTemaCompetencia,EvaCatPregunta,mciGrupoResultado,pcsAccionCompetencia"
    catBorrable = (NOT Predefinida()) AND (NOT bdExistenReferencias(IdCompetencia,referencias,"IdCompetencia")) _
                  AND (NOT bdNumReferenciasFilter(IdCompetencia,"GrupoEntidad360","IdEntidad","TipoEntidad=4"))
  end function

  function delete(conn)
    dim sq
    dim obj
    conn.execute "DELETE FROM mciReactivo WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM DBusuarioCompetencia WHERE IdCompetencia=" & IdCompetencia & ";"
    ''Resultados
    conn.execute "DELETE FROM ECpersonaPregunta WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM ECpersonaCompetencia WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM PersonalCompetencias WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM CompatibilidadCompetencias WHERE IdCompetencia=" & IdCompetencia & ";"
    ''Perfiles
    set obj=new clsPerfilEntrevista
    khorBorraMasivo conn, obj, "SELECT * FROM ecPerfilCompetencia WHERE IdCompetencia=" & IdCompetencia & ";"
    set obj=nothing
    conn.execute "DELETE FROM PuestoCompetencia WHERE IdCompetencia=" & IdCompetencia & ";"
    ''Definicion
    set obj=new clsConducta
    khorBorraMasivo conn, obj, "SELECT * FROM CatConductasObs360 WHERE IdCompetencia=" & IdCompetencia
    set obj=nothing
    conn.execute "DELETE FROM ECcompetenciaNivel WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM ECcompetenciaPregunta WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM MatrizCompetencias WHERE IdCompetencia=" & IdCompetencia & ";"
    conn.execute "DELETE FROM CatCompetencias360 WHERE IdCompetencia=" & IdCompetencia & ";"
    logAcceso LOG_BAJA, lblKHOR_Competencia, Competencia
  end function

  sub insert(conn)
    dim sq
    sq = "INSERT INTO CatCompetencias360 (Competencia,Tipo,IdEmpresa,Definicion,Clave,IdCompetenciaGrupo) VALUES ('" & sqsf(Competencia) & "'," & Tipo & "," & IdEmpresa & ",'" & sqsf(Definicion) & "','" & sqsf(Clave) & "'," & IdCompetenciaGrupo & ");"
    conn.execute sq
    sq = "SELECT MAX(IdCompetencia) AS lastid FROM CatCompetencias360 WHERE IdEmpresa="&IdEmpresa&" AND Competencia='"&sqsf(Competencia)&"';"
    IdCompetencia=getBD("lastid",sq)
    logAcceso LOG_ALTA, lblKHOR_Competencia, Competencia
  end sub

  sub update(conn)
    dim sq
    sq = "UPDATE CatCompetencias360 SET Competencia='" & sqsf(Competencia) & "', Tipo=" & Tipo & ", Definicion='" & sqsf(Definicion) & "', Clave='" & sqsf(Clave) & "', IdCompetenciaGrupo=" & IdCompetenciaGrupo & " WHERE IdCompetencia=" & IdCompetencia
    conn.execute (sq)
    logAcceso LOG_CAMBIO, lblKHOR_Competencia, Competencia
  end sub

  sub clone(conn)
    dim idoriginal
    idoriginal = IdCompetencia
    insert conn
    insertInto conn, false, "ECcompetenciaNivel", "IdCompetencia,Nivel,Titulo,Descripcion", "SELECT (" & IdCompetencia & ") AS IdCompetencia,Nivel,Titulo, Descripcion FROM ECcompetenciaNivel WHERE IdCompetencia="  & idoriginal & ";"
    insertInto conn, true, "ECcompetenciaPregunta", "IdCompetencia,Nivel,Pregunta,IdArea", "SELECT (" & IdCompetencia & ") AS IdCompetencia,Nivel,Pregunta,IdArea FROM ECcompetenciaPregunta WHERE IdCompetencia="  & idoriginal & ";"
    insertInto conn, true, "CatConductasObs360", "IdCompetencia,Peso,Conducta,Cursos,Niveles", "SELECT (" & IdCompetencia & ") AS IdCompetencia,Peso,Conducta,Cursos,Niveles FROM CatConductasObs360 WHERE IdCompetencia="  & idoriginal & ";"
  end sub

  function chkUname(conn)
    dim sq, rs
    sq = "SELECT * FROM vCompetencia WHERE Competencia='" & sqsf(Competencia) & "' AND IdEmpresa=" & IdEmpresa
    set rs = getrs(conn,sq)
    if rs.EOF then
      chkUname = true
    else
      chkUname = (rs("IdCompetencia")&""=IdCompetencia&"")
    end if
    rs.close
    set rs = nothing
  end function

  sub getfromRS(rs)
    IdCompetencia=rsNum(rs,"IdCompetencia")
    Competencia=rsStr(rs,"Competencia")
    Tipo=rsNum(rs,"Tipo")
    Definicion=rsStr(rs,"Definicion")
    IdEmpresa=rsNum(rs,"IdEmpresa")
    Inactiva=rsBool(rs,"Inactiva")
    Clave=rsStr(rs,"Clave")
    IdCompetenciaGrupo=rsNum(rs,"IdCompetenciaGrupo")
  end sub

  function getfromDB(conn,idcom)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM vCompetencia WHERE IdCompetencia=" & idcom
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    dim i
    IdCompetencia=0
    Competencia=""
    Tipo=0
    Definicion=""
    IdEmpresa=0
    Inactiva=false
    Clave=""
    IdCompetenciaGrupo=0
    selNivel=0
  end sub

  private sub class_initialize()
    clean
    set colNivel = new frsCollection
  end sub
  private sub class_terminate()
    colNivel.clean
    set colNivel = nothing
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- COMPETENCIA PREGUNTA ----------------------------------------
'' ----------------------------------------------------------------------------
class clsCompetenciaPregunta
  public IdCompetenciaPregunta
  public IdCompetencia
  public Nivel
  public Pregunta
  public IdArea

  function delete(conn)
    conn.execute "DELETE FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta=" & IdCompetenciaPregunta & ";"
    logAcceso LOG_BAJA, lblEnt_EntrevistaPorCompetencias & ":" & lblKHOR_Reactivo, getDescripcion("vCompetencia","IdCompetencia","Competencia",IdCompetencia)
  end function

  sub insert(conn)
    dim sq : sq = "INSERT INTO ECcompetenciaPregunta (IdCompetencia,Nivel,Pregunta,IdArea)" & _
                  " VALUES (" & IdCompetencia & "," & Nivel & ",'" & sqsf(Pregunta) & "'," & IdArea & ");"
    conn.execute sq
    sq = "SELECT MAX(IdCompetenciaPregunta) AS lastid FROM ECcompetenciaPregunta WHERE IdCompetencia="&IdCompetencia&" AND Nivel="&Nivel&" AND IdArea=" & IdArea & ";"
    IdCompetenciaPregunta = getBD("lastid",sq)
    logAcceso LOG_ALTA, lblEnt_EntrevistaPorCompetencias & ":" & lblKHOR_Reactivo, getDescripcion("vCompetencia","IdCompetencia","Competencia",IdCompetencia)
  end sub

  sub update(conn)
    conn.execute "UPDATE ECcompetenciaPregunta SET Nivel=" & Nivel & ", Pregunta='" & sqsf(Pregunta) & "', IdArea=" & IdArea & " WHERE IdCompetenciaPregunta=" & IdCompetenciaPregunta
    logAcceso LOG_CAMBIO, lblEnt_EntrevistaPorCompetencias & ":" & lblKHOR_Reactivo, getDescripcion("vCompetencia","IdCompetencia","Competencia",IdCompetencia)
  end sub

  sub getfromRS(rs)
    IdCompetenciaPregunta=rsNum(rs,"IdCompetenciaPregunta")
    IdCompetencia=rsNum(rs,"IdCompetencia")
    Nivel=rsNum(rs,"Nivel")
    Pregunta=rsStr(rs,"Pregunta")
    IdArea=rsNum(rs,"IdArea")
  end sub

  function getfromDB(conn,idcp)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta=" & idcp
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    IdCompetenciaPregunta=0
    IdCompetencia=0
    Nivel=0
    Pregunta=""
    IdArea=0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- ENTREVISTA --------------------------------------------------
'' ----------------------------------------------------------------------------
class ecEntrevista
  public IdEntrevista
  public IdPersona
  public IdPerfil
  public Fecha
  public Resultado
  public Compatibilidad
  public Entrevistador
  public Comentarios
  public Status
  public IdUsuarioEntrevistador
  
  function puedeEntrevistar()
    dim ok : ok = true
    dim ecXUsuario : ecXUsuario = ecEntrevistaXUsuario()
    dim IdPersonaSesion : IdPersonaSesion = personaSesion()
    if (IdPersonaSesion > 0) AND (ecXUsuario = 2) then
      ok = (IdUsuarioEntrevistador = IdPersonaSesion)
    elseif (ecXUsuario=1) AND (khorCurrentUser.tipo <> ADMIN_GENERAL) then
      ok = false
      if IdUsuarioEntrevistador > 0 then
        dim reg : set reg = new clsUsuario
        if reg.getfromDB(conn,IdUsuarioEntrevistador) then
          ok = ( inCSV( strAdd( reg.IdUsuario, ",", reg.parents ), khorCurrentUser.IdUsuario ) >= 0 )
        end if
        set reg = nothing
      end if
    end if
    puedeEntrevistar = ok
  end function
  
  sub popUpUsuario(IdEmpresa)
    dim sq, filtroUsers
    dim ecXUsuario : ecXUsuario = ecEntrevistaXUsuario()
    if ecXUsuario = 2 then
      'Variable sucia, declarar en khorLabelEspecial.asp o similar
      'ec_filtro_entrevistadores = "" '-- Filtro de personas que pueden ser entrevistadores (tabla personal)
      filtroUsers = strAdd( khorCondicionUsuario("IdPersonal"), " AND ", ec_filtro_entrevistadores )
      sq = "SELECT IdPersonal AS IdUsuario, (Nombre+(CASE WHEN statusper<>0 THEN ' - " & lblFRS_Inactivo & "' ELSE '' END)) AS NombreStatus" & _
          " FROM Personal"
      sq = strAdd(sq, " WHERE ", filtroUsers) & " ORDER BY Nombre"
    else
      filtroUsers = iif( khorMultiSucursal() AND IdEmpresa<>0, "EXISTS(SELECT * FROM DBusuarioSucursal WHERE IdUsuario=vDBusuarios.IdUsuario AND IdSucursal=" & IdEmpresa & ")", "" )
      filtroUsers = strAdd( filtroUsers, " AND ", "EXISTS(SELECT * FROM DBpermisos WHERE IdUsuario=vDBusuarios.IdUsuario AND IdModulo IN (" & Modulo_EntrevistaRegistro & "))" )
      sq = "SELECT IdUsuario, (Descripcion+(CASE WHEN Activo<>1 THEN ' - " & lblFRS_Inactivo & "' ELSE '' END)) AS NombreStatus" & _
          " FROM vDBusuarios WHERE Admin=1 OR (" & filtroUsers & ") ORDER BY Descripcion"
    end if
    popUpBegin "ecUsuario", lblKHOR_Entrevistador, "", "" %>
    <div style="text-align:center;">
      <select id="movIdUsuarioEC" name="movIdUsuarioEC" class="<%=iif(ecXusuario=2,"showSearch ","")%>whiteblur" onMouseOut="inOut(this);" onMouseOver="inOver(this);" onBlur="inBlur(this);" onFocus="inFocus(this);">
        <option value="0"></value>
        <% optionFromQuery "IdUsuario", "NombreStatus", 0, sq %>
      </select>
      <br/>
      <INPUT type="button" id="ecUsuarioAceptar" name="ecUsuarioAceptar" value="<%=lblFRS_Aceptar%>" onclick="ecUserAccept();" class="whitebtn" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this);" onmouseout="inOut(this);">
      <INPUT type="button" id="ecUsuarioCancelar" name="ecUsuarioCancelar" value="<%=lblFRS_Cancelar%>" onclick="popUp_hide('ecUsuario');" class="whitebtn" onblur="inBlur(this);" onmouseover="inOver(this);" onfocus="inFocus(this);" onmouseout="inOut(this);">
    </div> <%
    popUpEnd %>
    <input type="hidden" id="movUsuarioEC" name="movUsuarioEC" value="">
    <input type="hidden" id="movIdEntrevista" name="movIdEntrevista" value=""> <%
  end sub

  function enOferta(conn)
    enOferta=BDexists("SELECT IdOferta FROM pbtOfertaPersona WHERE IdEntrevista="&IdEntrevista)
  end function

  function respondida()
    respondida = (Status<>0)
    'respondida=bdExists("SELECT * FROM ECPersonaPregunta WHERE Respuesta<>0 AND IdEntrevista="&IdEntrevista)
  end function

  function create(conn,idps,idpf,idusuarioec)
    dim sq, regpf, ok
    ok = false
    set regpf = new clsPerfil
    regpf.getfromDB conn, clng("0"&idpf)
    if regpf.entrevistable() then
      IdPersona=idps
      IdPerfil=regpf.idperfil
      IdUsuarioEntrevistador = idusuarioec
      Fecha=Now
      Status=0
      insert conn
      insertInto conn, false, "ECPersonaCompetencia", "IdEntrevista,IdCompetencia,Nivel,Resultado,Comentarios", "SELECT ("& IdEntrevista &") AS IdEntrevista, IdCompetencia, Nivel, (0) AS Resultado, ('') AS Comentarios FROM ecPerfilCompetencia WHERE IdPerfil=" & IdPerfil
      if ecEntrevistaAleatoria() then
        dim presel,lista,obj,rs
        if ecEntrevistaPorPerfil() then
          sq="SELECT IdCompetenciaPregunta FROM ECpuestoCompetenciaPregunta WHERE IdPuesto="&IdPerfil
          presel=getBDlist("IdCompetenciaPregunta",sq,false)
        else
          presel=""
        end if
        set obj=new clsPerfilEntrevista
        sq="SELECT * FROM ecPerfilCompetencia WHERE IdPerfil="&IdPerfil
        set rs=getrs(conn,sq)
        while not rs.eof
          obj.getfromRS rs
          lista=obj.getEntrevistaAleatoria(presel)
          if lista<>"" then
            insertInto conn, true, "ECPersonaPregunta", "IdEntrevista,IdCompetencia,Nivel,Respuesta,Pregunta,IdArea", "SELECT ("& IdEntrevista &") AS IdEntrevista, ECCompetenciaPregunta.IdCompetencia, ECCompetenciaPregunta.Nivel, (0) AS Respuesta, Pregunta, IdArea FROM ECCompetenciaPregunta WHERE IdCompetenciaPregunta IN ("&lista&")"
          end if
          rs.movenext
        wend
        rs.close
        set rs=nothing
        set obj=nothing
      else
        insertInto conn, true, "ECPersonaPregunta", "IdEntrevista,IdCompetencia,Nivel,Respuesta,Pregunta,IdArea", "SELECT ("& IdEntrevista &") AS IdEntrevista, ECCompetenciaPregunta.IdCompetencia, ECCompetenciaPregunta.Nivel, (0) AS Respuesta, Pregunta, IdArea FROM ECCompetenciaPregunta INNER JOIN ecPerfilCompetencia ON ECCompetenciaPregunta.IdCompetencia=ecPerfilCompetencia.IdCompetencia WHERE IdPerfil=" & IdPerfil
      end if
      ok = true
    end if
    set regpf = nothing
    create=ok
  end function

  function evaluate(conn)
    dim escala : escala = ecEntrevistaEscala()
    dim rs,sql,res
    dim ncd: ncd = 0
    Resultado=0
    Compatibilidad=0
    ''Genera resultado de cada competencia
    sql="UPDATE ECpersonaCompetencia SET resultado=0, compatibilidad=0 WHERE IdEntrevista=" & IdEntrevista & ";"
    conn.execute sql
    if ansiSQL="yes" then
      sql="SELECT IdCompetencia, ROUND(AVG(Respuesta)) AS Resultado, ROUND(AVG(Respuesta/" & escala & "),4) AS Compatibilidad" & _
          " FROM ECPersonaPregunta WHERE IdEntrevista="&IdEntrevista&" AND Respuesta>0 GROUP BY IdCompetencia;"
    else
      sql="SELECT IdCompetencia, AVG(CONVERT(float,Respuesta)) AS Resultado, ROUND(AVG(CONVERT(float,Respuesta)/" & escala & "),4) AS Compatibilidad" & _
          " FROM ECPersonaPregunta WHERE IdEntrevista="&IdEntrevista&" AND Respuesta>0 GROUP BY IdCompetencia;"
    end if
    set rs=getrs(conn,sql)
    while not rs.eof
      res = cdbl(rs("resultado"))
      sql = "UPDATE ECpersonaCompetencia SET Resultado="&res&", Compatibilidad="&cdbl(rs("compatibilidad"))&" WHERE IdEntrevista="&IdEntrevista&" AND IdCompetencia="&rs("IdCompetencia")&";"
      conn.execute sql
      if res > 0 then
        ncd = ncd + 1
        Resultado = Resultado + res
        Compatibilidad = Compatibilidad + cdbl(rs("compatibilidad"))
      end if
      rs.movenext
    wend
    rs.close
    set rs = nothing
    ''Genera resultado del perfil
    dim nct : nct = ncd
    '-- Variable dirty, definir en khorLabelEspecial.asp o similar
    'ecExcluyeCompetenciasNoEvaluadas = true  '-- EC: Excluye competencias no evaluadas de la calificación general
    if not ecExcluyeCompetenciasNoEvaluadas then
      nct = getBDnum("numc","SELECT COUNT(IdCompetencia) As numc FROM ecPersonaCompetencia WHERE IdEntrevista = " & IdEntrevista)
    end if
    if nct > 0 then
      Resultado = round(Resultado / nct, 4)
      Compatibilidad = round(Compatibilidad / nct, 4)
      Status = 1
    else
      Resultado=0
      Compatibilidad=0
    end if
    sql="UPDATE ECpersonaPuesto SET Resultado="&Resultado&", Compatibilidad="&Compatibilidad&", Status=" & Status & " WHERE IdEntrevista="&IdEntrevista&";"
    conn.execute sql
    khorBorraCompatibilidadPersonaPuestos conn, idpersona, idperfil
    onBorraCompatibilidadPersonaRubro conn,idpersona,tipoCR_ENT,0,"",CStr(idperfil)
    dim sq
    sq = getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
    sq = sq & " - " & getDescripcion("vPersona","IdPersona","Clave",idpersona)
    logAcceso LOG_GENERICO, lblEnt_EntrevistaPorCompetencias&":"&lblFRS_Aplicacion, sq
  end function

  function delete(conn)
    conn.execute "UPDATE pbtOfertaPersona SET FechaUltimoCambio=" & formatDateSQL(Now,true) & ", IdEntrevista=0 WHERE IdEntrevista=" & IdEntrevista & ";"
    conn.execute "DELETE FROM ECPersonaPregunta WHERE IdEntrevista=" & IdEntrevista & ";"
    conn.execute "DELETE FROM ECPersonaCompetencia WHERE IdEntrevista=" & IdEntrevista & ";"
    conn.execute "DELETE FROM ECPersonaPuesto WHERE IdEntrevista=" & IdEntrevista & ";"
    khorBorraCompatibilidadPersonaPuestos conn, idpersona, idperfil
    Dim sq
    sq = getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
    sq = sq & " - " & getDescripcion("vPersona","IdPersona","Clave",idpersona)
    logAcceso LOG_BAJA, lblEnt_EntrevistaPorCompetencias, sq
  end function

  sub insert(conn)
    if ecEntrevistaXUsuario()<>0 then Entrevistador = ecEntrevistaNombreUsuario(IdUsuarioEntrevistador)
    dim sq : sq = "INSERT INTO ECPersonaPuesto (IdPersonal,IdPuesto,Fecha,Resultado,Compatibilidad,Entrevistador,Comentarios,Status,IdUsuarioEntrevistador)" & _
                  " VALUES (" & IdPersona & "," & IdPerfil & "," & formatDateSQL(Fecha,true) &"," & Resultado & "," & Compatibilidad & ",'" & sqsf(Entrevistador) & "','" & sqsf(Comentarios) & "'," & Status & "," & IdUsuarioEntrevistador & ");"
    conn.execute sq
    sq = "SELECT MAX(IdEntrevista) AS lastid FROM ECPersonaPuesto WHERE IdPersonal="&IdPersona&" AND IdPuesto="&IdPerfil&";"
    IdEntrevista=getBD("lastid",sq)
    sq = getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
    sq = sq & " - " & getDescripcion("vPersona","IdPersona","Clave",idpersona)
    logAcceso LOG_ALTA, lblEnt_EntrevistaPorCompetencias, sq
  end sub

  sub update(conn)
    if ecEntrevistaXUsuario()<>0 then Entrevistador = ecEntrevistaNombreUsuario(IdUsuarioEntrevistador)
    dim sq :  sq = "UPDATE ECPersonaPuesto SET Resultado=" & Resultado & ", Compatibilidad=" & Compatibilidad &", Entrevistador='" & sqsf(Entrevistador) & "', Comentarios='" & sqsf(Comentarios) & "', Status=" & Status & ", IdUsuarioEntrevistador=" & IdUsuarioEntrevistador & _
                  ", Fecha = " & formatDateSQL(Fecha,true) & " WHERE IdEntrevista=" & IdEntrevista
    conn.execute (sq)
    sq = getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
    sq = sq & " - " & getDescripcion("vPersona","IdPersona","Clave",idpersona)
    'logAcceso LOG_CAMBIO, lblEnt_EntrevistaPorCompetencias, sq
  end sub

  sub getfromRS(rs)
    IdEntrevista=rsNum(rs,"IdEntrevista")
    IdPersona=rsNum(rs,"IdPersona")
    IdPerfil=rsNum(rs,"IdPerfil")
    Fecha=rs("Fecha")
    Resultado=rsNum(rs,"Resultado")
    Compatibilidad=rsNum(rs,"Compatibilidad")
    Entrevistador=rsStr(rs,"Entrevistador")
    Comentarios=rsStr(rs,"Comentarios")
    Status=rsNum(rs,"Status")
    IdUsuarioEntrevistador = rsNum(rs,"IdUsuarioEntrevistador")
  end sub

  function getfromDB(conn,idcp)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM ecEntrevista WHERE IdEntrevista=" & idcp
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    IdEntrevista=0
    IdPersona=0
    IdPerfil=0
    Fecha=Date
    Resultado=0
    Compatibilidad=0
    Entrevistador=""
    Comentarios=""
    Status=0
    IdUsuarioEntrevistador=0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- PERFIL ENTREVISTA -------------------------------------------
'' ----------------------------------------------------------------------------
class clsPerfilEntrevista
  private myIdPerfil
  public IdCompetencia
  public Nivel
  public Peso
  public NumPreg(3)
  private myIdArea
  
  public property let IdPerfil(idp)
    myIdPerfil = idp
    myIdArea = getBDnum("IdArea","SELECT IdArea FROM Puestos WHERE IdPuesto=" & myIdPerfil)
  end property
  public property get IdPerfil()
    IdPerfil = myIdPerfil
  end property
  public property get IdArea()
    IdArea = myIdArea
  end property

  public function listaPreguntas()
    dim sq : sq = "SELECT ECpuestoCompetenciaPregunta.IdCompetenciaPregunta FROM ECpuestoCompetenciaPregunta" & _
                  " INNER JOIN ECCompetenciaPregunta ON ECpuestoCompetenciaPregunta.IdCompetenciaPregunta=ECCompetenciaPregunta.IdCompetenciaPregunta" & _
                  " WHERE IdPuesto =" & IdPerfil & " AND IdCompetencia =" & IdCompetencia
    if ecEntrevistaPorArea() then
      sq = sq & " AND (IdArea = " & myIdArea & iif(myIdArea=0,""," OR IdArea=0") & ")"
    end if
    sq = sq & " ORDER BY Nivel DESC, ECpuestoCompetenciaPregunta.IdCompetenciaPregunta"
    listaPreguntas = getBDlist( "IdCompetenciaPregunta", sq, false )
  end function

  function listaAleatoria(tipo,original)
    dim disponibles, presel 'listas
    dim nlim, ndis, nsel    'contadores
    dim sq
    '-- Obtiene preguntas disponibles del <tipo>
    sq = "SELECT IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetencia=" & IdCompetencia
    if tipo=-1 then
      nlim = clng(NumPreg(1))
      sq = sq & " AND (IdArea = " & myIdArea & iif(myIdArea=0,""," OR IdArea=0") & ")"
    else
      nlim = clng(NumPreg(tipo))
      dim operador
      if tipo=0 then
        operador="<"
      elseif tipo=1 then
        operador="="
      else
        operador=">"
      end if
      sq = sq & " AND Nivel" & operador & Nivel
    end if
    sq = sq & " ORDER BY IdCompetenciaPregunta"
    disponibles = getBDlist("IdCompetenciaPregunta",sq,false)
    arrDis = split(disponibles,",")
    ndis = ubound(arrDis)-lbound(arrDis)+1
    if nlim>ndis then nlim=ndis 'Tope al numero de disponibles
    '-- Preseleccionadas: las disponibles que ya estan en <original>
    presel=""
    if disponibles<>"" and original<>"" then
      sq="SELECT IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta IN ("&disponibles&") AND IdCompetenciaPregunta IN ("&original&") ORDER BY IdCompetenciaPregunta"
      presel=getBDlist("IdCompetenciaPregunta",sq,false)
    end if
    arrSel = split(presel,",")
    nsel = ubound(arrSel)-lbound(arrSel)+1
    if nsel>ndis or nsel>nlim then  'Excedente: reinicializa descartando <original>
      Redim arrSel(0)
      nsel=0
    end if
    '-- Completa <nlim> preguntas en <arrSel>
    dim i, rndnum, noesta
    Randomize
    do while nsel<nlim
      rndnum = Int(ndis * Rnd + 1)
      noesta=true
      for i=1 to nsel
        if trim(arrSel(lbound(arrSel)+i-1))=trim(arrDis(rndnum-1)) then
          noesta=false
          exit for
        end if
      next
      if noesta then
        redim preserve arrSel(nsel)
        arrSel(nsel)=arrDis(rndnum-1)
        nsel=nsel+1
      end if
    loop
    '-- Salida: <arrSel> en CSV
    presel=""
    for i=0 to nsel-1
      presel=presel&","&arrSel(i)
    next
    if presel<>"" then
      presel=mid(presel,2)  'Omite primera coma y lee de BD para ordenar
      sq="SELECT IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta IN ("&presel&") ORDER BY Nivel, IdCompetenciaPregunta"
      presel=getBDlist("IdCompetenciaPregunta",sq,false)
    end if
    listaAleatoria=presel
  end function

  function getEntrevistaAleatoria(lispreg)
    dim lista
    if ecEntrevistaPorArea() then
      lista = listaAleatoria(-1,lispreg)
    else
      lista = strAdd( listaAleatoria(0,lispreg), ",", strAdd( listaAleatoria(1,lispreg), ",", listaAleatoria(2,lispreg) )  )
    end if
    getEntrevistaAleatoria=lista
  end function

  sub updatePreguntas(conn,lispreg)
    dim lista : lista = ""
    dim sq : sq = "DELETE FROM ECpuestoCompetenciaPregunta WHERE IdPuesto=" & IdPerfil & _
                  " AND IdCompetenciaPregunta IN (SELECT IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetencia=" & IdCompetencia & ")"
    if ecEntrevistaPorPerfil() then
      lista = getEntrevistaAleatoria(lispreg)
      if lista<>"" then sq=sq&" AND IdCompetenciaPregunta NOT IN ("&lista&")"
    end if
    conn.execute sq
    if ecEntrevistaPorPerfil() AND lista<>"" then
      conn.execute "INSERT INTO ECpuestoCompetenciaPregunta (IdPuesto,IdCompetenciaPregunta)" & _
                  " SELECT (" & IdPerfil & ") AS IdPuesto, IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta IN (" & lista & ")" & _
                  " AND IdCompetenciaPregunta NOT IN (SELECT IdCompetenciaPregunta FROM ECPuestoCompetenciaPregunta aux WHERE aux.IdPuesto=" & IdPerfil & ")"
      'dim arrP,i
      'sq="SELECT IdCompetenciaPregunta FROM ECcompetenciaPregunta WHERE IdCompetenciaPregunta IN ("&lista&") AND IdCompetenciaPregunta NOT IN (SELECT IdCompetenciaPregunta FROM ECPuestoCompetenciaPregunta WHERE IdPuesto="&IdPerfil&")"
      'lista=getBDlist("IdCompetenciaPregunta",sq,false)
      'arrP=split(lista,",")
      'for i=lbound(arrP) to ubound(arrP)
      '  conn.execute "INSERT INTO ECpuestoCompetenciaPregunta (IdPuesto,IdCompetenciaPregunta) VALUES ("&IdPerfil&","&arrP(i)&")"
      'next
    end if
  end sub

  sub updatePerfil(conn,liscom,lisniv,lispes,lisnp,lispreg)
    dim sql,i,j,aux
    dim arrId,arrVal,arrNP,arrAux
    dim porAreaCargo : porAreaCargo = ecEntrevistaPorArea()
    sql="DELETE FROM ECPuestoCompetenciaPregunta WHERE IdPuesto="&IdPerfil
    if liscom<>"" then sql=sql&" AND IdCompetenciaPregunta IN (SELECT IdCompetenciaPregunta FROM ECCompetenciaPregunta WHERE IdCompetencia NOT IN ("&liscom&"))"
    conn.execute(sql)
    sql="DELETE FROM ECPuestoCompetenciaNivel WHERE IdPuesto="&IdPerfil
    if liscom<>"" then sql=sql&" AND IdCompetencia NOT IN ("&liscom&")"
    conn.execute(sql)
    sql=""
    arrId = split(liscom,",")
    arrVal = split(lisniv,",")
    arrPes = split(lispes,",")
    arrNP = split(lisnp,"|")
    arrLP = split(lispreg,"|")
    lispaux=""
    for i=lbound(arrLP) to ubound(arrLP)
      lispaux = strAdd(lispaux, ",", arrLP(i))
    next
    for i = lbound(arrId) to ubound(arrId)
      IdCompetencia = arrId(i)
      Nivel = arrVal(i)
      Peso = arrPes(i)
      if porAreaCargo then
        NumPreg(0) = 0
        if i<=ubound(arrNP) then
          NumPreg(1) = clng(arrNP(i))
        else
          NumPreg(1) = 0
        end if
        NumPreg(2) = 0
      else
        if i<=ubound(arrNP) then
          aux=arrNP(i)
        else
          aux="0,0,0"
        end if
        arrAux=split(aux,",")
        NumPreg(0)=arrAux(lbound(arrAux))
        NumPreg(1)=arrAux(lbound(arrAux)+1)
        NumPreg(2)=arrAux(lbound(arrAux)+2)
      end if
      update conn
      updatePreguntas conn, lispaux
    next
    logAcceso LOG_CAMBIO, lblKHOR_PerfilDeEntrevistaPorCompetencias, getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
  end sub

  sub delete(conn)
    conn.execute "DELETE FROM ECPuestoCompetenciaPregunta WHERE IdPuesto="&IdPerfil&" AND IdCompetenciaPregunta IN (SELECT IdCompetenciaPregunta FROM ECCompetenciaPregunta WHERE IdCompetencia="&IdCompetencia&");"
    conn.execute "DELETE FROM ECPuestoCompetenciaNivel WHERE IdPuesto="&IdPerfil&" AND IdCompetencia="&Idcompetencia&";"
    logAcceso LOG_BAJA, lblKHOR_PerfilDeEntrevistaPorCompetencias, getDescripcion("vPerfil","IdPerfil","Perfil",idperfil)
  end sub

  sub insert(conn)
    conn.execute "INSERT INTO ECPuestoCompetenciaNivel(IdPuesto,IdCompetencia,Nivel,Peso,npInferior,npNivel,npSuperior)" & _
                " VALUES ("&IdPerfil&","&IdCompetencia&","&Nivel&","&Peso&","&NumPreg(0)&","&NumPreg(1)&","&NumPreg(2)&")"
  end sub

  sub update(conn)
    if bdExists("SELECT * FROM ECPuestoCompetenciaNivel WHERE IdPuesto="&IdPerfil&" AND IdCompetencia="&IdCompetencia) then
      conn.execute "UPDATE ECPuestoCompetenciaNivel SET Nivel="&Nivel&",Peso="&Peso&",npInferior="&NumPreg(0)&",npNivel="&NumPreg(1)&",npSuperior="&NumPreg(2) & _
                  " WHERE IdPuesto="&IdPerfil&" AND IdCompetencia="&IdCompetencia
    else
      insert conn
    end if
  end sub

  sub getfromRS(rs)
    IdPerfil=rsNum(rs,"IdPerfil")
    IdCompetencia=rsNum(rs,"IdCompetencia")
    Nivel=rsNum(rs,"Nivel")
    Peso=rsNum(rs,"Peso")
    NumPreg(0)=rsNum(rs,"npInferior")
    NumPreg(1)=rsNum(rs,"npNivel")
    NumPreg(2)=rsNum(rs,"npSuperior")
  end sub

  function getfromDB(conn,idp,idc)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM ecPerfilCompetencia WHERE IdPerfil=" & idp & " AND IdCompetencia=" & idc
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    dim i
    myIdPerfil=0
    IdCompetencia=0
    Nivel=0
    Peso=0
    for i=0 to 2
      NumPreg(i)=0
    next
    myIdArea=0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- COMPETENCIA CONDUCTA --------------------------------------------
'' ----------------------------------------------------------------------------
class clsConducta
  public IdConducta
  public IdCompetencia
  public Conducta
  public Peso
  public Nivel
  public IdComportamiento

  function catBorrable()
    catBorrable = (bdNumReferenciasFilter(IdConducta,"entidadevaluacion360","identidad2","TipoEvaluacion=3") = 0) _
                  AND NOT bdExistenReferencias(IdConducta,"PuestoConducta360","IdConducta")
  end function
  
  public property get numCursos()
    numCursos=getBD("cuantos","SELECT COUNT(*) AS cuantos FROM ConductaCurso360 WHERE IdConducta="&IdConducta)
  end property

  function delete(conn)
    conn.execute "DELETE FROM ConductaCurso360 WHERE IdConducta=" & IdConducta & "; "
    conn.execute "DELETE FROM ConductaNivel360 WHERE IdConducta=" & IdConducta & "; "
    conn.execute "DELETE FROM CatConductasObs360 WHERE IdConducta=" & IdConducta & "; "
    logAcceso LOG_BAJA, lblKHOR_ConductaObservable, left(Conducta,245) & ".."
  end function

  sub insert(conn)
    dim sq
    sq = "INSERT INTO CatConductasObs360 (Conducta,IdCompetencia,Peso,Niveles,IdComportamiento) VALUES ('" & sqsf(Conducta) & "'," & IdCompetencia & ",0," & Nivel & "," & IdComportamiento & ");"
    conn.execute sq
    sq = "SELECT MAX(IdConducta) AS lastid FROM CatConductasObs360 WHERE IdCompetencia="&IdCompetencia&" AND Conducta='"&sqsf(Conducta)&"';"
    IdConducta=getBD("lastid",sq)
    logAcceso LOG_ALTA, lblKHOR_ConductaObservable, left(Conducta,245) & ".."
  end sub

  sub update(conn)
    dim sq
    sq = "UPDATE CatConductasObs360 SET Conducta='" & sqsf(Conducta) & "', Niveles=" & Nivel & ", IdComportamiento = " & IdComportamiento & _
        " WHERE IdConducta=" & IdConducta
    conn.execute (sq)
    logAcceso LOG_CAMBIO, lblKHOR_ConductaObservable, left(Conducta,245) & ".."
  end sub

  function chkUname(conn)
    dim sq, rs
    sq = "SELECT * FROM CatConductasObs360 WHERE Conducta='" & sqsf(Conducta) & "' AND IdCompetencia=" & IdCompetencia
    set rs = getrs(conn,sq)
    if rs.EOF then
      chkUname = true
    else
      chkUname = (rs("IdConducta")&""=IdConducta&"")
    end if
    rs.close
    set rs = nothing
  end function

  sub getfromRS(rs)
    IdConducta=rsNum(rs,"IdConducta")
    IdCompetencia=rsNum(rs,"IdCompetencia")
    Conducta=rsStr(rs,"Conducta")
    Peso=rsNum(rs,"Peso")
    Nivel=rsNum(rs,"Niveles")
    IdComportamiento = rsStr(rs,"IdComportamiento")
  end sub

  function getfromDB(conn,id)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM CatConductasObs360 WHERE IdConducta=" & id
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  sub clean()
    dim i
    IdConducta=0
    IdCompetencia=0
    Peso=0
    Conducta=""
    Nivel=0
    IdComportamiento=0
  end sub

  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class

'' ----------------------------------------------------------------------------
'' ---------- CompetenciaNivelGrupo -------------------------------------------
'' ----------------------------------------------------------------------------

class CompetenciaNivelGrupo
  public Secuencia
  public Titulo
  public niveles
  public maxDiferencia
  public col
  function getCompetenciasPersona(idpersona,IdPuesto,soloDelPuesto)
    dim escenario : escenario = khorEscenarioComportamiento(IdPuesto)
    dim sq, rsc
    col.clean
    if soloDelPuesto then
      sq = "SELECT vCompetencia.IdCompetencia, vCompetencia.Competencia, vCompetencia.Definicion" & _
            " FROM vCompetencia" & _
            " INNER JOIN PuestoPsicometria ON vCompetencia.IdCompetencia=PuestoPsicometria.IdCompetencia" & _
            " INNER JOIN PersonalCompetencias ON vCompetencia.IdCompetencia=PersonalCompetencias.IdCompetencia" & _
            " WHERE (vCompetencia.Inactiva=0 OR vCompetencia.Inactiva IS NULL)" & _
            " AND PersonalCompetencias.IdPersonal=" & idpersona & _
            " AND PuestoPsicometria.IdPuesto=" & IdPuesto
      if clasificaCompetenciasXDiferencia then
        sq = sq & " AND (PuestoPsicometria.Nivel - PersonalCompetencias.Valor" & escenario & ") <= " & maxDiferencia
        if Secuencia > 1 then
          sq = sq & " AND (PuestoPsicometria.Nivel - PersonalCompetencias.Valor" & escenario & ") > " & _
                    getBDnum("maxDiferencia","SELECT maxDiferencia FROM CompetenciaNivelGrupo WHERE Secuencia =" & (Secuencia-1))
        end if
      else
        sq = sq & " AND PersonalCompetencias.Valor" & escenario & " IN (" & Niveles & ")"
      end if
    else
      sq = "SELECT vCompetencia.IdCompetencia, vCompetencia.Competencia, vCompetencia.Definicion" & _
            " FROM vCompetencia" & _
            " INNER JOIN PersonalCompetencias ON vCompetencia.IdCompetencia=PersonalCompetencias.IdCompetencia" & _
            " WHERE (vCompetencia.Inactiva=0 OR vCompetencia.Inactiva IS NULL)" & _
            " AND PersonalCompetencias.IdPersonal=" & idpersona & _
            " AND PersonalCompetencias.Valor" & escenario & " IN (" & Niveles & ")"
    end if
    sq = sq & " ORDER BY PersonalCompetencias.Valor" & escenario & " DESC, Competencia"
    set rsc = getrs(conn,sq)
    while not rsc.eof
      col.addKeyDesc rsNum(rsc,"IdCompetencia"), rsStr(rsc,"Competencia"), rsStr(rsc,"Definicion")
      rsc.movenext
    wend
    rsc.close
    set rsc = nothing
    getCompetenciasPersona = col.count
  end function
  sub getFromRS(rs)
    Secuencia = rsNum(rs,"Secuencia")
    Titulo = rsStrLocale(rs,"Titulo")
    niveles = rsStr(rs,"listaNiveles")
    maxDiferencia = rsNum(rs,"maxDiferencia")
  end sub
  sub clean()
    Secuencia = 0
    Titulo = ""
    niveles = ""
    maxDiferencia = 0
  end sub
  private sub class_initialize()
    set col = new frsCollection
    clean
  end sub
  private sub class_terminate()
    col.clean
    set col = nothing
  end sub
end class

'' ----------------------------------------------------------------------------
'' ----------------------------------------------------------------------------

%>