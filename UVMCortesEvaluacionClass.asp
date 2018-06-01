
<%
'-Tipo de Corte de Evaluacion
CONST DIMENSION_ACTUALIZACION_PROFESIONAL = "6 , 21 , 33 , 44 , 55 , 67"
CONST DIMENSION_CAPACITACION_UVM = "5 , 20 , 32 , 43 , 54 , 66"
CONST DIMENSION_EVALUACION_COORDINADOR = "3 , 19 , 31 , 41 , 53 , 64"
CONST DIMENSION_OBSERVACION_CLASE = "42 , 65"
CONST DIMENSION_OPINION_ESTUDIANTIL = "1 , 17 , 29 , 39 , 51 , 62"
CONST DIMENSION_AUTOEVALUACION = "2 , 18 , 30 , 40 , 52 , 63"
CONST DIMENSION_CALIFICACION_BANNER = "9 , 24 , 36 , 47 , 58 , 70"
CONST DIMENSION_APLICACION_EXAMENES = "8 , 23 , 35 , 46 , 57 , 69"
CONST DIMENSION_TUTORIA_ASESORIA = "11 , 72"
CONST DIMENSION_ASIDUIDAD = "7 , 22 , 34 , 45 , 56 , 68"
CONST DIMENSION_ASISTENCIA_REUNIONES = "10 , 25 , 48 , 59 , 71"
CONST DIMENSION_CONGRESOS = "15"
CONST DIMENSION_DISENO_EXAMENES = "74"
CONST DIMENSION_ENTREGA_EXAMENES = "13 , 27 , 38 , 50 , 61"
CONST DIMENSION_PARTICIPACION_ACADEMIAS =" 16 , 28 , 75"
CONST DIMENSION_PLAN_CLASE = "73"
CONST DIMENSION_PLANEACION_DIDACTICA = "37"
CONST DIMENSION_REPORTE_REVISTAS = "14"
CONST DIMENSION_SYLLABUS_LLENADO = "12 , 26 , 49 , 60"

'Modo de Corte de Evaluacion
CONST CORTEEVA_GLOBAL = 1
CONST CORTEEVA_ESPECIFICO = 2
CONST CORTEEVA_GENERAL = 3

class corteEvaluacion
  public IdCorteEvaluacion
  public IdPeriodo
  public IdNivel
  public CorteEvaluacion
  public FechaIni
  public FechaFin
  public Tipo
  public Activo
  public Modo
  
  function descripcion()
    descripcion = formatDateDisp(FechaIni,false) & " - " & formatDateDisp(FechaFin,false)
  end function
  
  function enFechas()
    enFechas = (Date >= FechaIni) AND (Date <= FechaFin)
  end function
  
  function catBorrable()
    'catBorrable = NOT bdExistenReferencias(IdEvaluacion,"edc_grupo","IdEvaluacion")
    catBorrable = true
  end function

  function update(conn)
    dim sq, sqnom
    dim logAccion
    dim idp
    sqnom = "SELECT * FROM UVM_CorteEvaluacion WHERE IdPeriodo = " & IdPeriodo &" AND IdNivel = " & IdNivel & " AND Tipo = '" & sqsf(Tipo) & "'"
    
    idp = getBDnum("IdCorteEvaluacion", sqnom)
    idModo = getBDnum("Modo", sqnom)

    if (idp<>0) AND (IdCorteEvaluacion <> idp) AND (Modo = 2) AND (idModo=2) then 'EDIT
      update = false
    else
      update = true
      if IdCorteEvaluacion=0 and idp = 0 then
        logAccion = LOG_ALTA
        sq = "INSERT INTO UVM_CorteEvaluacion (IdPeriodo,IdNivel,CorteEvaluacion,FechaIni,FechaFin,Tipo,Activo,Modo" & _
              ") VALUES (" & IdPeriodo & "," & IdNivel & ",'" & sqsf(CorteEvaluacion) & "', "  & formatDateSQL(FechaIni,false) & ", " & formatDateSQL(FechaFin,false) & ",'" & sqsf(Tipo) & "', " & Activo & ", "& Modo&_
              "); "
        conn. execute sq
        IdCorteEvaluacion = getBDnum("IdCorteEvaluacion", sqnom)
      else
        if (Modo = 2 and idModo = 2) OR (Modo = 2 AND idModo = 3) OR (Modo = 3 AND idModo = 3) OR (Modo = 1 AND idModo = 1) OR (idModo = 0 AND (Modo = 1 OR Modo = 2 OR Modo = 3)) then
          modoAux =true
        else
          modoAux = false
        end if
        IdCorteEvaluacion = iif( IdCorteEvaluacion = 0 and idp <> 0 , idp, iif (Modo = 3,idp,IdCorteEvaluacion) )
        if modoAux then 
          sq = "UPDATE UVM_CorteEvaluacion SET CorteEvaluacion = '" & sqsf(CorteEvaluacion) & "'" & _
            ", IdPeriodo = " & IdPeriodo & _
            ", IdNivel = " & IdNivel & _
            ", FechaIni = " & formatDateSQL(FechaIni,false) & _
            ", FechaFin = " & formatDateSQL(FechaFin,false) & _
            ", Tipo = '" & sqsf(Tipo) & "'" & _
            ", Activo = " & Activo & _  
            ", Modo = " & Modo & _
            " WHERE IdCorteEvaluacion = " & IdCorteEvaluacion
            
            logAccion = LOG_CAMBIO
            conn.execute sq
        end if
      end if
      'logAcceso logAccion, lblED_EvaluacionAlDesempenio &":"& lblED_PeriodoDeEvaluacion, CorteEvaluacion &" ("& IdCorteEvaluacion &")"
    end if
  end function

  sub delete(conn)
    conn.execute "DELETE FROM UVM_CorteEvaluacion WHERE IdCorteEvaluacion=" & IdCorteEvaluacion
   ' logAcceso LOG_BAJA, lblED_EvaluacionAlDesempenio &":"& lblED_PeriodoDeEvaluacion, CorteEvaluacion &" ("& IdCorteEvaluacion &")"
  end sub

  sub getFromRS(rs)
    IdCorteEvaluacion = rsNum(rs,"IdCorteEvaluacion")
    IdPeriodo = rsNum(rs,"IdPeriodo")
    IdNivel = rsNum(rs,"IdNivel")
    CorteEvaluacion = rsStr(rs,"CorteEvaluacion")
    FechaIni = rs("FechaIni")
    FechaFin = rs("FechaFin")
    Tipo = rsStr(rs,"Tipo")
    Activo = rsNum(rs,"Activo")
    Modo = rsNum(rs,"Modo")
  end sub

  function getFromDB(conn,id)
    dim sq, rs
    getfromDB=false
    sq = "SELECT * FROM UVM_CorteEvaluacion WHERE IdCorteEvaluacion=" & id
    set rs = getrs(conn,sq)
    if not rs.EOF then
      getfromRS(rs)
      getfromDB=true
    end if
    rs.close
    set rs = nothing
  end function

  function getfromDBfilter(conn,filter)
    dim sq, rs
    getfromDBfilter=false
    if filter<>"" then
      sq = "SELECT * FROM UVM_CorteEvaluacion WHERE " & filter
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
    IdCorteEvaluacion = 0
    Activo = 1
    FechaIni = null
    IdPeriodo = 0
    IdNivel = 0
    CorteEvaluacion = ""
    FechaFin = null
    Tipo = ""
    Activo = 0
    Modo = 0
  end sub
  private sub class_initialize()
    clean
  end sub
  private sub class_terminate()
  end sub
end class
%>