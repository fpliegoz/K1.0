<!--#include file="./khorClass.asp"-->
<!--#include file="./lmsClass.asp"-->
<%
childwin = (1=1)
thispageid = "frm_scormResultsExpanded"

function interpretacion(data)
  select case data
    case "id":                interpretacion = lblSCO_X_Id & " ("&data&")"
    case "score":             interpretacion = lblSCO_X_Score & " ("&data&")"
    case "raw":               interpretacion = lblSCO_X_Raw & " ("&data&")"
    case "max":               interpretacion = lblSCO_X_Max & " ("&data&")"
    case "min":               interpretacion = lblSCO_X_Min & " ("&data&")"
    case "status":            interpretacion = lblSCO_X_Status & " ("&data&")"
    case "interactions"       interpretacion = lblSCO_X_Interactions & " ("&data&")"
    case "objectives":        interpretacion = lblSCO_X_Objectives & " ("&data&")"
    case "time":              interpretacion = lblSCO_X_Time & " ("&data&")"
    case "type":              interpretacion = lblSCO_X_Type & " ("&data&")"
    case "correct_responses": interpretacion = lblSCO_X_CorrectResponses & " ("&data&")"
    case "pattern":           interpretacion = lblSCO_X_Pattern & " ("&data&")"
    case "weighting":         interpretacion = lblSCO_X_Weighting & " ("&data&")"
    case "student_response":  interpretacion = lblSCO_X_StudentResponse & " ("&data&")"
    case "result":            interpretacion = lblSCO_X_Result & " ("&data&")"
    case "latency":           interpretacion = lblSCO_X_Latency & " ("&data&")"
    case else:                interpretacion = data
  end select
end function

sub addElement(element,value,node)
  set newNode = xmlData.createElement("row")
  set cellNode = xmlData.createElement("cell")
  cellNode.Text = element
  newNode.appendChild cellNode
  set cellNode = xmlData.createElement("cell")
  if not IsNull(value) then
    cellNode.Text = value
  else
    newNode.setAttribute "open", "1"
  end if
  newNode.appendChild cellNode
  node.appendChild newNode
end sub

sub addExcelRow(level,element,value,fillCells)%>
  <tr>
    <td style="padding-left:<%=level*2%>em;"><%=element%></td>
    <td style="text-align:right;"><%if not IsNull(value) then%><%=value%><%end if%></td><%
    if fillCells then%>
    <td></td><td></td><td></td><td></td><%
    end if%>
  </tr><%
end sub

sub buildTable(rs, fillCells)
  dim marginLevel, nBuffer, token
  marginLevel = 1
  nBuffer = ""
  Dim mData : Set mData = New AssocArrayClass
  while not rs.EOF
  token = split(rsStr(rs,"Element"),".")
  for i=2 to UBound(token)
    marginLevel = i-1
    select case i
      case 2:
        if nBuffer <> CStr(token(i)) then
          nBuffer = CStr(token(i))
          mData(nBuffer)("scoreAlreadySetted") = false
          mData(nBuffer)("objectivesAlreadySetted") = false
          mData(nBuffer)("correct_responsesAlreadySetted") = false
          addExcelRow marginLevel, token(i), null, fillCells
        end if
      case 3:
        if token(i) <> "score" and token(i) <> "objectives" and token(i) <> "correct_responses" then
          addExcelRow marginLevel, interpretacion(token(i)), interpretacion(rsStr(rs,"Value")), fillCells
        else
          if not mData(nBuffer)(token(i)&"AlreadySetted") then
            addExcelRow marginLevel, interpretacion(token(i)), null, fillCells
            mData(nBuffer)(token(i)&"AlreadySetted") = true
          end if
        end if
      case 4:
        if token(i) = "min" or token(i) = "max" or token(i) = "raw" then
          addExcelRow marginLevel, interpretacion(token(i)), interpretacion(rsStr(rs,"Value")), fillCells
        else
          addExcelRow marginLevel, interpretacion(token(i)), null, fillCells
        end if
      case 5:
        addExcelRow marginLevel, interpretacion(token(i)), interpretacion(rsStr(rs,"Value")), fillCells
    end select
  next
  rs.movenext
  wend
end sub

Dim counterData : Set counterData = New AssocArrayClass
thispage = "scormResultsExpanded.asp"
target = reqplain("target")
report_type = reqplain("report_type")
IdPersona = reqn("idper")
IdSCO = reqn("idsco")
attempt = reqn("att")

if report_type = "pdf" then
  sesionFromRequest
end if

sql = "SELECT Element, Value FROM ScormSCOTrack WHERE Element LIKE 'cmi."&target&".%' AND IdPersonal=" & IdPersona & " AND IdScormSCO=" & IdSCO & " AND Attempt=" & attempt & " ORDER BY cast(substring(Element, len('cmi.interactions.')+1, charindex('.', Element, len('cmi.interactions.')+1) - (len('cmi.interactions.')+1)) as int) "
set rs = getrs(conn,sql)

if report_type = "html" or report_type = "pdf" then
  if report_type = "html" then
    'inicio armado de cadena xml
    xmlString = "<rows><row open='1'><cell>"&interpretacion(target)&"</cell><cell></cell></row></rows>"
    set xmlData = Server.CreateObject("Microsoft.XMLDOM")
    xmlData.async = false
    xmlData.loadXML(xmlString)
    set root = xmlData.documentElement
    
    if target = "objectives" then
      nBuffer = ""
      scoreAlreadySetted = false
      rowCounter = 1
      while not rs.EOF
        token = split(rsStr(rs,"Element"),".")
        for i=2 to UBound(token)
          select case i
            case 2:
              if nBuffer <> CStr(token(i)) then
                nBuffer = CStr(token(i))
                scoreAlreadySetted = false
                rowCounter = rowCounter + 1
                addElement nBuffer,null,root.childNodes(0)
              end if
            case 3:
              set tmpNode = root.childNodes(0)
              for j=0 to tmpNode.childNodes.length-1
                set node = root.childNodes(0).childNodes(j)
                if node.nodeName = "row" then
                  if CStr(node.childNodes(0).text) = nBuffer then
                    exit for
                  end if
                end if
              next
              if token(i) = "score" then
                if not scoreAlreadySetted then
                  scoreAlreadySetted = true
                  addElement interpretacion(token(i)),null,node
                  rowCounter = rowCounter + 1
                end if
              else
                addElement interpretacion(token(i)),rsStr(rs,"Value"),node
                rowCounter = rowCounter + 1
              end if
            case else:
              breakAll = false
              for j=0 to root.childNodes(0).childNodes.length-1
                set node = root.childNodes(0).childNodes(j)
                if node.nodeName = "row" then
                  if CStr(node.childNodes(0).text) = nBuffer then
                    for k=0 to node.childNodes.length-1
                      if node.childNodes(k).nodeName = "row" then
                        set tmpNode = node.childNodes(k)
                        if tmpNode.childNodes(0).text = interpretacion(token(i-1)) then                    
                          breakAll = true
                          set node = tmpNode
                          exit for
                        end if
                      end if
                    next
                  end if
                end if
                if breakAll then
                  exit for
                end if
              next
              addElement interpretacion(token(i)),rsStr(rs,"Value"),node
              rowCounter = rowCounter + 1
          end select
        next
        rs.movenext
      wend
    elseif target = "interactions" then
      Dim nData : Set nData = New AssocArrayClass
      nBuffer = ""
      rowCounter = 1
      while not rs.EOF
        token = split(rsStr(rs,"Element"),".")
        for i=2 to UBound(token)
          select case i
            case 2:
              if nBuffer <> CStr(token(i)) then
                nBuffer = CStr(token(i))
                nData(nBuffer)("objecSetted") = false
                nData(nBuffer)("cRespSetted") = false
                rowCounter = rowCounter + 1
                addElement nBuffer,null,root.childNodes(0)
              end if
            case 3:
              for j=0 to root.childNodes(0).childNodes.length-1
                set node = root.childNodes(0).childNodes(j)
                if node.nodeName = "row" then
                  if CStr(node.childNodes(0).text) = nBuffer then
                    exit for
                  end if
                end if
              next
              if token(i) = "objectives" then
                if nData(nBuffer)("objecSetted") = false then
                  nData(nBuffer)("objecSetted") = true
                  addElement interpretacion(token(i)),null,node
                  rowCounter = rowCounter + 1
                end if
              end if
              if token(i) = "correct_responses" then
                if nData(nBuffer)("cRespSetted") = false then
                  nData(nBuffer)("cRespSetted") = true
                  addElement interpretacion(token(i)),null,node
                  rowCounter = rowCounter + 1
                end if
              end if
              if token(i) <> "objectives" and token(i) <> "correct_responses" then
                addElement interpretacion(token(i)),rsStr(rs,"Value"),node
                rowCounter = rowCounter + 1
              end if
            case 4:
              breakAll = false
              for j=0 to root.childNodes(0).childNodes.length-1
                set node = root.childNodes(0).childNodes(j)
                if node.nodeName = "row" then
                  if CStr(node.childNodes(0).text) = nBuffer then
                    for k=0 to node.childNodes.length-1
                      if node.childNodes(k).nodeName = "row" then
                        set tmpNode = node.childNodes(k)
                        if tmpNode.childNodes(0).text = interpretacion(token(i-1)) then   
                          breakAll = true
                          set node = tmpNode
                          exit for
                        end if
                      end if
                    next
                  end if
                end if
                if breakAll then
                  exit for
                end if
              next
              if IsObject(nData(nBuffer)(token(i-1))(CStr(token(i)))) then
                addElement interpretacion(CStr(token(i))),null,node
                nData(nBuffer)(token(i-1))(CStr(token(i))) = CStr(token(i))
                rowCounter = rowCounter + 1
              end if
            case 5:
              breakAll = false
              for j=0 to root.childNodes(0).childNodes.length-1
                set node = root.childNodes(0).childNodes(j)
                if node.nodeName = "row" then
                  if CStr(node.childNodes(0).text) = nBuffer then
                    for k=0 to node.childNodes.length-1
                      if node.childNodes(k).nodeName = "row" then
                        set tmpNode = node.childNodes(k)
                        if tmpNode.childNodes(0).text = interpretacion(token(i-2)) then
                          for n=0 to tmpNode.childNodes.length-1
                            set tmpN = tmpNode.childNodes(n)
                            if CStr(tmpN.text) = CStr(token(i-1)) then
                              breakAll = true
                              set node = tmpN
                              exit for
                            end if
                          next
                        end if
                      end if
                      if breakAll then
                        exit for
                      end if
                    next
                  end if
                end if
                if breakAll then
                  exit for
                end if
              next
              addElement interpretacion(token(i)),rsStr(rs,"Value"),node
              rowCounter = rowCounter + 1
          end select
        next
        rs.movenext
      wend
    end if
    xmlString = left(xmlData.xml,len(xmlData.xml)-2)
    'fin armado cadena xml
  end if
  tit1 = lblSCO_SCORMResults
%>
<HTML>
<HEAD>
<TITLE><%=khorAppName()%> - <%=tit1%></TITLE>
<META http-equiv="Content-Type" content="text/html; charset=<%=KHOR_CHARSET%>">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">

<link rel="STYLESHEET" type="text/css" href="./dhtmlx/css/dhtmlxgrid.css">
<link rel="STYLESHEET" type="text/css" href="./dhtmlx/css/dhtmlxgrid_dhx_skyblue.css">
<link rel="STYLESHEET" type="text/css" href="./customEstilos.css">
<script src="./dhtmlx/js/dhtmlxcommon.js"></script>
<script src="./dhtmlx/js/dhtmlxgrid.js"></script>
<script src="./dhtmlx/js/dhtmlxgridcell.js"></script>
<script src="./dhtmlx/js/dhtmlxtreegrid.js"></script>
<% includeJS %>
</head>
  <BODY>
    <!--#include file="khorHeader.asp"-->
    <table class="pagetable" cellspacing="5" width="<%=khorWinWidth()%>">
      <tr>
        <td class="pagetitle">
          <% call ponEncabezado(tit1, tit2, "") %>
        </td>
      </tr>
      <tr>
        <td class="bordegris">
          <%IF errmsg<>"" THEN%><div class="alerta"><%=errmsg%></div><br class="sp5"><%END IF%>
          <form name="TrueForm" action="<%=thispage%>" method="POST" class="plana" onSubmit="return false;">
            <%ponLigaBottom true%>
            <br/>
            <table align="center" border="0" width="95%" cellspacing="0">
              <tr>
                <td style="font-size: 10pt;">
                  <span style="font-family:Verdana,Arial,Helvetica,sans-serif;font-weight: bold;"><%=lblSCO_SCO%>:</span> <%=getBD("Title","SELECT Title from ScormSCO WHERE IdScormSCO=" & IdSCO)%><br/>
                  <span style="font-family:Verdana,Arial,Helvetica,sans-serif;font-weight: bold;"><%=lblLMS_Student%>:</span> <%=getBD("nombre","SELECT nombre from personal WHERE idpersonal=" & IdPersona)%></br>
                  <span style="font-family:Verdana,Arial,Helvetica,sans-serif;font-weight: bold;"><%=lblSCO_Attempt%>:</span> <%=attempt%>
                </td>
              </tr><%
            if report_type = "html" then%>
                <td valign="middle" align="right"><%=lblFRS_Exportar%></td>
                <td align="right" width="55px">
                  <a href="javascript:void(0);" onClick="myPrintPage('xls')"><%=lblLMS_Excel%>&nbsp;<img src="./khorImg/lms/icons/5.gif" border="0" /></a><br />
                  <a href="javascript:void(0);" onClick="myPrintPage('pdf')"><%=lblLMS_PDF%>&nbsp;<img src="./khorImg/lms/icons/1.gif" border="0" /></a>
                </td>
              </tr>
            </table>
            <br/><%
            elseif report_type = "pdf" then%>
              <tr>
                <td align="center" style="vertical-align:top;">
                  <table style="">
                    <tr class="celdatit">
                      <td><%=lblLMS_Segment%></td>
                      <td><%=lblLMS_Value%></td>
                    </tr>
                    <tr>
                      <td colspan="2"><%=interpretacion(target)%></td>
                    </tr><%
                    buildTable rs, false%>
                  </table>
                </td>
              </tr>
            </table><%
        end if
    if report_type = "html" then%>
            <div id="div_TreeGrid" align="center" style="width:570px;height:350px;background-color:white;"></div>
            <script>
              $(function(){init();})
              function init(){
                mygrid = new dhtmlXGridObject('div_TreeGrid');
                mygrid.selMultiRows = true;
                mygrid.imgURL = "./dhtmlx/imgs/icons_greenfolders/";
                mygrid.setHeader("Segmento,Valor");
                mygrid.setInitWidths("300,<%if report_type = "html" then%>253<%else%>268<%end if%>");
                mygrid.setColAlign("left,left");
                mygrid.setColTypes("tree,ro");
                mygrid.setColSorting("str,str");
                mygrid.enableTreeCellEdit(false);
                mygrid.enableMultiline = true;
                mygrid.init();
                mygrid.setSkin("dhx_skyblue");
                mygrid.loadXMLString('<%=xmlString%>');
              }
              function myPrintPage(format) {
                noPageBlocker = true;
                if(format=="pdf") {<%
                  pdfkey = initPDFurl(thispageid, pdf_URL() & thispage & "?report_type=pdf&target="& target &"&idper="& IdPersona &"&idsco="& IdSCO &"&att="& attempt)%>
                  openPDFjob('<%=pdfkey%>');
                } else if(format=="xls") {
                  var url = "<%=thispage%>?report_type="+format+"&target=<%=target%>&idper=<%=IdPersona%>&idsco=<%=IdSCO%>&att=<%=attempt%>";
                  url = "<%=pageURL(1)%>/" + url;
                  window.location = url;
                }     
              }
            </script>
            <div align="center" nowrap="nowrap">
              <%ponLigaRegreso(thispageid)%>
            </div>
            <%ponLigaTop true%>
          </form>
        </td>
      </tr>
    </table>
    <!--#include file="khorFooter.asp"-->
    <%end if%>
<%
elseif report_type = "xls" then 
response.ContentType = "application/vnd.ms-excel"
response.AddHeader "Content-Disposition", "attachment;filename=resSCORM.xls"
%>
<table>
  <tr><td colspan="6"><%=lblSCO_SCO%>: <%=getBD("Title","SELECT Title from ScormSCO WHERE IdScormSCO=" & IdSCO)%></td></tr>
  <tr><td colspan="6"><%=lblLMS_Student%>: <%=getBD("nombre","SELECT nombre from personal WHERE idpersonal=" & IdPersona)%></td></tr>
  <tr><td colspan="6"><%=lblSCO_Attempt%>: <%=attempt%></td></tr>
  <tr><td></td><td></td><td></td><td></td><td></td><td></td></tr>
  <tr>
    <td><%=lblLMS_Segment%></td>
    <td><%=lblLMS_Value%></td>
  </tr>
  <tr>
    <td><%=interpretacion(target)%></td>
    <td></td>
  </tr><%
  buildTable rs, true
end if
%>
    </table>    
  </BODY>
</HTML>
<%
  conn.close
%>