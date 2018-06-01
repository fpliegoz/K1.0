<!--#include file="./khorClass.asp"-->
<%
  thispageid = "ejemploPagina"
  thispage = "ejemploPagina.asp"
  
  mov = ucase(reqs("mov"))
  if mov="PDF" then
    sesionFromRequest
    childwin = true
  end if
  
  'checaSesion ses_super&","&ses_adminid, "", ""
  'ok2enter = khorPermisoModulo( Modulo_360config, khorModulosActivos())
  'validaEntrada ok2enter, "", thispageid

  x = reqs("x")
  
  tit1 = "Pagina de ejemplo"
  tit2 = "Segunda linea de titulo"
  tit3 = "Tercera linea de titulo"

  buttons = iif( mov="PDF", "", "A||accion('A');@@B||accion('B');@@Google||location.href='http://www.google.com';@@Luis Paz||abreAyuda();" )
  
%>
<% layoutHeadStart khorAppName() & " - " & strAdd(tit1, " - ", tit2) %>
<% includeJS %>
<script>
function accion(op) {
  sendval('', 'x',op);
}
<%IF pdf_enabled() then %>
function myPrintPage() {
 <%
  pdfkey = initPDFurl( thispageid , _
                        pdf_URL() & thispage & "?mov=pdf&x=" & x )
%>
  openPDFjob( '<%=pdfkey%>' );
}
<%END IF%>
</script>
<% layoutHeadEnd %>
<% layoutStart tit1, tit2, tit3, errmsg, khorWinWidth(),"" %>
<% defaultFormStart thispage,"",true %>

<input type="hidden" name="x" value="">
<input type="hidden" name="mov" value="">
Prueba de Camara de Video:
<div id="container">
  <video id="video" width="640" height="480" autoplay></video>
	<button id="snap" class="sexyButton">Snap Photo</button>
	<canvas id="canvas" width="640" height="480"></canvas>
</div>
<script>

		// Put event listeners into place
		window.addEventListener("DOMContentLoaded", function() {
			// Grab elements, create settings, etc.
            var canvas = document.getElementById('canvas');
            var context = canvas.getContext('2d');
            var video = document.getElementById('video');
            var mediaConfig =  { video: true };
            var errBack = function(e) {
            	console.log('An error has occurred!', e)
            };

			// Put video listeners into place
            if(navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
                navigator.mediaDevices.getUserMedia(mediaConfig).then(function(stream) {
                    video.src = window.URL.createObjectURL(stream);
                    video.play();
                });
            }

            /* Legacy code below! */
            else if(navigator.getUserMedia) { // Standard
				navigator.getUserMedia(mediaConfig, function(stream) {
					video.src = stream;
					video.play();
				}, errBack);
			} else if(navigator.webkitGetUserMedia) { // WebKit-prefixed
				navigator.webkitGetUserMedia(mediaConfig, function(stream){
					video.src = window.webkitURL.createObjectURL(stream);
					video.play();
				}, errBack);
			} else if(navigator.mozGetUserMedia) { // Mozilla-prefixed
				navigator.mozGetUserMedia(mediaConfig, function(stream){
					video.src = window.URL.createObjectURL(stream);
					video.play();
				}, errBack);
			}

			// Trigger photo take
			document.getElementById('snap').addEventListener('click', function() {
				context.drawImage(video, 0, 0, 640, 480);
			});
		}, false);

	</script>



<div style="page-break-before:always;">
 <%=x%>
</div>
<% defaultFormEnd buttons, "", (mov<>"PDF") %>
<% layoutEnd %>