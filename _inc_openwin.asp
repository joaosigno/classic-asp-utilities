<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
%>
<script type="text/javascript" language="javascript">
function openWin ( fileName, windowName, width, height )
{
	//window.open(fileName,windowName,'width=350,height=350,directories=no,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,resizable=no');
	var opts = "";
	if ( width ) opts += 'width=' + width + ',';
	if ( height ) opts += 'height=' + height + ',';
	return window.open(fileName,windowName,opts+'directories=no,location=no,menubar=yes,scrollbars=yes,status=no,toolbar=no,resizable=yes');
}
// example usage:
// <a href="#" onClick="return openWin('help.htm','_blank');">Help</a>
</script>