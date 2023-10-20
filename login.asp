<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("UserID"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="index.html"
  MM_redirectLoginFailed="LogFail.html"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_conn_STRING
  MM_rsUser.Source = "SELECT UserID, Passwd"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM reg WHERE UserID='" & Replace(MM_valUsername,"'","''") &"' AND Passwd='" & Replace(Request.Form("Passwd"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Login</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 36px;
	font-weight: bold;
	color: #FFFFFF;
}
body,td,th {
	font-family: Chiller;
	color: #FFFFFF;
}
body {
	background-color: #333333;
	margin-top: 150px;
	background-image: url(images/login1.jpg);
}
.STYLE4 {color: #FFFFFF; font-size: 24px; }
-->
</style>
<link rel="shortcut icon" href="favicon2.ico" /> 
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_LoginAction%>">
  <table width="371" height="292" border="0" align="center" bordercolor="#333333" bgcolor="#333333">
    <tr>
      <td height="81" colspan="2" bgcolor="#333333"><div align="center" class="STYLE1">Login</div></td>
    </tr>
    <tr>
      <td width="134" height="61" bgcolor="#333333"><div align="right" class="STYLE4">UserID:</div></td>
      <td width="215" bgcolor="#333333"><label>
        <input name="UserID" type="text" id="UserID" size="20" maxlength="50" />
      </label></td>
    </tr>
    <tr>
      <td height="58" bgcolor="#333333"><div align="right" class="STYLE4">Password:</div></td>
      <td bgcolor="#333333"><input name="Passwd" type="text" id="Passwd" size="20" maxlength="50" /></td>
    </tr>
    <tr>
      <td height="72" bgcolor="#333333">&nbsp;</td>
      <td bgcolor="#333333"><label>
        <input type="submit" name="Submit" value="Login" />
        <input name="Submit2" type="button" onclick="JavaScript:window.location='register.asp'" value="Register" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
