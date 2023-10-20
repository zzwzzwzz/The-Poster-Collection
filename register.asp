<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Redirect if username exists
MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="RegFail.html"
  MM_rsKeyConnection=MM_conn_STRING
  MM_dupKeyUsernameValue = CStr(Request.Form("UserID"))
  MM_dupKeySQL="SELECT UserID FROM reg WHERE UserID='" & Replace(MM_dupKeyUsernameValue,"'","''") & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_conn_STRING
  MM_editTable = "reg"
  MM_editRedirectUrl = "RegSuccess.html"
  MM_fieldsStr  = "UserID|value|Passwd|value|Email|value|QQ|value|MSN|value"
  MM_columnsStr = "UserID|',none,''|Passwd|',none,''|Email|',none,''|QQ|none,none,NULL|MSN|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Register</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 36px;
	font-weight: bold;
}
.STYLE2 {font-size: 24px}
body,td,th {
	font-family: Chiller;
	color: #FFFFFF;
}
body {
	background-image: url(images/login1.jpg);
	background-color: #333333;
	margin-top: 150px;
}
-->

</style>
<link rel="shortcut icon" href="favicon2.ico" /> 
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <table width="407" border="0" align="center" bordercolor="#333333" bgcolor="#333333">
    <tr>
      <td height="58" colspan="2" bgcolor="#333333"><div align="center" class="STYLE1">Register A New Account:</div></td>
    </tr>
    <tr>
      <td width="147" height="38" bgcolor="#333333"><div align="right" class="STYLE2">
        <div align="right">User ID:</div>
      </div></td>
      <td width="250" bgcolor="#333333"><label>
        <input name="UserID" type="text" id="UserID" size="20" maxlength="50" />
      *</label></td>
    </tr>
    <tr>
      <td height="38" bgcolor="#333333"><div align="right" class="STYLE2">Password:</div></td>
      <td bgcolor="#333333"><input name="Passwd" type="text" id="Passwd" size="20" maxlength="50" />
      *</td>
    </tr>
    <tr>
      <td height="39" bgcolor="#333333"><div align="right" class="STYLE2">Repeat: </div></td>
      <td bgcolor="#333333"><input name="repass" type="text" id="repass" size="20" maxlength="50" />
      *</td>
    </tr>
    <tr>
      <td height="43" bgcolor="#333333"><div align="right" class="STYLE2">Email:</div></td>
      <td bgcolor="#333333"><input name="Email" type="text" id="Email" size="20" maxlength="50" />
      *</td>
    </tr>
    <tr>
      <td height="40" bgcolor="#333333"><div align="right" class="STYLE2">QQ:</div></td>
      <td bgcolor="#333333"><input name="QQ" type="text" id="QQ" size="20" maxlength="50" /></td>
    </tr>
    <tr>
      <td height="39" bgcolor="#333333"><div align="right" class="STYLE2">MSN:</div></td>
      <td bgcolor="#333333"><input name="MSN" type="text" id="MSN" size="20" maxlength="50" /></td>
    </tr>
    <tr>
      <td height="44" bgcolor="#333333">&nbsp;</td>
      <td bgcolor="#333333"><label>
        <div align="justify" class="STYLE2">
          <input type="submit" name="Submit" value="Register" />
          <input name="reset" type="reset" id="reset" value="Reset" />
        </div>
      </label></td>
    </tr>
  </table>

    <input type="hidden" name="MM_insert" value="form1">
</form>

</body>
</html>
