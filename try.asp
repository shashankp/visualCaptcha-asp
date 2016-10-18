<%@ Language="VBScript" %>
<%
If Request.Form(Session("imageFieldName")) = Session("validImageOption") Then
    Response.Write "OK"
Else
    Response.Write "Fail"
End If
%>