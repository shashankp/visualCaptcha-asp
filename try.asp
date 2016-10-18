<%@ Language="VBScript" %>
<%
For Each item In Request.Form
    Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
Next    
%>