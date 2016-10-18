<%@ Language="VBScript" %>
<%
Set stream = Server.CreateObject("ADODB.Stream")    
stream.Type = 1
stream.Open
stream.LoadFromFile(Server.MapPath("images\airplane.png"))

Response.Expires = 0 
Response.Buffer = TRUE 
Response.Clear 
Response.ContentType = "image/png"
Response.BinaryWrite(stream.Read()) 
Response.Flush
Response.End

stream = Nothing
%>
