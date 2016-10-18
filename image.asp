<%@ Language="VBScript" %>
<%
Dim imageName 
imageName = Session("images")(Int(Request.QueryString("id")))

Set stream = Server.CreateObject("ADODB.Stream")    
stream.Type = 1
stream.Open
stream.LoadFromFile(Server.MapPath("images\" + imageName + ".png"))

Response.Expires = 0 
Response.Buffer = TRUE 
Response.Clear 
Response.ContentType = "image/png"
Response.BinaryWrite(stream.Read()) 
Response.Flush
Response.End

stream = Nothing
%>
