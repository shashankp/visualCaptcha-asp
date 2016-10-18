<%@ Language="VBScript" %>
<script language="JScript" runat="server" src='js/json2.js'></script>
<%
    
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set imagesFile = fs.OpenTextFile(Server.MapPath("images.json"),1,false)
Set imagesJson = JSON.parse(imagesFile.ReadAll)

Dim selectedImages(5)

For i = 0 To 4
    'Randomize
    'Response.Write Int(Rnd*imagesJson.length)
    'Response.Write imagesJson.[Int(Rnd*imagesJson.length)].name
    selectedImages(i) = imagesJson.shift().name
Next 

Session("images") = selectedImages

Set frontendDataFile = fs.OpenTextFile(Server.MapPath("start.json"),1,false)
Set frontendDataJson = JSON.parse(frontendDataFile.ReadAll)
Session("frontendData") = frontendDataJson

Session("validImageOption") = "3ddd5c563e533e8453bb4972611c5793391d962f"
Session("imageFieldName") = "37d251b7f67da50ee1ea56750bdd40ec9b757277"

Set stream = Server.CreateObject("ADODB.Stream")    
stream.Type = 1
stream.Open
stream.LoadFromFile(Server.MapPath("start.json"))

Response.Expires = 0 
Response.Buffer = TRUE 
Response.Clear 
Response.ContentType = "application/json"
Response.BinaryWrite(stream.Read()) 
Response.Flush
Response.End

stream = Nothing

%>
