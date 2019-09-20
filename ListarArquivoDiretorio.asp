<%
'https://github.com/fernando-nishino/ClassicASP.Functions

Sub Explorar(ByVal pasta, ByVal nos)
	Response.Write String(nos, ">") & pasta.Name & "<BR>"
	For Each subpasta In pasta.SubFolders
		Explorar subpasta, nos + 1
	Next
	For Each arquivos In pasta.Files
		Response.Write String(nos + 1, ">") & "<a href=""" & arquivos & """>" & arquivos.Name & "</a> (" & arquivos.Size & "b)<BR>"
	Next
End Sub

caminho = Server.MapPath("/")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set pasta2 = fso.GetFolder(caminho)
Explorar pasta2, 0
Set pasta2 = Nothing
Set fso = Nothing
%>