<%
'https://github.com/fernando-nishino/ClassicASP.ListTreeFolderFile

Sub Explore(ByVal folder, ByVal node)
	Response.Write String(node, ">") & folder.Name & "<BR>"
	For Each subfolder In folder.SubFolders
		Explore subfolder, node + 1
	Next
	For Each file In folder.Files
		Response.Write String(node + 1, ">") & "<a href=""" & file & """>" & file.Name & "</a> (" & file.Size & "b)<BR>"
	Next
End Sub

path = Server.MapPath("/")
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set folderRoot = fso.GetFolder(path)
Explore folderRoot, 0
Set folderRoot = Nothing
Set fso = Nothing
%>