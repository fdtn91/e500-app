' E500 3D Print — Lanzador sin ventana de consola
' Doble clic sobre este archivo para abrir la app sin terminal visible

Dim objShell, objFSO, strDir

Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' Directorio donde vive este script
strDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Cambiar al directorio de la app
objShell.CurrentDirectory = strDir

' Si no existe node_modules, instalar dependencias primero (silencioso)
If Not objFSO.FolderExists(strDir & "\node_modules") Then
    objShell.Run "cmd /c npm install --silent", 0, True
End If

' Lanzar Electron directamente usando el binario local para inicio instantáneo
If objFSO.FileExists(strDir & "\node_modules\.bin\electron.cmd") Then
    objShell.Run """" & strDir & "\node_modules\.bin\electron.cmd"" .", 0, False
Else
    objShell.Run "cmd /c npx electron .", 0, False
End If

Set objShell = Nothing
Set objFSO   = Nothing
