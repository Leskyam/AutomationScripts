Option Explicit

'*********************************************************************************************
'DECLARACIÓN DE VARIABLES LOCALES A ESTE SCRIPT.
'*********************************************************************************************
'CONSTANTES
'*********************************************************************************************
'PARA OBJETOS ESCRITURA EN FICHEROS DE TEXTO
Const ForReading = 1, ForWriting = 2, ForAppending = 8, TristateUseDefault = -2, _
      TristateTrue = -1, TristateFalse = 0

'*********************************************************************************************
'CONSTANTES DE CADENAS, GENERALMENTE PARA DAR FORMATO 
'A LO QUE SE ESCRIBE EN LOS FICHEROS DE TEXTO
'*********************************************************************************************
Public vbTab, VbCrLf
vbTab = Chr(9) 
VbCrLf = Chr(13) & Chr(10)

'*********************************************************************************************
'VARIABLES GENERALES
'*********************************************************************************************
Dim strComputerName, strUserName, strInventoryFile, strErrorsFile 

'*********************************************************************************************
'FUNCIONES UTILES
'*********************************************************************************************
'Chequea que se esté utilizando como engine para los scripts el motor CScript, sino es así,
'lo establece como el predeterminado.
'*********************************************************************************************
Sub CheckScriptHost()
Dim strMsg, strExec

On Error Resume Next

If InStr(LCase(wscript.fullname),"cscript") = 0 Then

	Dim oShell, oWSHEnv
	Set oShell=CreateObject("WScript.Shell")
	Set oWSHEnv = oShell.Environment 
	strExec = "cscript.exe //Logo //H:cscript //S"
	oShell.Run strExec,0,True

	strExec = ""
  If oWSHEnv("OS") = "Windows_NT" Then
    strExec = "cmd"
  else
    strExec = "command"
  End If

	strExec = strExec & " /c " & Chr(34) & "cscript.exe " & Chr(34) & _
	WScript.ScriptFullname & Chr(34)& Chr(34)

	oShell.Run strExec,,True

	strExec = "cscript.exe //Logo //H:WScript //S"
	oShell.Run strExec,0,False

	Wscript.Quit

End If

End Sub

'*********************************************************************************************
'Devuelve el camino desde donde se está ejecutando el script
'*********************************************************************************************
Function SetExecutingFromPath()

Dim strScriptPath

strScriptPath=Left(wscript.scriptfullname, _
Len(wscript.scriptfullname)-Len(wscript.scriptname))

If Right(strScriptPath,1) <> "\" Then
	strScriptPath=strScriptPath & "\"
End If

SetExecutingFromPath=strScriptPath

End Function

'*********************************************************************************************
'Devuelve el nombre de la máquina donde se está corriendo el script.
'*********************************************************************************************
Function GetComputerName()

Dim oNet
Set oNet = WScript.CreateObject("WScript.Network")
GetComputerName = UCase(oNet.ComputerName)
Set oNet = Nothing

End Function

'*********************************************************************************************
'Devuelve el nombre del usuario con derechos administrativos bajo el cual se ejecuta el script.
'*********************************************************************************************
Function GetUserName()

Dim oNet
Set oNet = WScript.CreateObject("WScript.Network")
GetUserName = UCase(oNet.UserName)
Set oNet = Nothing

End Function

'*********************************************************************************************
'ESTABLECER VALORES DE LAS VARIABLES GENERALES
strComputerName = GetComputerName
strInventoryFile = SetExecutingFromPath & strComputerName & ".XML"
strErrorsFile = SetExecutingFromPath & strComputerName & "_Errores.TXT"
strUserName = GetUserName

'*********************************************************************************************
'Chequea si existe el fichero, si es así lo elemina.
'*********************************************************************************************
Sub CheckIfExistsFiles()
Dim oFSO
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
'Fichero de inventario
If oFSO.FileExists(strInventoryFile) Then
	oFSO.DeleteFile(strInventoryFile)
End If
'Fichero de errores
If oFSO.FileExists(strErrorsFile) Then
	oFSO.DeleteFile(strErrorsFile)
End If

Set oFSO = Nothing

End Sub

'*********************************************************************************************
'Escribir el encabezado general del fichero de texto generado.
'*********************************************************************************************
Sub SetFileHeader()

On Error Resume Next

  WriteToInventoryFile "<?xml version='1.0' encoding='WINDOWS-1252'?>"
  WriteToInventoryFile "<Encabezado Utilidad='SCRIPT PARA INVENTARIO DE HARDWARE' Maquina='" & strComputerName & "' Fecha='" & Now & "'>"
  WScript.Echo "Utilidad='SCRIPT PARA INVENTARIO DE HARDWARE' Maquina='" & strComputerName & "' Fecha='" & Now & "'"

On Error Goto 0

End Sub

'*********************************************************************************************
'Escribir el pie general del fichero de texto generado.
'*********************************************************************************************
Sub SetFileFooter()

On Error Resume Next

  WriteToInventoryFile "</Encabezado>"

On Error Goto 0

End Sub

'*********************************************************************************************
'Escribe el encabezado de la clase que se le pase como parámetro.
'*********************************************************************************************
Sub SetClassHeader(strClassName)

On Error Resume Next

  WriteToInventoryFile "<Clase Nombre='" & strClassName & "'>"
  WScript.Echo "ENCABEZADO DE LA CLASE: '" & strClassName & "'"
  
On Error Goto 0

End Sub

'*********************************************************************************************
'Escribe el encabezado de la clase que se le pase como parámetro.
'*********************************************************************************************
Sub SetClassFooter(strClassName)

On Error Resume Next

  WriteToInventoryFile "</Clase>"
  WScript.Echo "FIN DE LA CLASE: '" & strClassName & "'"
    
On Error Goto 0

End Sub

'*********************************************************************************************
'Escribe en un fichero de error los detalles del mismo, en caso que suceda alguno.
'*********************************************************************************************
Sub LogError(strOrigen, intNumber, strDescription, strValue, strFile)

Dim oFSO, oErrorFile
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")

On Error Resume Next

Set oErrorFile = oFSO.OpenTextFile(strErrorsFile, ForAppending, True, False)
If Err <> 0 Then
  Err.Clear
Else
  oErrorFile.WriteLine "ERROR"
  oErrorFile.WriteLine vbTab & "Originado en: " & strOrigen
  oErrorFile.WriteLine vbTab & "Número: " & intNumber
  oErrorFile.WriteLine vbTab & "Descripción: " & strDescription
  oErrorFile.WriteLine vbTab & "Valor: " & strValue
  oErrorFile.WriteLine vbTab & "Fichero: " & strFile
  oErrorFile.Close
  Set oErrorFile = Nothing
  WScript.Echo "SE ESCRIBIO EL ERROR NUMERO " & intNumber & " REVISAR FICHERO: " & strErrorsFile
End If

On Error Goto 0

End Sub

'*********************************************************************************************
'Esta función reemplaza los valores del aspersan y los signos de menor que y mayor que por los
'que puedan ser interpretados por XML.
'*********************************************************************************************
Function MyXMLParser(ByVal strValue)

Dim strNewValue
strNewValue = strValue

If inStr(1,strNewValue,"&")>0 Then
  strNewValue = Replace(strNewValue,"&","&amp;",1,-1)
End If
If inStr(1,strNewValue,"<")>0 Then
  strNewValue = Replace(strNewValue,"<","&lt;",1,-1)
End If
If inStr(1,strNewValue,">")>0 Then
  strNewValue = Replace(strNewValue,">","&gt;",1,-1)
End if

MyXMLParser = strNewValue

End Function

'*********************************************************************************************
'Hace un ciclo que recorre todos los elementos de una clase que se le pase como parámetro y 
'cada una de sus propiedades o campos.
'*********************************************************************************************
Sub SetClassData(strClassName)

Dim x, strValue
Dim objWMIService, oClass, oClassItem, oCampo

On Error Resume Next

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2")
Set oClass = objWMIService.ExecQuery("SELECT * FROM " & strClassName)
x = 1
For Each oClassItem in oClass
    WriteToInventoryFile "<Elemento ID='" & x & "'>"
    For Each oCampo in oClassItem.Properties_
			If Err <> 0 Then
				LogError "Sub SetClassData. WMI Clase: " & strClassName, Err.Number, Err.Description, "", "Escribiendo en: " & strInventoryFile
				Err.Clear
			Else
			  If Not oCampo.IsArray  AND Not isNull(oCampo.value) Then
				  strValue = Trim(MyXMLParser(oCampo.value))
				  WriteToInventoryFile "<" & oCampo.name & ">" & strValue & "</" & oCampo.name & ">"
					WScript.Echo "PROPIEDAD: " & oCampo.name & " VALOR: '" & oCampo.Value & "'"
			  End If
      End If
    Next
    WriteToInventoryFile "</Elemento>"
    x = x + 1
Next

Set oClass = Nothing

End Sub

'*********************************************************************************************
'Esta es la parte del script que escribe los valores de las características del PC en el 
'fichero de texto que se le pasa como parámetro.
'*********************************************************************************************
Sub WriteToInventoryFile(strValue)

Dim oFSO, oInventoryFile
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")

On Error Resume Next

Set oInventoryFile = oFSO.OpenTextFile(strInventoryFile, ForAppending, True, False)
If Err <>0 Then
        LogError "Sub WriteToInventoryFile", Err.Number, Err.Description, strValue, strInventoryFile
  Err.Clear
Else
  oInventoryFile.WriteLine strValue
  oInventoryFile.Close
  Set oInventoryFile = Nothing
End If

On Error Goto 0

End Sub

'Chequear el engine predeterminado para correr los scripts.
CheckScriptHost

'Si existe algunos de los ficheros, el del inventario o el de errores los elimina.
CheckIfExistsFiles

'Escribir los datos en el fichero de texto.

'Establecer el Encabezado del fichero
SetFileHeader 

Dim oFSO, strXMLFile, oXMLFile, strNombrePath, oNodeNombreList, x
Set oXMLFile = CreateObject("Microsoft.XMLDOM")

Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
strXMLFile = "WMI_CLASES.XML"
If oFso.FileExists("WMI_CLASES_REDUCED.XML") Then 
	strXMLFile = "WMI_CLASES_REDUCED.XML"
End If

Set oFSO = Nothing

'Chequear si se puede cargar el documento XML.
If oXMLFile.Load(strXMLFile) Then

  strNombrePath = "/WMICLASES/CLASE/Nombre"
  Set oNodeNombreList = oXMLFile.selectNodes(strNombrePath)
  For x = 0 To oNodeNombreList.length - 1
				SetClassHeader oNodeNombreList.Item(x).text
        SetClassData oNodeNombreList.Item(x).text
        SetClassFooter oNodeNombreList.Item(x).text
  Next 

Else

  WScript.Echo "Error al cargar el fichero XML que contiene las clases. Se esperaba un fichero de nombre '" & strXMLFile & _
			   "' donde se definían las clases que se encuestarías para recuperar los datos necesarios."
  WScript.Quit

End If

SetFileFooter 

WScript.Quit

