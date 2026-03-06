url = "https://example.com/ejecutar_calculadora.bat"
savePath = "C:\Temp\ejecutar_calculadora.bat"

Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close
    MsgBox "Archivo descargado en: " & savePath
Else
    MsgBox "Error al descargar el archivo. Código: " & http.Status
End If

Set objShell = CreateObject("WScript.Shell")
objShell.Run "C:\ruta\a\tu\script\ejecutar_calculadora.bat", 1, True