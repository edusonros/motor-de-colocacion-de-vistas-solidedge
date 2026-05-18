Option Strict On

' Notas / pruebas manuales de auditoría DFT. No se compila con Extraer_dft_dxf_flatdxf.vbproj.
' Visual Studio lo abría como «Archivo varios»; se restaura para evitar error 0x80070002.

Namespace Tests

''' <summary>Ayuda para probar DftAuditService desde el depurador o la ventana inmediata.</summary>
Public Module DftAuditServiceTests

    ' Ejemplo (desde código con Logger instanciado):
    ' DftAuditService.AnalyzeDft("ruta\plano.dft", logger, keepSolidEdgeVisible:=False, outputFolder:="C:\temp\audit")

End Module

End Namespace
