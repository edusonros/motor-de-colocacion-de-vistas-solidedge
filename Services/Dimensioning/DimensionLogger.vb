Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>Prefija mensajes de acotado con [DIM] para trazabilidad en el log principal.</summary>
Friend NotInheritable Class DimensionLogger
    Private ReadOnly _log As Logger

    Public Sub New(logger As Logger)
        _log = logger
    End Sub

    Public Sub Info(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM] " & msg)
    End Sub

    Public Sub Warn(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][WARN] " & msg)
    End Sub

    Public Sub Err(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][ERR] " & msg)
    End Sub

    Public Sub Vert(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][VERT] " & msg)
    End Sub

    Public Sub VertWarn(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][VERT][WARN] " & msg)
    End Sub

    Public Sub VertFallback(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][VERT][FALLBACK] " & msg)
    End Sub

    ''' <summary>Log de colocación con prefijo exacto [DIM][PLACE][SRC] (sin añadir [DIM] extra).</summary>
    Public Sub PlaceSrc(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][PLACE][SRC] " & msg)
    End Sub

    ''' <summary>Prefijo [DIM][PLACE][ASSERT].</summary>
    Public Sub PlaceAssertMsg(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][PLACE][ASSERT] " & msg)
    End Sub

    ''' <summary>Prefijo [DIM][PLACE][FIX].</summary>
    Public Sub PlaceFixMsg(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][PLACE][FIX] " & msg)
    End Sub

    ''' <summary>Prefijo <c>[DIM][FRAME]</c> (marco único vista base).</summary>
    Public Sub Frame(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][FRAME] " & msg)
    End Sub

    ''' <summary>Prefijo <c>[DIM][PLACE]</c> (colocación con feature / local / hoja).</summary>
    Public Sub DimPlaceLine(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][PLACE] " & msg)
    End Sub

    ''' <summary>Prefijo <c>[DIM][ASSERT]</c>.</summary>
    Public Sub DimAssert(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][ASSERT] " & msg)
    End Sub

    ''' <summary>Cota insertada: rango en hoja y equivalente local respecto al marco base.</summary>
    Public Sub DimPost(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][POST] " & msg)
    End Sub

    ''' <summary>Puntos extremos geométricos reales (<c>[DIM][EXTPT]</c>).</summary>
    Public Sub ExtPt(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][EXTPT] " & msg)
    End Sub

    ''' <summary>Reposicionamiento post-inserción (<c>TrackDistance</c>, <c>SetKeyPoint</c>, etc.).</summary>
    Public Sub Repos(msg As String)
        If _log Is Nothing Then Return
        _log.Log("[DIM][REPOS] " & msg)
    End Sub

    ''' <summary>Línea literal al log principal (p. ej. prefijos <c>[DIM][COORD]</c> ya completos en el mensaje).</summary>
    Public Sub LogLine(msg As String)
        If _log Is Nothing Then Return
        _log.Log(msg)
    End Sub

    ''' <summary>Error COM detallado: método, objeto, tipo, HRESULT, mensaje.</summary>
    Public Sub ComFail(methodName As String, comTargetDescription As String, ex As Exception)
        If ex Is Nothing Then Return
        Dim sb As New StringBuilder()
        sb.Append("Error insertando cota: ")
        sb.Append(methodName)
        sb.Append(" | objeto=")
        sb.Append(comTargetDescription)
        sb.Append(" | ex=")
        sb.Append(ex.GetType().FullName)
        Dim cex = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            sb.Append(" | HRESULT=0x")
            sb.Append(cex.ErrorCode.ToString("X8", Globalization.CultureInfo.InvariantCulture))
        End If
        sb.Append(" | ")
        sb.Append(ex.Message)
        Err(sb.ToString())
    End Sub
End Class
