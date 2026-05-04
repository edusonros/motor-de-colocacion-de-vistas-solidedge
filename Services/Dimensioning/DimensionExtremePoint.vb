Option Strict Off

Imports System.Globalization

''' <summary>
''' Punto extremo geométrico real en hoja/local, con trazabilidad a la entidad 2D de vista base que lo aporta.
''' Usado para anclar cotas exteriores con picks explícitos en Dimensions.AddDistanceBetweenObjects.
''' </summary>
Friend NotInheritable Class DimensionExtremePoint

    ''' <summary>Coordenadas de hoja (m).</summary>
    Public XSheet As Double
    Public YSheet As Double

    ''' <summary>Coordenadas en marco local de la vista base (origen = MinX/MinY del Range).</summary>
    Public XLocal As Double
    Public YLocal As Double

    ''' <summary>Referencia COM (DVLine2d, DVArc2d, etc.) para la API de cotas.</summary>
    Public SourceObject As Object

    Public SourceEntityType As String
    Public SourceEntityIndex As Integer
    Public Description As String

    Public IsFromLineEndpoint As Boolean
    Public IsFromArcEndpoint As Boolean
    Public IsFromArcSample As Boolean
    Public IsFromCircleExtreme As Boolean
    Public IsFromEllipseSample As Boolean

    Public Function FormatOneLine() As String
        Return String.Format(CultureInfo.InvariantCulture,
            "type={0} idx={1} hoja=({2:0.######},{3:0.######}) local=({4:0.######},{5:0.######}) {6}",
            If(SourceEntityType, "?"), SourceEntityIndex,
            XSheet, YSheet, XLocal, YLocal,
            If(Description, ""))
    End Function

End Class
