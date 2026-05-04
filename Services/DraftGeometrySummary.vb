Option Strict Off

Imports System.Collections.Generic

''' <summary>Conteos de colecciones 2D en una <see cref="SolidEdgeDraft.DrawingView"/> (interop).</summary>
Public Class GeometryCollectionCounts
    Public Property DVLines2d As Integer
    Public Property DVArcs2d As Integer
    Public Property DVCircles2d As Integer
    Public Property DVEllipses2d As Integer
    Public Property DVPoints2d As Integer
    Public Property DVLineStrings2d As Integer
    Public Property DVBSplineCurves2d As Integer
End Class

''' <summary>Línea 2D proyectada en coordenadas de hoja (metros).</summary>
Public Class LineGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVLine2d"
    Public Property X1 As Double
    Public Property Y1 As Double
    Public Property X2 As Double
    Public Property Y2 As Double
    Public Property Length As Double
    Public Property DeltaX As Double
    Public Property DeltaY As Double
    Public Property AngleRadians As Double
    Public Property AngleDegrees As Double
    Public Property Orientation As String
    Public Property MidX As Double
    Public Property MidY As Double
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
End Class

''' <summary>Arco 2D; centro y extremos en hoja cuando se ha podido transformar.</summary>
Public Class ArcGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVArc2d"
    Public Property CenterSheetX As Double
    Public Property CenterSheetY As Double
    Public Property Radius As Double
    Public Property StartSheetX As Double
    Public Property StartSheetY As Double
    Public Property EndSheetX As Double
    Public Property EndSheetY As Double
    Public Property StartAngleRaw As Double
    Public Property EndAngleRaw As Double
    Public Property SweepAngleRaw As Double
    Public Property AngleUnitNote As String
    Public Property SweepSignNote As String
    Public Property ArcLengthApprox As Double
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
    Public Property CurveClass As String
    Public Property ApiNote As String
End Class

Public Class CircleGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVCircle2d"
    Public Property CenterSheetX As Double
    Public Property CenterSheetY As Double
    Public Property Radius As Double
    Public Property Diameter As Double
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
    Public Property CurveClass As String
End Class

Public Class EllipseGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVEllipse2d"
    Public Property CenterSheetX As Double
    Public Property CenterSheetY As Double
    Public Property MajorAxis As Double
    Public Property MinorAxis As Double
    Public Property OrientationRadians As Double
    Public Property ApiNote As String
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
End Class

Public Class PointGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVPoint2d"
    Public Property SheetX As Double
    Public Property SheetY As Double
    Public Property RelationNote As String
End Class

Public Class LineStringGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVLineString2d"
    Public Property NodeCount As Integer
    Public Property LengthApprox As Double
    Public Property StartSheetX As Double
    Public Property StartSheetY As Double
    Public Property EndSheetX As Double
    Public Property EndSheetY As Double
    Public Property Closed As Boolean
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
    Public Property ApiNote As String
End Class

Public Class BSplineGeometryInfo
    Public Property Index As Integer
    Public Property EntityType As String = "DVBSplineCurve2d"
    Public Property NodeCount As Integer
    Public Property LengthApprox As Double
    Public Property StartSheetX As Double
    Public Property StartSheetY As Double
    Public Property EndSheetX As Double
    Public Property EndSheetY As Double
    Public Property Closed As Boolean
    Public Property BboxMinX As Double
    Public Property BboxMaxX As Double
    Public Property BboxMinY As Double
    Public Property BboxMaxY As Double
    Public Property ApiNote As String
End Class

''' <summary>Resumen de geometría 2D de una vista (reutilizable por el motor de acotado).</summary>
Public Class ViewGeometrySummary
    Public Property ViewIndex As Integer
    Public Property Name As String
    Public Property DrawingViewType As String
    Public Property ViewOrientation As String
    Public Property ScaleFactor As Double
    Public Property RotationNote As String
    Public Property OriginSheetX As Double
    Public Property OriginSheetY As Double
    Public Property HasOrigin As Boolean
    Public Property RangeMinX As Double
    Public Property RangeMinY As Double
    Public Property RangeMaxX As Double
    Public Property RangeMaxY As Double
    Public Property ViewWidth As Double
    Public Property ViewHeight As Double
    Public Property HasModelLink As Boolean
    Public Property ModelLinkFileName As String
    Public Property Counts As New GeometryCollectionCounts()
    Public Property Lines As New List(Of LineGeometryInfo)()
    Public Property Arcs As New List(Of ArcGeometryInfo)()
    Public Property Circles As New List(Of CircleGeometryInfo)()
    Public Property Ellipses As New List(Of EllipseGeometryInfo)()
    Public Property Points As New List(Of PointGeometryInfo)()
    Public Property LineStrings As New List(Of LineStringGeometryInfo)()
    Public Property BSplines As New List(Of BSplineGeometryInfo)()
    ''' <summary>Clasificación agregada (conteos).</summary>
    Public Property ClassifyLinesH As Integer
    Public Property ClassifyLinesV As Integer
    Public Property ClassifyLinesI As Integer
    Public Property ClassifyArcSmall As Integer
    Public Property ClassifyArcLarge As Integer
    Public Property ClassifyArcOther As Integer
    Public Property ClassifyCircHoleCandidate As Integer
    Public Property ClassifyCircOther As Integer
    Public Property ClassifyCurveOpen As Integer
    Public Property ClassifyCurveClosed As Integer
    ''' <summary>Unión de bbox de entidades (hoja).</summary>
    Public Property EntityUnionMinX As Double
    Public Property EntityUnionMaxX As Double
    Public Property EntityUnionMinY As Double
    Public Property EntityUnionMaxY As Double
    Public Property EntityUnionHasData As Boolean
    Public Property ContributorMinX As String
    Public Property ContributorMaxX As String
    Public Property ContributorMinY As String
    Public Property ContributorMaxY As String
    Public Property TolAxisM As Double

    ' Referencias reutilizables para el motor de acotado (solo en memoria; no para serialización).
    Public Property ExtremeMinXObject As Object
    Public Property ExtremeMaxXObject As Object
    Public Property ExtremeMinYObject As Object
    Public Property ExtremeMaxYObject As Object

    ' Puntos “cercanos” a la entidad para ayudar a AddDistanceBetweenObjects (coordenadas de hoja, m).
    Public Property ExtremeMinXPickX As Double
    Public Property ExtremeMinXPickY As Double
    Public Property ExtremeMaxXPickX As Double
    Public Property ExtremeMaxXPickY As Double
    Public Property ExtremeMinYPickX As Double
    Public Property ExtremeMinYPickY As Double
    Public Property ExtremeMaxYPickX As Double
    Public Property ExtremeMaxYPickY As Double
End Class
