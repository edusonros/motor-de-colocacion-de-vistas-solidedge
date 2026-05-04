Option Strict Off

Imports SolidEdgeDraft

Public Enum DimensionCandidateType
    LinearDistance
    TotalHorizontal
    TotalVertical
    PartialHorizontal
    PartialVertical
    Diameter
    Radius
    Angle
    Auxiliary
End Enum

Public Enum DimensionOrientation
    Horizontal
    Vertical
    Oblique
    Radial
    Angular
End Enum

Public Enum DimensionSide
    Top
    Bottom
    Left
    Right
    Inside
    Unknown
End Enum

Public Class Point2D
    Public Property X As Double
    Public Property Y As Double
End Class

Public Class BoundingBox2D
    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double

    Public ReadOnly Property Width As Double
        Get
            Return MaxX - MinX
        End Get
    End Property

    Public ReadOnly Property Height As Double
        Get
            Return MaxY - MinY
        End Get
    End Property
End Class

Public Class DimensionCandidate
    Public Property View As DrawingView

    Public Property Type As DimensionCandidateType
    Public Property Orientation As DimensionOrientation

    Public Property SourceObject1 As Object
    Public Property SourceObject2 As Object

    Public Property P1 As Point2D
    Public Property P2 As Point2D

    Public Property NominalValue As Double

    Public Property Priority As Integer

    Public Property IsTotalDimension As Boolean
    Public Property IsFunctionalDimension As Boolean
    Public Property IsRepeatedFeature As Boolean
    Public Property IsAuxiliary As Boolean
    Public Property IsRedundant As Boolean

    Public Property UsesHiddenGeometry As Boolean
    Public Property ViewShowsFeatureClearly As Boolean = True

    Public Property PlacementSide As DimensionSide
    Public Property PlacementPoint As Point2D

    Public Property RequiredSymbol As String

    ''' <summary>Reservado para ISO 129: referencias <see cref="DvLineGeomInfo"/> en extremos/parciales.</summary>
    Public Property IsoAuxLine1 As Object
    Public Property IsoAuxLine2 As Object
End Class

Public Class DimensionCreationResult
    Public Property Ok As Boolean
    Public Property Dimension As Object
    Public Property MethodUsed As String
    Public Property WasFallback As Boolean
    Public Property IsConnected As Boolean
    Public Property FloatingButVisible As Boolean
    Public Property ErrorMessage As String
End Class
