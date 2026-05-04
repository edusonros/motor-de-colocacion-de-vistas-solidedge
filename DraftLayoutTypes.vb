Option Strict Off
Imports System.IO

''' <summary>Estructuras de datos para el sistema de layout automático de Drafts.
''' Unidades: metros (Solid Edge Draft). Origen: esquina inferior izquierda, Y hacia arriba.</summary>
Public Module DraftLayoutTypes

#Region "Estructuras geométricas"

    ''' <summary>Rectángulo 2D: coordenadas min/max en el sistema del sheet.</summary>
    Public Structure ViewRect
        Public MinX As Double
        Public MinY As Double
        Public MaxX As Double
        Public MaxY As Double

        Public ReadOnly Property Width As Double
            Get
                Return Math.Max(0, MaxX - MinX)
            End Get
        End Property

        Public ReadOnly Property Height As Double
            Get
                Return Math.Max(0, MaxY - MinY)
            End Get
        End Property

        Public ReadOnly Property CenterX As Double
            Get
                Return (MinX + MaxX) / 2.0
            End Get
        End Property

        Public ReadOnly Property CenterY As Double
            Get
                Return (MinY + MaxY) / 2.0
            End Get
        End Property

        Public ReadOnly Property Area As Double
            Get
                Return Width * Height
            End Get
        End Property

        Public Shared Function Create(minX As Double, minY As Double, maxX As Double, maxY As Double) As ViewRect
            Dim r As New ViewRect
            r.MinX = minX : r.MinY = minY
            r.MaxX = maxX : r.MaxY = maxY
            Return r
        End Function
    End Structure

    ''' <summary>Dimensiones de una vista (ancho x alto).</summary>
    Public Structure ViewSize
        Public Width As Double
        Public Height As Double

        Public ReadOnly Property Area As Double
            Get
                Return Width * Height
            End Get
        End Property

        Public Shared Function Create(w As Double, h As Double) As ViewSize
            Dim vs As New ViewSize
            vs.Width = Math.Max(0, w)
            vs.Height = Math.Max(0, h)
            Return vs
        End Function
    End Structure

    ''' <summary>Punto 2D (X, Y) en metros.</summary>
    Public Structure Point2D
        Public X As Double
        Public Y As Double
        Public Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub
    End Structure

#End Region

#Region "Medición a escala 1"

    ''' <summary>Tamaños de vistas a escala 1, indexados por orientación Solid Edge.</summary>
    Public Class ViewSizesAt1
        Private _sizes As New Dictionary(Of Integer, (Double, Double))

        Public Sub SetSize(ori As Integer, w As Double, h As Double)
            _sizes(ori) = (w, h)
        End Sub

        Public Function GetWidth(ori As Integer) As Double
            If _sizes.ContainsKey(ori) Then Return _sizes(ori).Item1
            Return 0
        End Function

        Public Function GetHeight(ori As Integer) As Double
            If _sizes.ContainsKey(ori) Then Return _sizes(ori).Item2
            Return 0
        End Function
    End Class

#End Region

#Region "Candidatos y opciones de layout"

    ''' <summary>Candidato de vista base: orientación + rotación 0° o 90°.
    ''' Incluye dimensiones efectivas tras rotación (base, derecha, inferior) a escala 1.</summary>
    Public Class BaseViewCandidate
        ''' <summary>Nombre: Front, Top, Left</summary>
        Public Property BaseOriName As String
        ''' <summary>Constante Solid Edge: igFrontView, igTopView, igLeftView</summary>
        Public Property BaseOri As Integer
        ''' <summary>True = base girada 90° (H efectivo = W original, W efectivo = H original)</summary>
        Public Property Rotated90 As Boolean
        ''' <summary>Rotación en grados: 0 o 90</summary>
        Public Property RotationDeg As Integer
        ''' <summary>Dimensiones base a escala 1 (tras rotación)</summary>
        Public Property BaseSize As ViewSize
        ''' <summary>Dimensiones vista derecha (AddByFold Right) a escala 1</summary>
        Public Property RightSize As ViewSize
        ''' <summary>Dimensiones vista inferior (AddByFold Down) a escala 1</summary>
        Public Property DownSize As ViewSize

        Public Overrides Function ToString() As String
            Return $"{BaseOriName} {(If(Rotated90, "90°", "0°"))}"
        End Function
    End Class

    ''' <summary>Opción de layout: un candidato + escala calculada + dimensiones del bloque total.
    ''' Usado para comparar y elegir el mejor.</summary>
    Public Class FoldLayoutOption
        Public Property Candidate As BaseViewCandidate
        Public Property Scale As Double
        ''' <summary>Ancho total bloque (base + gap + right) en metros, escalado</summary>
        Public Property BlockWidth As Double
        ''' <summary>Alto total bloque (base + gap + down) en metros, escalado</summary>
        Public Property BlockHeight As Double
        Public Property RejectReason As String
        Public Property LayoutScore As Double

        Public ReadOnly Property IsValid As Boolean
            Get
                Return String.IsNullOrEmpty(RejectReason)
            End Get
        End Property
    End Class

    ''' <summary>Layout resuelto: posiciones exactas de base, right, down, iso y flat.
    ''' Todo en metros, esquina superior-izquierda de cada vista.</summary>
    Public Class ResolvedLayout
        Public Property TemplatePath As String
        Public Property Scale As Double
        Public Property BaseOri As Integer
        Public Property BaseOriName As String
        Public Property Rotated90 As Boolean
        Public Property OriRight As Integer
        Public Property OriDown As Integer

        ' Posiciones top-left (X = left edge, Y = top edge; Y crece hacia arriba)
        Public Property BaseTopLeftX As Double
        Public Property BaseTopLeftY As Double
        Public Property RightTopLeftX As Double
        Public Property RightTopLeftY As Double
        Public Property DownTopLeftX As Double
        Public Property DownTopLeftY As Double
        Public Property IsoTopLeftX As Double
        Public Property IsoTopLeftY As Double
        Public Property FlatTopLeftX As Double
        Public Property FlatTopLeftY As Double

        Public Property IncludeIso As Boolean = True
        Public Property IncludeFlat As Boolean
        Public Property BaseWidth As Double
        Public Property BaseHeight As Double
        Public Property BlockWidth As Double
        Public Property BlockHeight As Double

        ' Para logging: dimensiones a escala 1 y escaladas
        Public Property BaseWidthAt1 As Double
        Public Property BaseHeightAt1 As Double
        Public Property RightWidthAt1 As Double
        Public Property RightHeightAt1 As Double
        Public Property DownWidthAt1 As Double
        Public Property DownHeightAt1 As Double
        Public Property IsoWidth As Double
        Public Property IsoHeight As Double
        ''' <summary>Borde izquierdo y ancho del zona flat (para recalcular posición tras rotación).</summary>
        Public Property FlatZoneLeft As Double
        Public Property FlatZoneWidth As Double
    End Class

#End Region

End Module
