Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Diagnóstico duro: sistema hoja ↔ vista, escala, round-trip y estado de la cota creada.
''' Prefijos: [DIM][COORD], [DIM][COORD][VIEW], [DIM][COORD][SHEET], [DIM][COORD][ROUNDTRIP], [DIM][COORD][DIM_CREATED].
''' </summary>
Friend NotInheritable Class DimensionCoordinateDiagnostics

    Friend Shared Sub LogViewAndSheetContext(
        dv As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        tag As String)

        If dv Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim sf As Double = 1.0R
            Try
                sf = CDbl(dv.ScaleFactor)
            Catch
            End Try
            Dim rminX As Double = 0, rminY As Double = 0, rmaxX As Double = 0, rmaxY As Double = 0
            Try
                dv.Range(rminX, rminY, rmaxX, rmaxY)
            Catch
            End Try
            Dim ox As Double = 0, oy As Double = 0
            Try
                ox = CDbl(dv.OriginSheetX)
                oy = CDbl(dv.OriginSheetY)
            Catch
            End Try
            Dim cx As Double = 0, cy As Double = 0
            Try
                cx = CDbl(dv.CenterX)
                cy = CDbl(dv.CenterY)
            Catch
            End Try
            log.LogLine("[DIM][COORD] " & tag & " ScaleFactor=" & sf.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) &
                        " OriginSheet=(" & ox.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & oy.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " Center=(" & cx.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & cy.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " dv.Range=(" & rminX.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & rminY.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")-(" &
                        rmaxX.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & rmaxY.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")")
            If frame IsNot Nothing Then
                log.LogLine("[DIM][COORD] " & tag & " frame.Range=(" & frame.MinX.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & frame.MinY.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")-(" &
                            frame.MaxX.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & frame.MaxY.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")")
            End If
        Catch
        End Try
    End Sub

    Friend Shared Sub LogPicksSheetAndView(
        dv As DrawingView,
        x1 As Double, y1 As Double, x2 As Double, y2 As Double,
        log As DimensionLogger,
        tag As String)

        If dv Is Nothing OrElse log Is Nothing Then Return
        log.LogLine("[DIM][COORD][SHEET] " & tag & " P1=(" & x1.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & y1.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                    " P2=(" & x2.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & y2.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")")
        Try
            Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
            dv.SheetToView(x1, y1, vx1, vy1)
            dv.SheetToView(x2, y2, vx2, vy2)
            log.LogLine("[DIM][COORD][VIEW] " & tag & " SheetToView P1=(" & vx1.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & vy1.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " P2=(" & vx2.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & vy2.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")")
        Catch ex As Exception
            log.LogLine("[DIM][COORD][VIEW] " & tag & " SheetToView failed: " & ex.Message)
        End Try
    End Sub

    Friend Shared Sub LogRoundTrip(
        dv As DrawingView,
        xSheet As Double, ySheet As Double,
        log As DimensionLogger,
        tag As String)

        If dv Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim vx As Double = 0, vy As Double = 0
            dv.SheetToView(xSheet, ySheet, vx, vy)
            Dim bx As Double = 0, by As Double = 0
            dv.ViewToSheet(vx, vy, bx, by)
            Dim dx As Double = bx - xSheet
            Dim dy As Double = by - ySheet
            log.LogLine("[DIM][COORD][ROUNDTRIP] " & tag & " sheet=(" & xSheet.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & ySheet.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " -> view=(" & vx.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & vy.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " -> sheet=(" & bx.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & "," & by.ToString("G17", System.Globalization.CultureInfo.InvariantCulture) & ")" &
                        " err=(" & dx.ToString("G2", System.Globalization.CultureInfo.InvariantCulture) & "," & dy.ToString("G2", System.Globalization.CultureInfo.InvariantCulture) & ")")
        Catch ex As Exception
            log.LogLine("[DIM][COORD][ROUNDTRIP] " & tag & " failed: " & ex.Message)
        End Try
    End Sub

    Friend Shared Sub LogCreatedDimensionState(
        d As FrameworkDimension,
        log As DimensionLogger,
        tag As String,
        desiredP1X As Double, desiredP1Y As Double, desiredP2X As Double, desiredP2Y As Double)

        If d Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim rminX As Double = 0, rminY As Double = 0, rmaxX As Double = 0, rmaxY As Double = 0
            Try
                d.Range(rminX, rminY, rmaxX, rmaxY)
            Catch
            End Try
            Dim minX As Double = System.Math.Min(rminX, rmaxX)
            Dim maxX As Double = System.Math.Max(rminX, rmaxX)
            Dim minY As Double = System.Math.Min(rminY, rmaxY)
            Dim maxY As Double = System.Math.Max(rminY, rmaxY)
            log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " Range sheet=(" & minX.ToString("G17", CultureInfo.InvariantCulture) & "," & minY.ToString("G17", CultureInfo.InvariantCulture) & ")-(" &
                        maxX.ToString("G17", CultureInfo.InvariantCulture) & "," & maxY.ToString("G17", CultureInfo.InvariantCulture) & ")")
            Try
                Dim tdx As Double = 0, tdy As Double = 0
                d.GetTextOffsets(tdx, tdy)
                log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " GetTextOffsets=(" & tdx.ToString("G17", CultureInfo.InvariantCulture) & "," & tdy.ToString("G17", CultureInfo.InvariantCulture) & ")")
            Catch
            End Try
            Try
                Dim td As Double = CDbl(d.TrackDistance)
                log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " TrackDistance=" & td.ToString("G17", CultureInfo.InvariantCulture))
            Catch
            End Try
            Try
                Dim mx As Object = Nothing
                mx = CallByName(d, "MeasurementAxisEx", CallType.Get)
                If mx IsNot Nothing Then
                    log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " MeasurementAxisEx legible=" & mx.ToString())
                End If
            Catch
            End Try
            Dim nKp As Integer = 0
            Try
                nKp = CInt(d.KeyPointCount)
            Catch
                nKp = 0
            End Try
            Dim sb As New System.Text.StringBuilder()
            sb.Append("[DIM][COORD][DIM_CREATED] ").Append(tag).Append(" KeyPointCount=").Append(nKp.ToString(CultureInfo.InvariantCulture))
            Dim lim As Integer = System.Math.Min(nKp, 12)
            For i As Integer = 0 To lim - 1
                Try
                    Dim px As Double = 0, py As Double = 0, pz As Double = 0
                    Dim kpt As SolidEdgeConstants.KeyPointType
                    Dim hdl As SolidEdgeConstants.HandleType
                    d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                    sb.Append(" K").Append(i.ToString(CultureInfo.InvariantCulture)).Append("=(").Append(px.ToString("G17", CultureInfo.InvariantCulture)).Append(",").Append(py.ToString("G17", CultureInfo.InvariantCulture)).Append(") kpt=").Append(CInt(kpt).ToString(CultureInfo.InvariantCulture))
                Catch
                End Try
            Next
            log.LogLine(sb.ToString())
            Dim midDesX As Double = (desiredP1X + desiredP2X) * 0.5R
            Dim midDesY As Double = (desiredP1Y + desiredP2Y) * 0.5R
            Dim midRngX As Double = (minX + maxX) * 0.5R
            Dim midRngY As Double = (minY + maxY) * 0.5R
            Dim dMid As Double = System.Math.Sqrt((midDesX - midRngX) * (midDesX - midRngX) + (midDesY - midRngY) * (midDesY - midRngY))
            log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " delta_mid(desired_picks vs Range_bbox)=" & dMid.ToString("G4", CultureInfo.InvariantCulture) & " m")
        Catch ex As Exception
            log.LogLine("[DIM][COORD][DIM_CREATED] " & tag & " failed: " & ex.Message)
        End Try
    End Sub

End Class
