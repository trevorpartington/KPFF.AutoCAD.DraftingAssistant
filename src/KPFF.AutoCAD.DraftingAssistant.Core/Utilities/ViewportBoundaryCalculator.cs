using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Gile.AutoCAD.R25.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Utilities;

/// <summary>
/// Calculates model space boundaries for viewports using Gile's ViewportExtension transformations
/// </summary>
public static class ViewportBoundaryCalculator
{
    /// <summary>
    /// Gets the model space footprint of a viewport as an ordered collection of points.
    /// Uses the proper PSDCS → DCS → WCS transformation chain via Gile's ViewportExtension.
    /// Handles both rectangular and polygonal viewports.
    /// </summary>
    /// <param name="viewport">The viewport to calculate boundaries for</param>
    /// <param name="transaction">Optional transaction for accessing clip entities. If null, creates internal transaction.</param>
    /// <returns>Ordered vertices of the model-space footprint in WCS coordinates</returns>
    /// <exception cref="ArgumentNullException">Thrown when viewport is null</exception>
    /// <exception cref="NotSupportedException">Thrown when viewport uses unsupported clip entity type</exception>
    /// <exception cref="InvalidOperationException">Thrown when transformation fails</exception>
    public static Point3dCollection GetViewportFootprint(Viewport viewport, Transaction? transaction = null)
    {
        if (viewport == null)
            throw new ArgumentNullException(nameof(viewport), "Viewport cannot be null");

        var result = new Point3dCollection();
        
        try
        {
            // Get the transformation chain: PSDCS → DCS → WCS using Gile's ViewportExtension
            Matrix3d psdcs2dcs = viewport.PSDCS2DCS();
            Matrix3d dcs2wcs = viewport.DCS2WCS();
            
            // Order matters: matrices apply right-to-left.
            // Start in PSDCS, transform → DCS, then → WCS
            Matrix3d fullTransform = dcs2wcs * psdcs2dcs;

            if (viewport.NonRectClipOn && viewport.NonRectClipEntityId.IsValid)
            {
                // === Polygonal viewport with clip entity ===
                var points = GetPolygonalViewportPoints(viewport, transaction, fullTransform);
                foreach (var point in points)
                    result.Add(point);
            }
            else
            {
                // === Rectangular viewport ===
                var points = GetRectangularViewportPoints(viewport, fullTransform);
                foreach (var point in points)
                    result.Add(point);
            }
        }
        catch (System.Exception ex) when (!(ex is ArgumentNullException || ex is NotSupportedException))
        {
            throw new InvalidOperationException($"Failed to calculate viewport footprint: {ex.Message}", ex);
        }

        return result;
    }

    /// <summary>
    /// Gets the model space footprint points for a rectangular viewport.
    /// </summary>
    /// <param name="viewport">The viewport to process</param>
    /// <param name="fullTransform">Combined PSDCS → DCS → WCS transformation matrix</param>
    /// <returns>Ordered list of corner points in WCS coordinates</returns>
    private static List<Point3d> GetRectangularViewportPoints(Viewport viewport, Matrix3d fullTransform)
    {
        var points = new List<Point3d>();

        // Define viewport corners in PSDCS (paper space display coordinates)
        double halfWidth = viewport.Width / 2.0;
        double halfHeight = viewport.Height / 2.0;

        // Counter-clockwise from bottom-left
        var psdcsCorners = new[]
        {
            new Point3d(viewport.CenterPoint.X - halfWidth, viewport.CenterPoint.Y - halfHeight, 0), // Bottom-left
            new Point3d(viewport.CenterPoint.X - halfWidth, viewport.CenterPoint.Y + halfHeight, 0), // Top-left
            new Point3d(viewport.CenterPoint.X + halfWidth, viewport.CenterPoint.Y + halfHeight, 0), // Top-right
            new Point3d(viewport.CenterPoint.X + halfWidth, viewport.CenterPoint.Y - halfHeight, 0)  // Bottom-right
        };

        // Transform each corner: PSDCS → DCS → WCS
        foreach (var psdcsPoint in psdcsCorners)
        {
            points.Add(psdcsPoint.TransformBy(fullTransform));
        }

        return points;
    }

    /// <summary>
    /// Gets the model space footprint points for a polygonal viewport with clip entity.
    /// </summary>
    /// <param name="viewport">The viewport to process</param>
    /// <param name="externalTransaction">Optional external transaction</param>
    /// <param name="fullTransform">Combined PSDCS → DCS → WCS transformation matrix</param>
    /// <returns>Ordered list of boundary points in WCS coordinates</returns>
    /// <exception cref="NotSupportedException">Thrown for unsupported clip entity types</exception>
    private static List<Point3d> GetPolygonalViewportPoints(Viewport viewport, Transaction? externalTransaction, Matrix3d fullTransform)
    {
        var points = new List<Point3d>();
        Database db = viewport.Database;
        bool ownTransaction = externalTransaction == null;

        Transaction tr = externalTransaction ?? db.TransactionManager.StartTransaction();
        try
        {
            var entity = tr.GetObject(viewport.NonRectClipEntityId, OpenMode.ForRead) as Entity;

            switch (entity)
            {
                case Polyline polyline:
                    for (int i = 0; i < polyline.NumberOfVertices; i++)
                    {
                        Point3d psdcsPoint = polyline.GetPoint3dAt(i);
                        points.Add(psdcsPoint.TransformBy(fullTransform));
                    }
                    break;

                case Polyline2d polyline2d:
                    foreach (ObjectId vertexId in polyline2d)
                    {
                        var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                        if (vertex != null)
                        {
                            Point3d psdcsPoint = vertex.Position;
                            points.Add(psdcsPoint.TransformBy(fullTransform));
                        }
                    }
                    break;

                case Polyline3d polyline3d:
                    foreach (ObjectId vertexId in polyline3d)
                    {
                        var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as PolylineVertex3d;
                        if (vertex != null)
                        {
                            Point3d psdcsPoint = vertex.Position;
                            points.Add(psdcsPoint.TransformBy(fullTransform));
                        }
                    }
                    break;

                default:
                    string entityType = entity?.GetType().Name ?? "null";
                    throw new NotSupportedException($"Unsupported clip entity type: {entityType}. Only Polyline, Polyline2d, and Polyline3d are supported.");
            }

            if (ownTransaction)
                tr.Commit();
        }
        finally
        {
            if (ownTransaction)
                tr.Dispose();
        }

        return points;
    }

    /// <summary>
    /// Gets the bounding box of a viewport's model space footprint
    /// </summary>
    /// <param name="viewport">The viewport to calculate bounds for</param>
    /// <param name="transaction">Optional transaction for accessing clip entities</param>
    /// <returns>Extents3d representing the bounding box, or null if calculation fails</returns>
    public static Extents3d? GetViewportBounds(Viewport viewport, Transaction? transaction = null)
    {
        var footprint = GetViewportFootprint(viewport, transaction);
        
        if (footprint.Count == 0)
            return null;

        var bounds = new Extents3d();
        
        // Initialize with first point
        bounds.Set(footprint[0], footprint[0]);
        
        // Expand to include all points
        for (int i = 1; i < footprint.Count; i++)
        {
            bounds.AddPoint(footprint[i]);
        }

        return bounds;
    }

    /// <summary>
    /// Checks if a point is within the model space footprint of a viewport
    /// </summary>
    /// <param name="viewport">The viewport to test against</param>
    /// <param name="testPoint">The point to test (in model space coordinates)</param>
    /// <param name="transaction">Optional transaction for accessing clip entities</param>
    /// <returns>True if the point is within the viewport boundary</returns>
    public static bool IsPointInViewport(Viewport viewport, Point3d testPoint, Transaction? transaction = null)
    {
        var footprint = GetViewportFootprint(viewport, transaction);
        
        if (footprint.Count < 3)
            return false;

        // Use point-in-polygon algorithm (will be implemented when PointInPolygonDetector is created)
        // For now, return false as placeholder
        // TODO: Integrate with PointInPolygonDetector when it's implemented
        return false;
    }

    /// <summary>
    /// Creates diagnostic information about viewport transformations for debugging.
    /// </summary>
    /// <param name="viewport">The viewport to analyze</param>
    /// <returns>Diagnostic information string</returns>
    public static string GetTransformationDiagnostics(Viewport viewport)
    {
        if (viewport == null)
            return "Viewport is null";

        try
        {
            var psdcs2dcs = viewport.PSDCS2DCS();
            var dcs2wcs = viewport.DCS2WCS();
            
            // Correct order: apply PSDCS→DCS first, then DCS→WCS
            var fullTransform = dcs2wcs * psdcs2dcs;

            // Test sample corner transformation to show intermediate results
            var sampleCornerPSDCS = new Point3d(
                viewport.CenterPoint.X + viewport.Width / 2.0, 
                viewport.CenterPoint.Y + viewport.Height / 2.0, 
                0);
            var sampleCornerDCS = sampleCornerPSDCS.TransformBy(psdcs2dcs);
            var sampleCornerWCS = sampleCornerDCS.TransformBy(dcs2wcs);

            return $"Viewport Transformation Diagnostics:\n" +
                   $"  Center Point (PSDCS): {viewport.CenterPoint}\n" +
                   $"  Dimensions: {viewport.Width:F3} x {viewport.Height:F3}\n" +
                   $"  View Center (DCS): {viewport.ViewCenter}\n" +
                   $"  View Target (WCS): {viewport.ViewTarget}\n" +
                   $"  View Direction: {viewport.ViewDirection}\n" +
                   $"  View Height: {viewport.ViewHeight:F3}\n" +
                   $"  Custom Scale: {viewport.CustomScale:F6}\n" +
                   $"  Twist Angle: {viewport.TwistAngle:F6} rad ({viewport.TwistAngle * 180.0 / Math.PI:F2}°)\n" +
                   $"  Non-Rect Clip: {viewport.NonRectClipOn}\n" +
                   $"  Clip Entity Valid: {viewport.NonRectClipEntityId.IsValid}\n" +
                   $"  \n" +
                   $"  Transformation Matrices:\n" +
                   $"  PSDCS→DCS: {FormatMatrix(psdcs2dcs)}\n" +
                   $"  DCS→WCS: {FormatMatrix(dcs2wcs)}\n" +
                   $"  Full (DCS2WCS * PSDCS2DCS): {FormatMatrix(fullTransform)}\n" +
                   $"  \n" +
                   $"  Sample Corner Transformation (top-right):\n" +
                   $"  PSDCS: {sampleCornerPSDCS}\n" +
                   $"  → DCS: {sampleCornerDCS}\n" +
                   $"  → WCS: {sampleCornerWCS}";
        }
        catch (System.Exception ex)
        {
            return $"Error getting diagnostics: {ex.Message}";
        }
    }

    /// <summary>
    /// Formats a transformation matrix for display.
    /// </summary>
    private static string FormatMatrix(Matrix3d matrix)
    {
        return $"[{matrix[0,0]:F3},{matrix[0,1]:F3},{matrix[0,2]:F3},{matrix[0,3]:F3}]" +
               $"[{matrix[1,0]:F3},{matrix[1,1]:F3},{matrix[1,2]:F3},{matrix[1,3]:F3}]" +
               $"[{matrix[2,0]:F3},{matrix[2,1]:F3},{matrix[2,2]:F3},{matrix[2,3]:F3}]" +
               $"[{matrix[3,0]:F3},{matrix[3,1]:F3},{matrix[3,2]:F3},{matrix[3,3]:F3}]";
    }
}