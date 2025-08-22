using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Utilities;

/// <summary>
/// Calculates model space boundaries for viewports using manual transformation matrices.
/// Works on any layout (active or not) by reconstructing transformations from viewport properties.
/// </summary>
public static class ViewportBoundaryCalculator
{
    /// <summary>
    /// Gets the model space footprint of a viewport as an ordered collection of points.
    /// Uses manual DCS → WCS transformation built from viewport properties.
    /// Works on any layout (active or not). Handles both rectangular and polygonal viewports.
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
            if (viewport.NonRectClipOn && viewport.NonRectClipEntityId.IsValid)
            {
                // === Polygonal viewport with clip entity ===
                var points = GetPolygonalViewportPoints(viewport, transaction);
                foreach (var point in points)
                    result.Add(point);
            }
            else
            {
                // === Rectangular viewport ===
                var points = GetRectangularViewportPoints(viewport);
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
    /// Transforms a point from paper space coordinates to model space coordinates.
    /// Uses the correct transformation sequence: scale from ViewCenter, then rotate around origin.
    /// </summary>
    /// <param name="paperPoint">Point in paper space coordinates</param>
    /// <param name="vp">The viewport to use for transformation</param>
    /// <returns>Point in model space coordinates</returns>
    private static Point3d TransformPaperToModel(Point3d paperPoint, Viewport vp)
    {
        // Step 1: Scale from ViewCenter by 1/CustomScale (paper units → model units)
        double scaleFactor = 1.0 / vp.CustomScale;
        
        // Vector from ViewCenter to the point
        Vector3d fromCenter = new Vector3d(
            paperPoint.X - vp.ViewCenter.X,
            paperPoint.Y - vp.ViewCenter.Y,
            0);
        
        // Scale this vector
        Vector3d scaledFromCenter = fromCenter * scaleFactor;
        
        // Apply scaling from ViewCenter
        Point3d scaledPoint = new Point3d(
            vp.ViewCenter.X + scaledFromCenter.X,
            vp.ViewCenter.Y + scaledFromCenter.Y,
            0);
        
        // Step 2: Rotate around origin (0,0) by TwistAngle
        Matrix3d rotation = Matrix3d.Rotation(vp.TwistAngle, Vector3d.ZAxis, Point3d.Origin);
        Point3d rotatedPoint = scaledPoint.TransformBy(rotation);
        
        return rotatedPoint;
    }


    /// <summary>
    /// Gets the model space footprint points for a rectangular viewport.
    /// </summary>
    /// <param name="vp">The viewport to process</param>
    /// <returns>Ordered list of corner points in model space coordinates</returns>
    private static List<Point3d> GetRectangularViewportPoints(Viewport vp)
    {
        // Build rectangle in paper space coordinates around ViewCenter
        // Half-sizes in paper units:
        double halfW = vp.Width / 2.0;
        double halfH = vp.Height / 2.0;

        var paperCorners = new[]
        {
            new Point3d(vp.ViewCenter.X - halfW, vp.ViewCenter.Y - halfH, 0), // BL
            new Point3d(vp.ViewCenter.X - halfW, vp.ViewCenter.Y + halfH, 0), // TL
            new Point3d(vp.ViewCenter.X + halfW, vp.ViewCenter.Y + halfH, 0), // TR
            new Point3d(vp.ViewCenter.X + halfW, vp.ViewCenter.Y - halfH, 0), // BR
        };

        // Transform each paper space point to model space using correct sequence
        return paperCorners.Select(p => TransformPaperToModel(p, vp)).ToList();
    }

    /// <summary>
    /// Gets the model space footprint points for a polygonal viewport with clip entity.
    /// </summary>
    /// <param name="viewport">The viewport to process</param>
    /// <param name="externalTransaction">Optional external transaction</param>
    /// <returns>Ordered list of boundary points in model space coordinates</returns>
    /// <exception cref="NotSupportedException">Thrown for unsupported clip entity types</exception>
    private static List<Point3d> GetPolygonalViewportPoints(Viewport viewport, Transaction? externalTransaction)
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
                        Point3d clipPoint = polyline.GetPoint3dAt(i);
                        points.Add(TransformPaperToModel(clipPoint, viewport));
                    }
                    break;

                case Polyline2d polyline2d:
                    foreach (ObjectId vertexId in polyline2d)
                    {
                        var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                        if (vertex != null)
                        {
                            Point3d clipPoint = vertex.Position;
                            points.Add(TransformPaperToModel(clipPoint, viewport));
                        }
                    }
                    break;

                case Polyline3d polyline3d:
                    foreach (ObjectId vertexId in polyline3d)
                    {
                        var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as PolylineVertex3d;
                        if (vertex != null)
                        {
                            Point3d clipPoint = vertex.Position;
                            points.Add(TransformPaperToModel(clipPoint, viewport));
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
            // Test sample corner transformation using the new method
            // Top-right corner in paper space
            double halfWidth = viewport.Width / 2.0;
            double halfHeight = viewport.Height / 2.0;
            
            var sampleCornerPaper = new Point3d(
                viewport.ViewCenter.X + halfWidth, 
                viewport.ViewCenter.Y + halfHeight, 
                0);
            var sampleCornerModel = TransformPaperToModel(sampleCornerPaper, viewport);

            return $"Viewport Transformation Diagnostics (Simplified Approach):\n" +
                   $"  Center Point (Paper): {viewport.CenterPoint}\n" +
                   $"  Dimensions (Paper): {viewport.Width:F3} x {viewport.Height:F3}\n" +
                   $"  View Center: {viewport.ViewCenter}\n" +
                   $"  View Target: {viewport.ViewTarget}\n" +
                   $"  View Direction: {viewport.ViewDirection}\n" +
                   $"  View Height: {viewport.ViewHeight:F3}\n" +
                   $"  Custom Scale: {viewport.CustomScale:F6} (Scale Factor: {1.0/viewport.CustomScale:F1})\n" +
                   $"  Twist Angle: {viewport.TwistAngle:F6} rad ({viewport.TwistAngle * 180.0 / Math.PI:F2}°)\n" +
                   $"  Non-Rect Clip: {viewport.NonRectClipOn}\n" +
                   $"  Clip Entity Valid: {viewport.NonRectClipEntityId.IsValid}\n" +
                   $"  \n" +
                   $"  Transformation Steps:\n" +
                   $"  1. Build rectangle in paper space around ViewCenter\n" +
                   $"  2. Scale from ViewCenter by {1.0/viewport.CustomScale:F1}\n" +
                   $"  3. Rotate around origin (0,0) by {viewport.TwistAngle * 180.0 / Math.PI:F2}°\n" +
                   $"  \n" +
                   $"  Sample Corner Transformation (top-right):\n" +
                   $"  Paper Space: {sampleCornerPaper}\n" +
                   $"  → Model Space: {sampleCornerModel}";
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