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
            // Build DCS→WCS transformation manually from viewport data
            // This works even if the layout is not active by using only stored viewport properties
            Matrix3d dcs2wcs = BuildDcsToWcs(viewport);

            if (viewport.NonRectClipOn && viewport.NonRectClipEntityId.IsValid)
            {
                // === Polygonal viewport with clip entity ===
                var points = GetPolygonalViewportPoints(viewport, transaction, dcs2wcs);
                foreach (var point in points)
                    result.Add(point);
            }
            else
            {
                // === Rectangular viewport ===
                var points = GetRectangularViewportPoints(viewport, dcs2wcs);
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
    /// Builds a DCS→WCS transformation matrix from viewport data.
    /// Works even if the layout is not active by manually constructing coordinate system from viewport properties.
    /// </summary>
    /// <param name="viewport">The viewport to build transformation for</param>
    /// <returns>Transformation matrix from DCS to WCS coordinates</returns>
    private static Matrix3d BuildDcsToWcs(Viewport vp)
    {
        // View Z (view direction, normalized)
        Vector3d vz = vp.ViewDirection.GetNormal();

        // A provisional X axis in the view plane (perp to vz).
        // If vz is near World Z, use World X as helper to avoid degeneracy.
        Vector3d helper = Math.Abs(vz.DotProduct(Vector3d.ZAxis)) > 0.999999 ? Vector3d.XAxis : Vector3d.ZAxis;
        Vector3d vx0 = helper.CrossProduct(vz).GetNormal();        // lies in view plane
        Vector3d vy0 = vz.CrossProduct(vx0).GetNormal();            // completes right-handed basis

        // Apply TwistAngle: rotate basis around view normal (vz)
        Matrix3d twist = Matrix3d.Rotation(vp.TwistAngle, vz, Point3d.Origin);
        Vector3d vx = vx0.TransformBy(twist);
        Vector3d vy = vy0.TransformBy(twist);

        // Scale: DCS (paper units) → model units is 1 / CustomScale
        double s = 1.0 / vp.CustomScale;

        // Build an affine transform such that:
        // W = ViewTarget + vx*( (Dx - ViewCenter.X) * s ) + vy*( (Dy - ViewCenter.Y) * s )
        // We can assemble this with AlignCoordinateSystem from a local "DCS-like" frame.

        // Matrix that maps from a local XY frame to model (columns are basis vectors, origin at ViewTarget)
        Matrix3d basisToWcs =
            Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                vp.ViewTarget, vx, vy, vz);

        // Matrix that recenters DCS to have origin at ViewCenter and scales by s
        Matrix3d dcsCenterAndScale =
            Matrix3d.Displacement(new Vector3d(-vp.ViewCenter.X, -vp.ViewCenter.Y, 0.0)) *
            Matrix3d.Scaling(s, Point3d.Origin);

        // Full DCS → WCS
        return basisToWcs * dcsCenterAndScale;
    }


    /// <summary>
    /// Gets the model space footprint points for a rectangular viewport.
    /// </summary>
    /// <param name="vp">The viewport to process</param>
    /// <param name="dcs2wcs">DCS → WCS transformation matrix</param>
    /// <returns>Ordered list of corner points in WCS coordinates</returns>
    private static List<Point3d> GetRectangularViewportPoints(Viewport vp, Matrix3d dcs2wcs)
    {
        // In DCS, the viewport rectangle is centered at ViewCenter and its size is the PAPER size.
        // Half-sizes in DCS (paper units):
        double halfW = vp.Width / 2.0;
        double halfH = vp.Height / 2.0;

        var dcsCorners = new[]
        {
            new Point3d(vp.ViewCenter.X - halfW, vp.ViewCenter.Y - halfH, 0), // BL
            new Point3d(vp.ViewCenter.X - halfW, vp.ViewCenter.Y + halfH, 0), // TL
            new Point3d(vp.ViewCenter.X + halfW, vp.ViewCenter.Y + halfH, 0), // TR
            new Point3d(vp.ViewCenter.X + halfW, vp.ViewCenter.Y - halfH, 0), // BR
        };

        return dcsCorners.Select(p => p.TransformBy(dcs2wcs)).ToList();
    }

    /// <summary>
    /// Gets the model space footprint points for a polygonal viewport with clip entity.
    /// </summary>
    /// <param name="viewport">The viewport to process</param>
    /// <param name="externalTransaction">Optional external transaction</param>
    /// <param name="dcs2wcs">DCS → WCS transformation matrix</param>
    /// <returns>Ordered list of boundary points in WCS coordinates</returns>
    /// <exception cref="NotSupportedException">Thrown for unsupported clip entity types</exception>
    private static List<Point3d> GetPolygonalViewportPoints(Viewport viewport, Transaction? externalTransaction, Matrix3d dcs2wcs)
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
                        points.Add(clipPoint.TransformBy(dcs2wcs));
                    }
                    break;

                case Polyline2d polyline2d:
                    foreach (ObjectId vertexId in polyline2d)
                    {
                        var vertex = tr.GetObject(vertexId, OpenMode.ForRead) as Vertex2d;
                        if (vertex != null)
                        {
                            Point3d clipPoint = vertex.Position;
                            points.Add(clipPoint.TransformBy(dcs2wcs));
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
                            points.Add(clipPoint.TransformBy(dcs2wcs));
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
            var dcs2wcs = BuildDcsToWcs(viewport);

            // Test sample corner transformation from DCS to show results
            // Use ViewCenter and ViewHeight to define a sample corner in DCS
            double halfHeight = viewport.ViewHeight / 2.0;
            double halfWidth = (viewport.ViewHeight * viewport.Width / viewport.Height) / 2.0;
            
            var sampleCornerDCS = new Point3d(
                viewport.ViewCenter.X + halfWidth, 
                viewport.ViewCenter.Y + halfHeight, 
                0);
            var sampleCornerWCS = sampleCornerDCS.TransformBy(dcs2wcs);

            return $"Viewport Transformation Diagnostics (Manual DCS→WCS):\n" +
                   $"  Center Point (Paper): {viewport.CenterPoint}\n" +
                   $"  Dimensions (Paper): {viewport.Width:F3} x {viewport.Height:F3}\n" +
                   $"  View Center (DCS): {viewport.ViewCenter}\n" +
                   $"  View Target (WCS): {viewport.ViewTarget}\n" +
                   $"  View Direction: {viewport.ViewDirection}\n" +
                   $"  View Height (DCS): {viewport.ViewHeight:F3}\n" +
                   $"  Custom Scale: {viewport.CustomScale:F6}\n" +
                   $"  Twist Angle: {viewport.TwistAngle:F6} rad ({viewport.TwistAngle * 180.0 / Math.PI:F2}°)\n" +
                   $"  Non-Rect Clip: {viewport.NonRectClipOn}\n" +
                   $"  Clip Entity Valid: {viewport.NonRectClipEntityId.IsValid}\n" +
                   $"  \n" +
                   $"  Manual Transformation Matrix:\n" +
                   $"  DCS→WCS: {FormatMatrix(dcs2wcs)}\n" +
                   $"  \n" +
                   $"  Sample Corner Transformation (DCS top-right):\n" +
                   $"  DCS: {sampleCornerDCS}\n" +
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