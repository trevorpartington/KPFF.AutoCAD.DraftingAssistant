using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Gile.AutoCAD.R25.Geometry;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Utilities;

/// <summary>
/// Calculates model space boundaries for viewports using AutoCAD's coordinate transformation matrix
/// </summary>
public static class ViewportBoundaryCalculator
{
    /// <summary>
    /// Gets the model space footprint of a viewport as an ordered collection of points
    /// </summary>
    /// <param name="viewport">The viewport to calculate boundaries for</param>
    /// <returns>Ordered vertices of the model-space footprint</returns>
    public static Point3dCollection GetViewportFootprint(Viewport viewport)
    {
        var result = new Point3dCollection();
        if (viewport == null)
            return result;

        try
        {
            // Use consistent scale calculation (both methods are equivalent)
            double scale = viewport.ViewHeight / viewport.Height; // == 1.0 / viewport.CustomScale

            // Model-space center of view (from ViewCenter with Y-axis inversion for DCS→model conversion)
            Point3d msCenter = new Point3d(
                viewport.ViewCenter.X,
                -viewport.ViewCenter.Y, // Invert Y for DCS to model space conversion
                0);

            // Transformation for twist rotation (negate because AutoCAD twist is clockwise)
            Matrix3d rot = Math.Abs(viewport.TwistAngle) < 1e-9
                ? Matrix3d.Identity
                : Matrix3d.Rotation(-viewport.TwistAngle, Vector3d.ZAxis, msCenter);

            if (!viewport.NonRectClipOn)
            {
                // === Rectangular viewport ===
                double halfWidth = (viewport.Width * scale) / 2.0;
                double halfHeight = (viewport.Height * scale) / 2.0;

                var modelCorners = new[]
                {
                    new Point3d(msCenter.X - halfWidth, msCenter.Y - halfHeight, 0), // BL
                    new Point3d(msCenter.X - halfWidth, msCenter.Y + halfHeight, 0), // TL
                    new Point3d(msCenter.X + halfWidth, msCenter.Y + halfHeight, 0), // TR
                    new Point3d(msCenter.X + halfWidth, msCenter.Y - halfHeight, 0)  // BR
                };

                foreach (var pt in modelCorners)
                    result.Add(pt.TransformBy(rot));
            }
            else
            {
                // === Polygonal viewport ===
                var pts = new List<Point3d>();
                Database db = viewport.Database;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ent = tr.GetObject(viewport.NonRectClipEntityId, OpenMode.ForRead) as Entity;

                    if (ent is Polyline pl)
                    {
                        for (int i = 0; i < pl.NumberOfVertices; i++)
                        {
                            var v = pl.GetPoint2dAt(i);

                            // Paper → model delta with both X and Y flips based on user's "double mirror" analysis
                            double dx = (viewport.CenterPoint.X - v.X) * scale; // <-- X inversion too
                            double dy = (v.Y - viewport.CenterPoint.Y) * scale; // <-- Y inversion

                            var p = new Point3d(msCenter.X + dx, msCenter.Y + dy, 0).TransformBy(rot);
                            pts.Add(p);
                        }

                        // Close if needed
                        if (!pl.Closed && pts.Count > 0)
                            pts.Add(pts[0]);
                    }
                    else
                    {
                        // Fallback: approximate with rectangular footprint for Circle/Ellipse clips
                        double halfWidth = (viewport.Width * scale) / 2.0;
                        double halfHeight = (viewport.Height * scale) / 2.0;

                        pts.AddRange(new[]
                        {
                            new Point3d(msCenter.X - halfWidth, msCenter.Y - halfHeight, 0),
                            new Point3d(msCenter.X - halfWidth, msCenter.Y + halfHeight, 0),
                            new Point3d(msCenter.X + halfWidth, msCenter.Y + halfHeight, 0),
                            new Point3d(msCenter.X + halfWidth, msCenter.Y - halfHeight, 0)
                        }.Select(pt => pt.TransformBy(rot)));
                    }

                    tr.Commit();
                }

                // Optional: normalize order to CCW for stable/consistent winding
                if (pts.Count >= 3)
                {
                    var cx = pts.Average(p => p.X);
                    var cy = pts.Average(p => p.Y);
                    pts = pts
                        .OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx)) // CCW from centroid
                        .ToList();
                }

                // Add to result
                foreach (var p in pts) 
                    result.Add(p);
            }
        }
        catch (System.Exception)
        {
            // Return empty collection on any transformation errors
            // This allows calling code to handle the empty result gracefully
            result.Clear();
        }

        return result;
    }

    /// <summary>
    /// Gets the bounding box of a viewport's model space footprint
    /// </summary>
    /// <param name="viewport">The viewport to calculate bounds for</param>
    /// <returns>Extents3d representing the bounding box, or null if calculation fails</returns>
    public static Extents3d? GetViewportBounds(Viewport viewport)
    {
        var footprint = GetViewportFootprint(viewport);
        
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
    /// <returns>True if the point is within the viewport boundary</returns>
    public static bool IsPointInViewport(Viewport viewport, Point3d testPoint)
    {
        var footprint = GetViewportFootprint(viewport);
        
        if (footprint.Count < 3)
            return false;

        // Use point-in-polygon algorithm (will be implemented when PointInPolygonDetector is created)
        // For now, return false as placeholder
        // TODO: Integrate with PointInPolygonDetector when it's implemented
        return false;
    }

    /// <summary>
    /// Calculates the transformation matrix from DCS to WCS for a viewport
    /// This implementation uses viewport properties to reconstruct the transformation
    /// </summary>
    /// <param name="viewport">The viewport to calculate the matrix for</param>
    /// <returns>Transformation matrix from DCS to WCS</returns>
    private static Matrix3d GetDcsToWcsMatrix(Viewport viewport)
    {
        // Get viewport properties
        Point2d viewCenter2d = viewport.ViewCenter;
        Point3d viewTarget = viewport.ViewTarget;
        Vector3d viewDirection = viewport.ViewDirection;
        double viewHeight = viewport.ViewHeight;
        double customScale = viewport.CustomScale;
        
        // Calculate the scale factor from viewport height to model space
        // ViewHeight is the height of the model space area shown in the viewport
        double modelToViewportScale = viewport.Height / viewHeight;
        
        // Calculate transformation matrix
        // 1. Scale from DCS units to model space units
        double scaleFactor = viewHeight / viewport.Height;
        Matrix3d scaleMatrix = Matrix3d.Scaling(scaleFactor, Point3d.Origin);
        
        // 2. Translate to position in model space
        // The viewport center in paper space corresponds to the view target in model space
        Vector3d translation = viewTarget.GetAsVector();
        Matrix3d translationMatrix = Matrix3d.Displacement(translation);
        
        // Combine transformations: scale first, then translate
        return scaleMatrix * translationMatrix;
    }
}