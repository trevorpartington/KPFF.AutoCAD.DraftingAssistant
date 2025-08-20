using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Gile.AutoCAD.R25.Geometry;

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
            if (!viewport.NonRectClipOn)
            {
                // Rectangular viewport boundary
                // Apply the custom scale factor manually since GeometryExtensions DCS2WCS isn't working correctly
                double scaleFactor = 1.0 / viewport.CustomScale; // CustomScale 0.05 -> scale factor 20
                double halfWidth = (viewport.Width * scaleFactor) / 2.0;
                double halfHeight = (viewport.Height * scaleFactor) / 2.0;

                // Get the view center (center of the displayed area in model space coordinates)
                // ViewCenter is in paper space coordinates, but represents the model space center
                Point3d modelSpaceCenter = new Point3d(viewport.ViewCenter.X, viewport.ViewCenter.Y, 0);

                // Define corners in model space relative to the model space center
                var modelCorners = new[]
                {
                    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y - halfHeight, 0), // Bottom-Left
                    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y + halfHeight, 0), // Top-Left
                    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y + halfHeight, 0), // Top-Right
                    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y - halfHeight, 0)  // Bottom-Right
                };

                // Add corners directly (no transformation needed since we calculated them in model space)
                foreach (var modelPoint in modelCorners)
                {
                    result.Add(modelPoint);
                }
            }
            else
            {
                // Polygonal viewport boundary (non-rectangular clipping)
                // For now, fall back to rectangular boundary with proper scaling
                double scaleFactor = 1.0 / viewport.CustomScale;
                double halfWidth = (viewport.Width * scaleFactor) / 2.0;
                double halfHeight = (viewport.Height * scaleFactor) / 2.0;
                
                Point3d modelSpaceCenter = new Point3d(viewport.ViewCenter.X, viewport.ViewCenter.Y, 0);

                var modelCorners = new[]
                {
                    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y - halfHeight, 0),
                    new Point3d(modelSpaceCenter.X - halfWidth, modelSpaceCenter.Y + halfHeight, 0),
                    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y + halfHeight, 0),
                    new Point3d(modelSpaceCenter.X + halfWidth, modelSpaceCenter.Y - halfHeight, 0)
                };

                foreach (var modelPoint in modelCorners)
                {
                    result.Add(modelPoint);
                }
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