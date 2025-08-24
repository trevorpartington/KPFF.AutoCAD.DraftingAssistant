using Autodesk.AutoCAD.Geometry;
using System;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Utilities;

/// <summary>
/// Provides point-in-polygon detection using ray casting algorithm.
/// Based on the proven LISP implementation from DBRT_UNB_DWG.lsp.
/// </summary>
public static class PointInPolygonDetector
{
    /// <summary>
    /// Determines if a test point is inside a polygon using the ray casting algorithm.
    /// This is a classic computational geometry algorithm that casts a ray from the test point
    /// to infinity and counts edge intersections. An odd number means the point is inside.
    /// </summary>
    /// <param name="testPoint">The point to test (X,Y coordinates used, Z ignored)</param>
    /// <param name="polygon">Collection of polygon vertices in order</param>
    /// <returns>True if the point is inside the polygon, false otherwise</returns>
    /// <exception cref="ArgumentNullException">Thrown when testPoint or polygon is null</exception>
    /// <exception cref="ArgumentException">Thrown when polygon has fewer than 3 vertices</exception>
    public static bool IsPointInPolygon(Point3d testPoint, Point3dCollection polygon)
    {
        if (polygon == null)
            throw new ArgumentNullException(nameof(polygon), "Polygon cannot be null");
            
        if (polygon.Count < 3)
            throw new ArgumentException("Polygon must have at least 3 vertices", nameof(polygon));

        return IsPointInPolygon(testPoint.X, testPoint.Y, polygon);
    }

    /// <summary>
    /// Internal implementation of point-in-polygon test using X,Y coordinates.
    /// Implements the ray casting algorithm exactly as in the LISP version.
    /// </summary>
    /// <param name="testX">X coordinate of test point</param>
    /// <param name="testY">Y coordinate of test point</param>
    /// <param name="polygon">Collection of polygon vertices</param>
    /// <returns>True if point is inside polygon</returns>
    private static bool IsPointInPolygon(double testX, double testY, Point3dCollection polygon)
    {
        int numVertices = polygon.Count;
        bool isInside = false;

        // Loop through all edges of the polygon
        for (int i = 0; i < numVertices; i++)
        {
            // Get current and next vertices (wrapping to start for last edge)
            Point3d currentPoint = polygon[i];
            Point3d nextPoint = polygon[(i + 1) % numVertices];

            double x1 = currentPoint.X;
            double y1 = currentPoint.Y;
            double x2 = nextPoint.X;
            double y2 = nextPoint.Y;

            // Check if test point Y is between edge Y bounds and ray intersects edge
            // This implements the standard ray casting intersection test
            if (((y1 < testY && testY <= y2) || (y2 < testY && testY <= y1)) &&
                (testX < (x1 + ((testY - y1) / (y2 - y1)) * (x2 - x1))))
            {
                isInside = !isInside; // Toggle the inside flag
            }
        }

        return isInside;
    }

    /// <summary>
    /// Tests if a point is inside a polygon with tolerance for floating-point comparisons.
    /// Useful when dealing with coordinate transformations that may introduce small numerical errors.
    /// </summary>
    /// <param name="testPoint">The point to test</param>
    /// <param name="polygon">Collection of polygon vertices in order</param>
    /// <param name="tolerance">Tolerance for floating-point comparisons (default: 1e-9)</param>
    /// <returns>True if the point is inside the polygon within tolerance</returns>
    public static bool IsPointInPolygonWithTolerance(Point3d testPoint, Point3dCollection polygon, double tolerance = 1e-9)
    {
        if (polygon == null)
            throw new ArgumentNullException(nameof(polygon), "Polygon cannot be null");
            
        if (polygon.Count < 3)
            throw new ArgumentException("Polygon must have at least 3 vertices", nameof(polygon));

        if (tolerance < 0)
            throw new ArgumentException("Tolerance must be non-negative", nameof(tolerance));

        // For very small tolerance, use the standard algorithm
        if (tolerance < 1e-12)
            return IsPointInPolygon(testPoint, polygon);

        // Test the exact point and points slightly offset in cardinal directions
        // If any of these tests return true, consider the point as inside
        var testPoints = new[]
        {
            testPoint,
            new Point3d(testPoint.X + tolerance, testPoint.Y, testPoint.Z),
            new Point3d(testPoint.X - tolerance, testPoint.Y, testPoint.Z),
            new Point3d(testPoint.X, testPoint.Y + tolerance, testPoint.Z),
            new Point3d(testPoint.X, testPoint.Y - tolerance, testPoint.Z)
        };

        foreach (var point in testPoints)
        {
            if (IsPointInPolygon(point, polygon))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Calculates the winding number of a point with respect to a polygon.
    /// An alternative to ray casting that can provide more information about
    /// the point's relationship to the polygon.
    /// </summary>
    /// <param name="testPoint">The point to test</param>
    /// <param name="polygon">Collection of polygon vertices in order</param>
    /// <returns>Winding number (0 = outside, non-zero = inside)</returns>
    public static int GetWindingNumber(Point3d testPoint, Point3dCollection polygon)
    {
        if (polygon == null)
            throw new ArgumentNullException(nameof(polygon), "Polygon cannot be null");
            
        if (polygon.Count < 3)
            throw new ArgumentException("Polygon must have at least 3 vertices", nameof(polygon));

        int windingNumber = 0;
        double testX = testPoint.X;
        double testY = testPoint.Y;

        for (int i = 0; i < polygon.Count; i++)
        {
            Point3d currentPoint = polygon[i];
            Point3d nextPoint = polygon[(i + 1) % polygon.Count];

            double x1 = currentPoint.X;
            double y1 = currentPoint.Y;
            double x2 = nextPoint.X;
            double y2 = nextPoint.Y;

            if (y1 <= testY)
            {
                if (y2 > testY) // Upward crossing
                {
                    if (IsLeft(x1, y1, x2, y2, testX, testY) > 0) // Point is left of edge
                        windingNumber++;
                }
            }
            else
            {
                if (y2 <= testY) // Downward crossing
                {
                    if (IsLeft(x1, y1, x2, y2, testX, testY) < 0) // Point is right of edge
                        windingNumber--;
                }
            }
        }

        return windingNumber;
    }

    /// <summary>
    /// Tests if a point is left of a line segment.
    /// Helper method for winding number calculation.
    /// </summary>
    /// <param name="x1">X coordinate of line start</param>
    /// <param name="y1">Y coordinate of line start</param>
    /// <param name="x2">X coordinate of line end</param>
    /// <param name="y2">Y coordinate of line end</param>
    /// <param name="testX">X coordinate of test point</param>
    /// <param name="testY">Y coordinate of test point</param>
    /// <returns>Positive if left, negative if right, zero if on line</returns>
    private static double IsLeft(double x1, double y1, double x2, double y2, double testX, double testY)
    {
        return ((x2 - x1) * (testY - y1) - (testX - x1) * (y2 - y1));
    }
}