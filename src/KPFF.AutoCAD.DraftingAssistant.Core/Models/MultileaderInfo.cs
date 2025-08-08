using Autodesk.AutoCAD.Geometry;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Models;

public class MultileaderInfo
{
    public string StyleName { get; set; } = string.Empty;
    public Point3d Position { get; set; }
    public string TextContent { get; set; } = string.Empty;
    public int NoteNumber { get; set; }
    public bool IsValidNoteNumber => NoteNumber > 0;
}

public class ViewportInfo
{
    public Point3d Center { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public Point3d[] Boundary { get; set; } = Array.Empty<Point3d>();
    public double Scale { get; set; }
}