using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using KPFF.AutoCAD.DraftingAssistant.Core.Interfaces;

namespace KPFF.AutoCAD.DraftingAssistant.Core.Services;

/// <summary>
/// Service for inserting blocks from external DWG files into the current drawing
/// </summary>
public class BlockInsertionService
{
    private readonly ILogger _logger;

    public BlockInsertionService(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Inserts a stack of 24 construction note blocks (NT01 through NT24)
    /// </summary>
    /// <param name="blockFilePath">Path to the DWG file containing the block</param>
    /// <returns>True if insertion was successful</returns>
    public bool InsertConstructionNoteBlockStack(string blockFilePath)
    {
        try
        {
            _logger.LogInformation($"Starting construction note block stack insertion from {blockFilePath}");
            
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            if (!File.Exists(blockFilePath))
            {
                _logger.LogError($"Block file not found: {blockFilePath}");
                ed.WriteMessage($"\nError: Block file not found: {blockFilePath}");
                return false;
            }

            // Get insertion point from user for NT01 (the top block)
            var ppr = ed.GetPoint($"\nSelect insertion point for construction note stack: ");
            if (ppr.Status != PromptStatus.OK)
            {
                _logger.LogInformation("User cancelled block stack insertion");
                return false;
            }

            var baseInsertionPoint = ppr.Value;
            var successfulInsertions = 0;

            // Perform entire operation with document lock to avoid lock violations
            using (var docLock = doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Insert all 24 blocks stacked vertically, each with unique name
                        for (int i = 1; i <= 24; i++)
                        {
                            var blockName = $"NT{i:D2}"; // NT01, NT02, etc.
                            
                            // Import each block with its unique name
                            var blockDefId = ImportBlockDefinition(db, blockFilePath, blockName);
                            if (blockDefId == ObjectId.Null)
                            {
                                _logger.LogError($"Failed to import block definition: {blockName}");
                                continue;
                            }
                            
                            var yOffset = -(i - 1) * 0.5; // 0, -0.5, -1.0, etc.
                            var insertionPoint = new Point3d(
                                baseInsertionPoint.X,
                                baseInsertionPoint.Y + yOffset,
                                baseInsertionPoint.Z);

                            var success = InsertBlockReference(db, tr, blockDefId, blockName, insertionPoint);
                            if (success)
                            {
                                successfulInsertions++;
                                _logger.LogDebug($"Successfully inserted block: {blockName} at {insertionPoint}");
                            }
                            else
                            {
                                _logger.LogError($"Failed to insert block: {blockName}");
                            }
                        }

                        tr.Commit();
                        
                        _logger.LogInformation($"Construction note block stack insertion completed. {successfulInsertions}/24 blocks inserted successfully.");
                        ed.WriteMessage($"\nConstruction note stack inserted: {successfulInsertions}/24 blocks successful.");
                        
                        return successfulInsertions > 0;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during block stack insertion transaction: {ex.Message}", ex);
                        tr.Abort();
                        return false;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error inserting construction note block stack: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Inserts a title block from an external DWG file at the origin (0,0)
    /// Supports both dynamic and static blocks from the specified DWG file.
    /// </summary>
    /// <param name="blockFilePath">Path to the title block DWG file (created with WBLOCK)</param>
    /// <returns>True if insertion was successful</returns>
    public bool InsertTitleBlock(string blockFilePath)
    {
        try
        {
            _logger.LogInformation($"Starting title block insertion from {blockFilePath}");
            
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            if (!File.Exists(blockFilePath))
            {
                _logger.LogError($"Title block file not found: {blockFilePath}");
                ed.WriteMessage($"\nError: Title block file not found: {blockFilePath}");
                return false;
            }

            // Insert at origin (0,0,0)
            var insertionPoint = Point3d.Origin;
            
            // The block name is the same as the DWG filename (without extension)
            var blockName = Path.GetFileNameWithoutExtension(blockFilePath);
            _logger.LogInformation($"Using title block: {blockName}");

            // Perform entire operation with document lock to avoid lock violations
            using (var docLock = doc.LockDocument())
            {
                // Import the block definition
                var blockDefId = ImportBlockDefinition(db, blockFilePath, blockName);
                if (blockDefId == ObjectId.Null)
                {
                    _logger.LogError("Failed to import title block definition");
                    ed.WriteMessage("\nError: Failed to import title block definition");
                    return false;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Insert the block reference at origin
                        var success = InsertBlockReference(db, tr, blockDefId, blockName, insertionPoint);
                        if (success)
                        {
                            tr.Commit();
                            _logger.LogInformation($"Successfully inserted title block: {blockName} at {insertionPoint}");
                            ed.WriteMessage($"\nTitle block '{blockName}' inserted at origin (0,0).");
                            return true;
                        }
                        else
                        {
                            _logger.LogError($"Failed to create title block reference: {blockName}");
                            ed.WriteMessage("\nError: Failed to create title block reference");
                            tr.Abort();
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during title block insertion transaction: {ex.Message}", ex);
                        tr.Abort();
                        return false;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error inserting title block: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Inserts a construction note block from an external DWG file
    /// </summary>
    /// <param name="blockFilePath">Path to the DWG file containing the block</param>
    /// <param name="blockName">Name of the block to insert</param>
    /// <returns>True if insertion was successful</returns>
    public bool InsertConstructionNoteBlock(string blockFilePath, string blockName = "NT01")
    {
        try
        {
            _logger.LogInformation($"Starting block insertion: {blockName} from {blockFilePath}");
            
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            if (!File.Exists(blockFilePath))
            {
                _logger.LogError($"Block file not found: {blockFilePath}");
                ed.WriteMessage($"\nError: Block file not found: {blockFilePath}");
                return false;
            }

            // Get insertion point from user first
            var ppr = ed.GetPoint($"\nSelect insertion point for {blockName}: ");
            if (ppr.Status != PromptStatus.OK)
            {
                _logger.LogInformation("User cancelled block insertion");
                return false;
            }

            // Perform entire operation with document lock to avoid lock violations
            using (var docLock = doc.LockDocument())
            {
                // Import the block definition
                var blockDefId = ImportBlockDefinition(db, blockFilePath, blockName);
                if (blockDefId == ObjectId.Null)
                {
                    _logger.LogError($"Failed to import block definition: {blockName}");
                    return false;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Create and insert the block reference
                        var success = InsertBlockReference(db, tr, blockDefId, blockName, ppr.Value);
                        
                        if (success)
                        {
                            tr.Commit();
                            _logger.LogInformation($"Successfully inserted block: {blockName}");
                            ed.WriteMessage($"\nBlock {blockName} inserted successfully.");
                            return true;
                        }
                        else
                        {
                            _logger.LogError($"Failed to create block reference: {blockName}");
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error during block insertion transaction: {ex.Message}", ex);
                        tr.Abort();
                        return false;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error inserting construction note block: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Imports a dynamic block definition from an external DWG file using Database.Insert
    /// Operates outside of transaction scope to avoid lock violations
    /// </summary>
    private ObjectId ImportBlockDefinition(Database targetDb, string blockFilePath, string blockName)
    {
        try
        {
            _logger.LogDebug($"Importing dynamic block definition: {blockName} from {blockFilePath}");

            // Check if block already exists using a read-only transaction
            using (var readTr = targetDb.TransactionManager.StartTransaction())
            {
                var blockTable = readTr.GetObject(targetDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (blockTable != null && blockTable.Has(blockName))
                {
                    _logger.LogDebug($"Block definition {blockName} already exists, using existing definition");
                    var existingId = blockTable[blockName];
                    readTr.Commit();
                    return existingId;
                }
                readTr.Commit();
            }

            // Import dynamic block using Database.Insert outside of main transaction
            using (var sourceDb = new Database(false, true))
            {
                try
                {
                    _logger.LogDebug($"Reading DWG file: {blockFilePath}");
                    sourceDb.ReadDwgFile(blockFilePath, FileOpenMode.OpenForReadAndAllShare, true, "");
                    _logger.LogDebug($"Successfully read DWG file");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    _logger.LogError($"Error reading {blockFilePath}: {ex.ErrorStatus} - {ex.Message}");
                    return ObjectId.Null;
                }

                try
                {
                    // Use Database.Insert to preserve dynamic block properties
                    var insertedBlockId = targetDb.Insert(blockName, sourceDb, false);
                    
                    if (insertedBlockId != ObjectId.Null)
                    {
                        _logger.LogDebug($"Successfully imported dynamic block definition: {blockName}");
                        return insertedBlockId;
                    }
                    else
                    {
                        _logger.LogError($"Database.Insert returned null ObjectId for block: {blockName}");
                        return ObjectId.Null;
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    _logger.LogError($"Error inserting dynamic block definition {blockName}: {ex.ErrorStatus} - {ex.Message}");
                    return ObjectId.Null;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Unexpected error importing dynamic block definition: {ex.Message}", ex);
            return ObjectId.Null;
        }
    }

    /// <summary>
    /// Creates and inserts a block reference into the current space with proper attribute synchronization
    /// </summary>
    private bool InsertBlockReference(Database db, Transaction tr, ObjectId blockDefId, string blockName, Point3d insertionPoint)
    {
        try
        {
            _logger.LogDebug($"Creating block reference for: {blockName} at {insertionPoint}");

            // Create the block reference
            var blockRef = new BlockReference(insertionPoint, blockDefId);
            blockRef.ScaleFactors = new Scale3d(1.0, 1.0, 1.0);
            blockRef.Rotation = 0.0;

            // Add to the current space (model space or layout)
            var currentSpaceId = db.CurrentSpaceId;
            var currentSpace = tr.GetObject(currentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            
            if (currentSpace != null)
            {
                currentSpace.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);
                
                // Synchronize attributes from the block definition
                // This is crucial for dynamic blocks to have their attributes available
                SyncBlockAttributes(db, tr, blockRef, blockDefId);
                
                _logger.LogDebug($"Block reference created successfully with attributes: {blockName}");
                return true;
            }
            else
            {
                _logger.LogError("Could not access current space for block insertion");
                blockRef.Dispose();
                return false;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error creating block reference: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Creates and inserts a block reference into the specified space with proper attribute synchronization
    /// Optimized version that accepts the target space directly to avoid lock violations
    /// </summary>
    private bool InsertBlockReferenceToSpace(Database db, Transaction tr, ObjectId blockDefId, string blockName, Point3d insertionPoint, BlockTableRecord targetSpace)
    {
        try
        {
            _logger.LogDebug($"Creating block reference for: {blockName} at {insertionPoint}");

            // Create the block reference
            var blockRef = new BlockReference(insertionPoint, blockDefId);
            blockRef.ScaleFactors = new Scale3d(1.0, 1.0, 1.0);
            blockRef.Rotation = 0.0;

            // Add to the specified space (already opened for write in the transaction)
            targetSpace.AppendEntity(blockRef);
            tr.AddNewlyCreatedDBObject(blockRef, true);
            
            // Synchronize attributes from the block definition
            // This is crucial for dynamic blocks to have their attributes available
            SyncBlockAttributes(db, tr, blockRef, blockDefId);
            
            _logger.LogDebug($"Block reference created successfully with attributes: {blockName}");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error creating block reference: {ex.Message}", ex);
            return false;
        }
    }


    /// <summary>
    /// Synchronizes attributes from the block definition to the block reference
    /// This ensures that dynamic blocks have their attributes properly initialized
    /// </summary>
    private void SyncBlockAttributes(Database db, Transaction tr, BlockReference blockRef, ObjectId blockDefId)
    {
        try
        {
            var blockDef = tr.GetObject(blockDefId, OpenMode.ForRead) as BlockTableRecord;
            if (blockDef == null) return;

            _logger.LogDebug($"Synchronizing attributes for block: {blockDef.Name}");

            // Find all attribute definitions in the block
            foreach (ObjectId objId in blockDef)
            {
                var obj = tr.GetObject(objId, OpenMode.ForRead);
                if (obj is AttributeDefinition attDef)
                {
                    _logger.LogDebug($"Creating attribute reference for: {attDef.Tag}");

                    // Create an attribute reference from the definition
                    var attRef = new AttributeReference();
                    attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                    
                    // Add the attribute reference to the block reference
                    blockRef.AttributeCollection.AppendAttribute(attRef);
                    tr.AddNewlyCreatedDBObject(attRef, true);
                    
                    _logger.LogDebug($"Added attribute reference: {attDef.Tag} = '{attRef.TextString}'");
                }
            }

            _logger.LogDebug($"Attribute synchronization completed for block: {blockDef.Name}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error synchronizing block attributes: {ex.Message}", ex);
        }
    }
}