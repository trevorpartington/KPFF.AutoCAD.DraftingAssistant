using Autodesk.AutoCAD.DatabaseServices;
using KPFF.AutoCAD.DraftingAssistant.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPFF.AutoCAD.DraftingAssistant.Classes.Services
{
    public class BlockProcessor
    {
        public bool ProcessBlock(Transaction trans, BlockReference blockRef, string actualBlockName, string sheetNumber, List<Note> notesForSheet, SheetNamingConfig namingConfig, UpdateResult result, NamingConventionHelper namingConventionHelper)
        {
            int blockNoteNumber = namingConventionHelper.ExtractBlockNoteNumber(actualBlockName, namingConfig);
            if (blockNoteNumber <= 0) 
            {
                return false;
            }

            blockRef.UpgradeOpen();

            try
            {
                Note? noteForBlock = (blockNoteNumber <= notesForSheet.Count) ? notesForSheet[blockNoteNumber - 1] : null;

                bool blockWasModified = false;

                if (blockRef.IsDynamicBlock)
                {
                    bool visibilityChanged = UpdateDynamicBlockProperties(blockRef, noteForBlock);
                    if (visibilityChanged) blockWasModified = true;
                }

                bool attributesChanged = UpdateAttributes(trans, blockRef, noteForBlock, blockNoteNumber);
                if (attributesChanged) blockWasModified = true;

                if (blockWasModified)
                {
                    blockRef.RecordGraphicsModified(true);
                    result.UpdatedBlocks++;
                }

                return blockWasModified;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Error updating block {blockRef.Name}: {ex.Message}");
                return false;
            }
        }

        private bool UpdateDynamicBlockProperties(BlockReference blockRef, Note? noteForBlock)
        {
            bool wasModified = false;
            foreach (DynamicBlockReferenceProperty prop in blockRef.DynamicBlockReferencePropertyCollection)
            {
                if (prop.PropertyName.Equals("Visibility1", StringComparison.OrdinalIgnoreCase) ||
                    prop.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase))
                {
                    string newValue = noteForBlock != null ? "Hex" : "No Hex";
                    string currentValue = prop.Value?.ToString() ?? "";
                    if (currentValue != newValue)
                    {
                        prop.Value = newValue;
                        wasModified = true;
                    }
                }
            }
            return wasModified;
        }

        private bool UpdateAttributes(Transaction trans, BlockReference blockRef, Note? noteForBlock, int blockNoteNumber)
        {
            bool wasModified = false;
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                if (trans.GetObject(attId, OpenMode.ForWrite) is AttributeReference attRef)
                {
                    bool attributeChanged = UpdateAttributeWithAlignment(attRef, noteForBlock, blockNoteNumber);
                    if (attributeChanged) wasModified = true;
                }
            }
            return wasModified;
        }

        private bool UpdateAttributeWithAlignment(AttributeReference attRef, Note? note, int noteNumber)
        {
            string tag = attRef.Tag.ToUpper();
            bool wasModified = false;
            
            switch (tag)
            {
                case "NUM":
                    string newNumValue = note != null ? noteNumber.ToString() : "";
                    string currentNumValue = attRef.TextString ?? "";
                    if (currentNumValue != newNumValue)
                    {
                        attRef.Justify = AttachmentPoint.MiddleCenter;
                        attRef.TextString = newNumValue;
                        attRef.AdjustAlignment(attRef.Database);
                        wasModified = true;
                    }
                    break;
                case "NOTE":
                    string newNoteValue = note?.Text ?? "";
                    string currentNoteValue = attRef.TextString ?? "";
                    if (currentNoteValue != newNoteValue)
                    {
                        attRef.TextString = newNoteValue;
                        attRef.AdjustAlignment(attRef.Database);
                        wasModified = true;
                    }
                    break;
            }
            
            return wasModified;
        }
    }
}
