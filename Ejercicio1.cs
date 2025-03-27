//LMB 3/27/2025
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(InsertarBloquesEnEndpoints))]

public class InsertarBloquesEnEndpoints
{
    [CommandMethod("InsertarBloquesEnEndpoints")]
    public void InsertBlocks()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        // Buscar el bloque a insertar "Tick"
        ObjectId blockId = ObjectId.Null;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            if (bt.Has("Tick"))
            {
                blockId = bt["Tick"];
            }
            tr.Commit();
        }

        if (blockId == ObjectId.Null)
        {
            ed.WriteMessage("\nBloque 'Tick' no encontrado.");
            return;
        }

        // Diccionario para rastrear puntos donde ya se insertaron bloques
        HashSet<Point3d> processedPoints = new HashSet<Point3d>();

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            // Iterar a través de las polilíneas
            TypedValue[] filterlist = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = ed.SelectAll(filter);

            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet selSet = selRes.Value;
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj != null)
                    {
                        Polyline pline = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Polyline;
                        if (pline != null)
                        {
                            InsertBlockAtPoint(pline.StartPoint, modelSpace, blockId, tr, processedPoints);
                            InsertBlockAtPoint(pline.EndPoint, modelSpace, blockId, tr, processedPoints);
                        }
                    }
                }
            }

            tr.Commit();
        }
    }


    private void InsertBlockAtPoint(Point3d point, BlockTableRecord modelSpace, ObjectId blockId, Transaction tr, HashSet<Point3d> processedPoints)
    {
        // Verifica si ya se insertó un bloque en este punto
        if (!processedPoints.Contains(point))
        {
            using (BlockReference blockRef = new BlockReference(point, blockId))
            {
                modelSpace.AppendEntity(blockRef);
                tr.AddNewlyCreatedDBObject(blockRef, true);
                processedPoints.Add(point); // Agrega el punto al conjunto de los ya procesados
            }
        }
    }
}