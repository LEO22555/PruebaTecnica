//LMB 3/27/2025
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;


[assembly: CommandClass(typeof(ExportarDXF))]

public class ExportarDXF
{
    [CommandMethod("ExportarDXF")]
    public void Export()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        string dwgPath = HostApplicationServices.Current.FindFile(doc.Name, doc.Database, FindFileHint.Default);
        string dxfPath = Path.ChangeExtension(dwgPath, ".dxf");

        try
        {
            using (Database destDb = new Database(true, true))
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                    // Obtener los ObjectIds de las entidades en el espacio modelo
                    ObjectIdCollection objectIds = new ObjectIdCollection(modelSpace.Cast<ObjectId>().ToArray());


                    IdMapping mapping = new IdMapping();
                    db.WblockCloneObjects(objectIds,
                                         destDb.CurrentSpaceId,
                                         mapping,
                                         DuplicateRecordCloning.Ignore,
                                         false);
                    tr.Commit();
                }

                destDb.DxfOut(dxfPath, 16, true);
                
                ed.WriteMessage($"\nArchivo DXF exportado a: {dxfPath}");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\nError al exportar a DXF: {ex.Message}");
        }
    }
}