//LMB 3/27/2025
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;


[assembly: CommandClass(typeof(ClasificarPolilineas))]

public class ClasificarPolilineas
{

    [CommandMethod("ClasificarPolilineas")]
    public void Classify()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        Dictionary<string, int> categoryCounts = new Dictionary<string, int>()
        {
            {"Menor de 200 ft", 0},
            {"Entre 200 y 1000 ft", 0},
            {"Mayor de 1000 ft", 0},
            {"Vertices <= 10", 0},
            {"Vertices > 10", 0},
            {"Cerrada", 0},
            {"Abierta", 0}
        };


        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            TypedValue[] filterlist = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(filterlist);
            PromptSelectionResult selRes = ed.SelectAll(filter);


            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet selSet = selRes.Value;

                foreach (SelectedObject selObj in selSet)
                {
                    Polyline pline = tr.GetObject(selObj.ObjectId, OpenMode.ForWrite) as Polyline;

                    if (pline != null)
                    {
                        double length = pline.Length; 
                        int numVertices = pline.NumberOfVertices;
                        bool isClosed = pline.Closed;


                        if (length < 200) { categoryCounts["Menor de 200 ft"]++; pline.ColorIndex = 1; } // Rojo
                        else if (length >= 200 && length <= 1000) { categoryCounts["Entre 200 y 1000 ft"]++; pline.ColorIndex = 2; } // Amarillo
                        else { categoryCounts["Mayor de 1000 ft"]++; pline.ColorIndex = 3; } // Verde


                        if (numVertices <= 10) { categoryCounts["Vertices <= 10"]++; }
                        else { categoryCounts["Vertices > 10"]++; }


                        if (isClosed) { categoryCounts["Cerrada"]++; }
                        else { categoryCounts["Abierta"]++; }
                    }
                }
            }
            tr.Commit();
        }

        // Obtener la ruta del archivo DWG actual
        string currentDrawingPath = doc.Name;
        string drawingDirectory = Path.GetDirectoryName(currentDrawingPath);

        // Construir las rutas de los archivos de salida
        string excelPath = Path.Combine(drawingDirectory, "PolylineReport.xlsx");
        string txtPath = Path.Combine(drawingDirectory, "PolylineReport.txt");

        // Generar el informe en la misma carpeta del DWG
        GenerateReport(categoryCounts, excelPath, txtPath);


        // Generar el informe (Excel y TXT)
        //GenerateReport(categoryCounts, "PolylineReport.xlsx", "PolylineReport.txt");

        ed.WriteMessage("\nClasificación de polilíneas completada. Informe generado.");

        string currentDrawingDirectory = HostApplicationServices.Current.FindFile(doc.Name, doc.Database, FindFileHint.Default);
        string directory = Path.GetDirectoryName(currentDrawingDirectory);
        ed.WriteMessage($"\nDirectorio del dibujo: {directory}");
    }



    private void GenerateReport(Dictionary<string, int> data, string excelPath, string txtPath)
    {
        // Excel
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Polilíneas");
            int row = 1;
            foreach (var kvp in data)
            {
                worksheet.Cell(row, 1).Value = kvp.Key;
                worksheet.Cell(row, 2).Value = kvp.Value;
                row++;
            }
            workbook.SaveAs(excelPath);
        }


        // TXT
        using (StreamWriter writer = new StreamWriter(txtPath))
        {
            foreach (var kvp in data)
            {
                writer.WriteLine($"{kvp.Key}: {kvp.Value}");
            }
        }
    }
}