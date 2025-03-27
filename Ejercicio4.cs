//LMB 3/27/2025
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Npgsql;

[assembly: CommandClass(typeof(ExportarPostGIS.MyCommands))]

namespace ExportarPostGIS
{
    public class MyCommands
    {
        [CommandMethod("ExportarPostGIS")]
        public void ExportarPostGIS()
        {
            // Obtener el documento y el editor activo de AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Configuración de conexión a PostGIS // Credenciales de ejemplo //
            string connString = "Host=localhost;Username=postgres;Password=12345;Database=dwg_to_postGIS";

            // Variables para el informe final
            int exportedEntities = 0;
            List<string> errors = new List<string>();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    // Recorriendo cada entidad en ModelSpace
                    foreach (ObjectId objId in modelSpace)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (ent == null)
                            continue;

                        // Procesar polilíneas (tipo Polyline)
                        if (ent is Polyline pline)
                        {
                            // Convertir la polilínea a formato WKT
                            string wkt = PolylineToWKT(pline);
                            // Insertar en la tabla "polylines" de PostGIS
                            try
                            {
                                InsertGeometryToPostGIS(connString, "polylines", wkt, pline.Layer, "NAD83 Florida East 2236");
                                exportedEntities++;
                            }
                            catch (System.Exception ex)
                            {
                                errors.Add($"Error exportando polilínea (ID: {pline.ObjectId}): {ex.Message}");
                            }
                        }
                        // Procesar bloques (tipo BlockReference)
                        else if (ent is BlockReference blkRef)
                        {
                            // Convertir el bloque a WKT
                            string wkt = BlockReferenceToWKT(blkRef, tr);
                            try
                            {
                                InsertGeometryToPostGIS(connString, "blocks", wkt, blkRef.Layer, "NAD83 Florida East 2236");
                                exportedEntities++;
                            }
                            catch (System.Exception ex)
                            {
                                errors.Add($"Error exportando bloque (ID: {blkRef.ObjectId}): {ex.Message}");
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                errors.Add($"Error general: {ex.Message}");
            }

            // Generar el mensaje final para la línea de comandos de AutoCAD
            StringBuilder report = new StringBuilder();
            report.AppendLine($"\nExportación completada. Entidades exportadas: {exportedEntities}");
            if (errors.Count > 0)
            {
                report.AppendLine("\nErrores durante la exportación:");
                foreach (string err in errors)
                {
                    report.AppendLine($"\n- {err}");
                }
            }

            // Mostrar el reporte en la línea de comandos
            ed.WriteMessage(report.ToString());

            // Escribir el reporte en un archivo de texto
            string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ExportDWGToPostGIS_Log.txt");
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true)) // 'true' para agregar al archivo
                {
                    writer.WriteLine("========== Reporte de Exportación ==========");
                    writer.WriteLine("Autor: Leonardo Martinez");
                    writer.WriteLine($"Fecha: {DateTime.Now}");
                    writer.WriteLine(report.ToString());
                    writer.WriteLine("============================================");
                }
            }
            catch (System.Exception ex)
            {
                // Si ocurre un error al escribir el archivo, se notifica en la línea de comandos
                ed.WriteMessage($"\nError al escribir el archivo de log: {ex.Message}");
            }
        }

        // Método para convertir una Polyline a formato WKT
        private string PolylineToWKT(Polyline pline)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("LINESTRING(");
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                Point2d pt = pline.GetPoint2dAt(i);
                sb.AppendFormat("{0} {1}", pt.X, pt.Y);
                if (i < pline.NumberOfVertices - 1)
                    sb.Append(", ");
            }
            sb.Append(")");
            return sb.ToString();
        }

        // Método para convertir un BlockReference a formato WKT
        private string BlockReferenceToWKT(BlockReference blkRef, Transaction tr)
        {
            // El bloque se representa por su punto de inserción.
            Point3d insPt = blkRef.Position;
            string wkt = $"POINT({insPt.X} {insPt.Y})";
            return wkt;
        }

        // Método para insertar la geometría y atributos en la tabla de PostGIS
        private void InsertGeometryToPostGIS(string connString, string tableName, string wkt, string layer, string sridName)
        {
            // Se asume que la tabla en PostGIS posee las siguientes columnas:
            // id (serial), geom (geometry), layer (text) y srid (text)
            using (NpgsqlConnection conn = new NpgsqlConnection(connString))
            {
                conn.Open();
                // Convertir el nombre del SRS al SRID numérico; NAD83 Florida East 2236 es 2236.
                int srid = GetSRID(sridName);
                string sql = $"INSERT INTO {tableName} (geom, layer, srid) VALUES (ST_SetSRID(ST_GeomFromText(@wkt), {srid}), @layer, @sridName)";
                using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("wkt", wkt);
                    cmd.Parameters.AddWithValue("layer", layer);
                    cmd.Parameters.AddWithValue("sridName", sridName);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // Método auxiliar para obtener el SRID a partir del nombre
        private int GetSRID(string sridName)
        {
            // NAD83 Florida East 2236 corresponde a SRID 2236.
            return 2236;
        }
    }
}
