using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Text;

namespace RHPro.ReportesAFD.Clases
{
    public class WordImport
    {
        public static void InsertarTabla(string Tag, string FileNameOriginal, string FileNameSalida, DataTable Tabla, int TotalRegistros, bool EsPrimerReemplazo)
        {
            StreamReader reader;
            if (EsPrimerReemplazo)
                reader = new StreamReader(FileNameOriginal);
            else
                reader = new StreamReader(FileNameSalida);


            string StreamOriginal = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(FileNameSalida);

            #region Stream Tabla
            string StreamTabla = "";

            StreamTabla += "<w:tbl>";
            StreamTabla += "    <w:tblPr>";
            StreamTabla += "        <w:tblW w:w=\"0\" w:type=\"auto\"/>";           
            StreamTabla += "        <w:jc w:val=\"center\"/>";
            StreamTabla += "        <w:tblBorders>";
            StreamTabla += "            <w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "            <w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "            <w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "            <w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "            <w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "            <w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamTabla += "        </w:tblBorders>";
            StreamTabla += "    </w:tblPr>";

            StreamTabla += "    <w:tblGrid>";
            foreach (DataColumn dc in Tabla.Columns) StreamTabla += "        <w:gridCol w:w=\"1500\"/>";
            StreamTabla += "    </w:tblGrid>";

            StreamTabla += "    <w:tr>";
            foreach (DataColumn dc in Tabla.Columns)
            {
                StreamTabla += "        <w:tc>";
                StreamTabla += "            <w:tcPr>";
                StreamTabla += "                <w:tcW w:w=\"0\" w:type=\"auto\"/>";               
                StreamTabla += "                <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"CC0000\"/>";
                StreamTabla += "                <w:vAlign w:val=\"center\"/>";
                StreamTabla += "                <w:hAlign w:val=\"center\"/>";
                StreamTabla += "            </w:tcPr>";
                StreamTabla += "            <w:p>";
                StreamTabla += "                <w:r>";
                StreamTabla += "                    <w:pPr>";
                StreamTabla += "                        <w:jc w:val=\"center\"/>";
                StreamTabla += "                        <w:rPr>";
                StreamTabla += "                            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
                StreamTabla += "                            <w:color w:val=\"FFFFFF\"/>";
                StreamTabla += "                            <w:sz w:val=\"18\"/>";
                StreamTabla += "                            <w:szCs w:val=\"18\"/>";
                StreamTabla += "                        </w:rPr>";
                StreamTabla += "                    </w:pPr>";
                StreamTabla += "                    <w:t>" + dc.Caption + "</w:t>";
                StreamTabla += "                </w:r>";
                StreamTabla += "            </w:p>";
                StreamTabla += "        </w:tc>";
            }
            StreamTabla += "    </w:tr>";

            foreach (DataRow dr in Tabla.Rows)
            {
                StreamTabla += "    <w:tr>";
                foreach (DataColumn dc in Tabla.Columns)
                {
                    StreamTabla += "        <w:tc>";
                    StreamTabla += "            <w:tcPr>";
                    StreamTabla += "                <w:tcW w:w=\"0\" w:type=\"auto\"/>";                    
                    StreamTabla += "                <w:vAlign w:val=\"center\"/>";
                    StreamTabla += "                <w:hAlign w:val=\"center\"/>";
                    StreamTabla += "            </w:tcPr>";
                    StreamTabla += "            <w:p>";
                    StreamTabla += "                <w:r>";
                    StreamTabla += "                    <w:pPr>";
                    StreamTabla += "                        <w:jc w:val=\"center\"/>";
                    StreamTabla += "                        <w:rPr>";
                    StreamTabla += "                            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
                    StreamTabla += "                            <w:sz w:val=\"18\"/>";
                    StreamTabla += "                            <w:szCs w:val=\"18\"/>";
                    StreamTabla += "                        </w:rPr>";
                    StreamTabla += "                    </w:pPr>";
                    StreamTabla += "                    <w:t>" + dr[dc.ColumnName].ToString().Replace("&", "&amp;").Replace("<", "&lt;").Replace(">","&gt;") + "</w:t>";
                    StreamTabla += "                </w:r>";
                    StreamTabla += "            </w:p>";
                    StreamTabla += "        </w:tc>";
                }
                StreamTabla += "    </w:tr>";
                StreamTabla += Tag;
                StreamOriginal = StreamOriginal.Replace(Tag, StreamTabla);
                StreamTabla = "";
            }
            StreamTabla += "</w:tbl>";

            StreamTabla += "<w:p w:rsidR=\"007F5A66\" w:rsidRDefault=\"007F5A66\" w:rsidP=\"007F5A66\">";
            StreamTabla += "    <w:pPr>";
            StreamTabla += "        <w:jc w:val=\"both\"/>";
            StreamTabla += "        <w:rPr>";
            StreamTabla += "            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
            StreamTabla += "            <w:b/>";
            StreamTabla += "            <w:bCs/>";
            StreamTabla += "            <w:color w:val=\"3366FF\"/>";
            StreamTabla += "            <w:sz w:val=\"18\"/>";
            StreamTabla += "            <w:szCs w:val=\"18\"/>";
            StreamTabla += "            <w:u w:val=\"single\"/>";
            StreamTabla += "            <w:lang w:val=\"es-AR\" w:eastAsia=\"es-ES\"/>";
            StreamTabla += "            </w:rPr>";
            StreamTabla += "    </w:pPr>";
            StreamTabla += "</w:p>";
            StreamTabla += "<w:p w:rsidR=\"00F13363\" w:rsidRPr=\"00F13363\" w:rsidRDefault=\"00F13363\" w:rsidP=\"007F5A66\">";
            StreamTabla += "    <w:pPr>";
            StreamTabla += "        <w:jc w:val=\"center\"/>";
            StreamTabla += "        <w:rPr>";
            StreamTabla += "            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
            StreamTabla += "            <w:bCs/>";
            StreamTabla += "            <w:sz w:val=\"16\"/>";
            StreamTabla += "            <w:szCs w:val=\"16\"/>";
            StreamTabla += "            <w:lang w:val=\"es-AR\" w:eastAsia=\"es-ES\"/>";
            StreamTabla += "        </w:rPr>";
            StreamTabla += "    </w:pPr>";
            StreamTabla += "    <w:r w:rsidRPr=\"00F13363\">";
            StreamTabla += "        <w:rPr>";
            StreamTabla += "            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
            StreamTabla += "            <w:bCs/>";
            StreamTabla += "            <w:sz w:val=\"16\"/>";
            StreamTabla += "            <w:szCs w:val=\"16\"/>";
            StreamTabla += "            <w:lang w:val=\"es-AR\" w:eastAsia=\"es-ES\"/>";
            StreamTabla += "        </w:rPr>";
            StreamTabla += "        <w:t>Cantidad de Registros mostrados: " + Tabla.Rows.Count + " de " + TotalRegistros + ".</w:t>";
            StreamTabla += "    </w:r>";
            StreamTabla += "</w:p>";
            StreamOriginal = StreamOriginal.Replace(Tag, StreamTabla);
            StreamTabla = "";
            #endregion

            #region Stream Sin Registros
            string StreamSinRegistros = "";

            StreamSinRegistros += "<w:tbl>";
            StreamSinRegistros += "    <w:tblPr>";
            StreamSinRegistros += "        <w:tblW w:w=\"0\" w:type=\"auto\"/>";           
            StreamSinRegistros += "        <w:jc w:val=\"center\"/>";
            StreamSinRegistros += "        <w:tblBorders>";
            StreamSinRegistros += "            <w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "            <w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "            <w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "            <w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "            <w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "            <w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>";
            StreamSinRegistros += "        </w:tblBorders>";
            StreamSinRegistros += "    </w:tblPr>";
            StreamSinRegistros += "    <w:tblGrid>";
            StreamSinRegistros += "        <w:gridCol w:w=\"4500\"/>";
            StreamSinRegistros += "    </w:tblGrid>";
            StreamSinRegistros += "    <w:tr>";
            StreamSinRegistros += "        <w:tc>";
            StreamSinRegistros += "            <w:tcPr>";            
            StreamSinRegistros += "                <w:tcW w:w=\"0\" w:type=\"auto\"/>";            
            StreamSinRegistros += "                <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"CC0000\"/>";
            StreamSinRegistros += "                <w:vAlign w:val=\"center\"/>";
            StreamSinRegistros += "                <w:hAlign w:val=\"center\"/>";
            StreamSinRegistros += "            </w:tcPr>";
            StreamSinRegistros += "            <w:p>";
            StreamSinRegistros += "                <w:r>";
            StreamSinRegistros += "                    <w:pPr>";
            StreamSinRegistros += "                        <w:jc w:val=\"center\"/>";
            StreamSinRegistros += "                        <w:rPr>";
            StreamSinRegistros += "                            <w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>";
            StreamSinRegistros += "                            <w:color w:val=\"FFFFFF\"/>";
            StreamSinRegistros += "                            <w:sz w:val=\"22\"/>";
            StreamSinRegistros += "                            <w:szCs w:val=\"22\"/>";
            StreamSinRegistros += "                        </w:rPr>";
            StreamSinRegistros += "                    </w:pPr>";
            StreamSinRegistros += "                    <w:t>" + "No se han Encontrado Registros" + "</w:t>";
            StreamSinRegistros += "                </w:r>";
            StreamSinRegistros += "            </w:p>";
            StreamSinRegistros += "        </w:tc>";
            StreamSinRegistros += "    </w:tr>";
            StreamSinRegistros += "</w:tbl>";
            #endregion


            //if (Tabla.Rows.Count != 0) StreamOriginal = StreamOriginal.Replace(Tag, StreamTabla);
            //else StreamOriginal = StreamOriginal.Replace(Tag, StreamSinRegistros);

            if (Tabla.Rows.Count != 0) StreamOriginal = StreamOriginal.Replace(Tag, "");
            else StreamOriginal = StreamOriginal.Replace(Tag, StreamSinRegistros);


            writer.Write(StreamOriginal);

            writer.Close();
        }
    }
}
