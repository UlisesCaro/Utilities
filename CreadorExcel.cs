using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Utilities
{
    public class ServicioCrearExcel
    {
        private readonly string Extension;
        private readonly string NombreDeArchivoPorDefecto;
        private readonly string NombreDeHojaPorDefecto;
        private readonly string DirectorioTemporal;
        private readonly int LargoMaximoDeHoja;

        public ServicioCrearExcel(DirectoryInfo directorio = null)
        {
            var directorioTemporal = directorio == null || !directorio.Exists ? Path.GetTempPath() : directorio.FullName;
            if (!directorioTemporal.EndsWith("\\"))
                directorioTemporal += "\\";
            DirectorioTemporal = directorioTemporal;

            Extension = ".xlsx";
            NombreDeArchivoPorDefecto = "Libro1";
            NombreDeHojaPorDefecto = "Hoja";

            LargoMaximoDeHoja = 31;
            FontIndex = 0;
            FontIndexEncabezado = 2;
            FontIndexDetalle = 1;
            FontIndexTitulo = 3;

            FillIndex = 0;
            FillIndexEncabezado = 2;
            FillIndexDetalle = 0;
            FillIndexTitulo = 3;

            BorderIndex = 0;
            BorderIndexTodo = 1;
        }


        public uint FontIndex { get; set; }
        public uint FontIndexEncabezado { get; set; }
        public uint FontIndexDetalle { get; set; }
        public uint FontIndexTitulo { get; set; }
        public uint FillIndex { get; set; }
        public uint FillIndexEncabezado { get; set; }
        public uint FillIndexDetalle { get; set; }
        public uint FillIndexTitulo { get; set; }
        public uint BorderIndex { get; set; }
        public uint BorderIndexTodo { get; set; }

        public FileInfo Exportar(IEnumerable<object> objetos, string nombreDeArchivo = "", string nombreDeHoja = "", Dictionary<string, string> nombresDeColumna = null, string titulo = "")
        {
            var nombreDeArchivoDeducido = DeducirNombreDeArchivo(ref objetos, nombreDeArchivo);
            var nombreDeArchivoFisicoDeducido = DeducirNombreDeArchivoFisico(nombreDeArchivoDeducido);

            using (var spreadsheetDocument = SpreadsheetDocument.Create(nombreDeArchivoFisicoDeducido.FullName, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheetPartId = workbookPart.GetIdOfPart(worksheetPart);

                var formatosDeCelda = ObtenerFormatosDeCeldaPorDefecto();
                var usarTitulo = !string.IsNullOrWhiteSpace(titulo);

                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                var hojaDeEstilos = CrearHojaDeEstilos(formatosDeCelda, usarTitulo);
                workbookStylesPart.Stylesheet = hojaDeEstilos;
                workbookStylesPart.Stylesheet.Save();

                var sheets = new Sheets();
                var sheet = new Sheet { Name = DeducirNombreDeHoja(nombreDeHoja), SheetId = 1, Id = worksheetPartId };
                sheets.Append(sheet);

                if ((nombresDeColumna == null || nombresDeColumna.Count.Equals(0)) && objetos.Any())
                    nombresDeColumna = ObtenerNombresDeColumnaDesdeObjeto(objetos.First());
                var sheetData = CrearHojaDeDatos(objetos, hojaDeEstilos, formatosDeCelda, nombresDeColumna, titulo);

                var workbook = new Workbook();
                workbook.Append(sheets);

                var columns = AjustarTamañoColumnas(nombresDeColumna, sheetData, usarTitulo);
                var worksheet = new Worksheet();

                worksheet.Append(columns);
                worksheet.Append(sheetData);

                worksheetPart.Worksheet = worksheet;
                if (usarTitulo)
                {
                    var celdasCombinadas = new MergeCells();
                    var hasta = columns.Elements<Column>().Count() - 1;
                    celdasCombinadas.AppendChild(new MergeCell { Reference = new StringValue(DeterminarLetraDeColumna(0) + "1:" + DeterminarLetraDeColumna(hasta) + "1") });
                    worksheetPart.Worksheet.InsertAfter(celdasCombinadas, worksheetPart.Worksheet.Elements<SheetData>().First());
                }

                worksheetPart.Worksheet.Save();

                spreadsheetDocument.WorkbookPart.Workbook = workbook;
                spreadsheetDocument.WorkbookPart.Workbook.Save();
                spreadsheetDocument.Close();
            }

            return nombreDeArchivoFisicoDeducido;
        }

        #region >>>> Metodos de Creación de datos

        private SheetData CrearHojaDeDatos(IEnumerable<object> objetos, Stylesheet hojaDeEstilos, Dictionary<string, OpenXml_NumberFormatId> formatosDeCeldaUtilizados, Dictionary<string, string> nombresDeColumna = null, string titulo = "")
        {
            var openXmlElements = new List<OpenXmlElement>();
            var formatosCelda = hojaDeEstilos.CellFormats.Elements<CellFormat>().ToList();

            if (!string.IsNullOrWhiteSpace(titulo))
                openXmlElements.Add(CrearFilaTitulo(titulo, formatosCelda, formatosDeCeldaUtilizados, FontIndexTitulo, FillIndexTitulo, BorderIndexTodo));

            openXmlElements.Add(CrearFilaEncabezado(nombresDeColumna, formatosCelda, formatosDeCeldaUtilizados, FontIndexEncabezado, FillIndexEncabezado, BorderIndexTodo));
            foreach (var objeto in objetos)
                openXmlElements.Add(CrearFila(objeto, nombresDeColumna, formatosCelda, formatosDeCeldaUtilizados, FontIndexDetalle, FillIndexDetalle, BorderIndexTodo));

            var hojaDeDatos = new SheetData(openXmlElements);

            return hojaDeDatos;
        }
        private Dictionary<string, string> ObtenerNombresDeColumnaDesdeObjeto(object objeto)
        {
            var nombresColumna = new Dictionary<string, string>();
            var propiedades = ServicioRefleccion.ObtenerPropiedades(objeto);
            foreach (var propiedad in propiedades)
            {
                //Más adelante se podrán leer atributos
                nombresColumna.Add(propiedad.Name, propiedad.Name);
            }
            return nombresColumna;
        }
        private Row CrearFilaTitulo(string titulo, IList<CellFormat> formatosDeCelda, Dictionary<string, OpenXml_NumberFormatId> formatosDeCeldaUtilizados, uint? indiceFormatoFuente, uint? indiceFormatoRelleno, uint? indiceFormatoBordes)
        {
            var fila = new Row();
            var celdas = new Cell[1];
            celdas[0] = CrearCelda(titulo, formatosDeCelda, formatosDeCeldaUtilizados, indiceFormatoFuente, indiceFormatoRelleno, indiceFormatoBordes, null);
            fila.Append(celdas);
            return fila;
        }
        private Row CrearFilaEncabezado(Dictionary<string, string> nombresDeColumna, IList<CellFormat> formatosDeCelda, Dictionary<string, OpenXml_NumberFormatId> formatosDeCeldaUtilizados, uint? indiceFormatoFuente, uint? indiceFormatoRelleno, uint? indiceFormatoBordes)
        {
            var fila = new Row();
            var celdas = new Cell[nombresDeColumna.Count];
            for (var i = 0; i < nombresDeColumna.Count; i++)
                celdas[i] = CrearCelda(nombresDeColumna.ElementAt(i).Value
                    , formatosDeCelda
                    , formatosDeCeldaUtilizados
                    , FontIndexEncabezado
                    , FillIndexEncabezado
                    , BorderIndexTodo
                    , null);
            fila.Append(celdas);
            return fila;
        }
        private Row CrearFila(object objeto, Dictionary<string, string> nombresDeColumna, IList<CellFormat> formatosDeCelda, Dictionary<string, OpenXml_NumberFormatId> formatosDeCeldaUtilizados, uint? indiceFormatoFuente, uint? indiceFormatoRelleno, uint? indiceFormatoBordes)
        {
            var fila = new Row();
            var celdas = new Cell[nombresDeColumna.Count];

            int? idMoneda = null;
            var propiedaes = ServicioRefleccion.ObtenerPropiedades(objeto);
            var propiedadesMoneda = propiedaes.Where(w => w.Name.ToLower().Contains("idmoneda") && w.PropertyType.FullName == typeof(int).FullName).ToList();
            var dd = propiedadesMoneda.Select(s => (int)ServicioRefleccion.ObtenerValorDePropiedad(objeto, s.Name)).ToList();
            if (dd.Any())
                idMoneda = dd.GroupBy(g => g).First().Key;
            for (var i = 0; i < nombresDeColumna.Count; i++)
            {
                celdas[i] = CrearCelda(ServicioRefleccion.ObtenerValorDePropiedad(objeto, nombresDeColumna.ElementAt(i).Key)
                    , formatosDeCelda
                    , formatosDeCeldaUtilizados
                    , indiceFormatoFuente
                    , indiceFormatoRelleno
                    , indiceFormatoBordes
                    , idMoneda);
            }
            fila.Append(celdas);
            return fila;
        }
        private Cell CrearCelda(object valor, IList<CellFormat> formatosDeCelda, Dictionary<string, OpenXml_NumberFormatId> formatosDeCeldaUtilizados, uint? indiceFormatoFuente, uint? indiceFormatoRelleno, uint? indiceFormatoBordes, int? idMonedaFila)
        {
            var tipo = valor == null ? typeof(string) : valor.GetType();

            var idFormatoDeCeldaEntero = OpenXml_NumberFormatId.Numero;
            if (formatosDeCeldaUtilizados.ContainsKey(typeof(int).FullName))
                idFormatoDeCeldaEntero = formatosDeCeldaUtilizados[typeof(int).FullName];

            var idFormatoDeCelda = OpenXml_NumberFormatId.Ninguno;
            if (formatosDeCeldaUtilizados.ContainsKey(tipo.FullName))
                idFormatoDeCelda = formatosDeCeldaUtilizados[tipo.FullName];

            var celda = new Cell();
            celda.StyleIndex = ObtenerEstiloDeCelda(idFormatoDeCelda, formatosDeCelda, indiceFormatoFuente, indiceFormatoRelleno, indiceFormatoBordes);
            var culture = CultureInfo.CurrentCulture;
            switch (tipo.Name)
            {
                case "DateTime":
                    celda.DataType = CellValues.Number;
                    celda.CellValue = new CellValue(((DateTime)valor).ToOADate().ToString(CultureInfo.InvariantCulture));
                    break;
                case "Boolean":
                    celda.DataType = CellValues.String;
                    if (culture == CultureInfo.CurrentUICulture)
                        celda.CellValue = new CellValue(((bool)valor) ? "Sí" : "No");
                    else
                        celda.CellValue = new CellValue(((bool)valor)? "Yes" : "No");
                    break;
                case "Int32":
                    celda.DataType = CellValues.Number;
                    celda.CellValue = new CellValue(Convert.ToInt32(valor).ToString(CultureInfo.InvariantCulture));
                    break;
                case "Decimal":
                    celda.DataType = CellValues.Number;
                    var valorDecimal = Convert.ToDecimal(valor);
                    var redondeo = 2;
                    var decArray = valorDecimal.ToString().Split(new[] { '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (decArray.Length.Equals(1) || (decArray.Length > 1 && (decimal.Parse(decArray[1]).Equals(0) || idMonedaFila==1)))
                    {
                        var idFormato = idMonedaFila == null || idMonedaFila == 1 ? OpenXml_NumberFormatId.Moneda_SinDecimales : OpenXml_NumberFormatId.Moneda_Decimales;
                        celda.StyleIndex = ObtenerEstiloDeCelda(idFormato, formatosDeCelda, indiceFormatoFuente, indiceFormatoRelleno, indiceFormatoBordes);
                        var ss = decimal.Parse(decArray[0]).ToString(culture);
                        celda.CellValue = new CellValue(ss);
                    }
                    else
                        celda.CellValue = new CellValue(decimal.Round(valorDecimal, redondeo).ToString(CultureInfo.InvariantCulture));
                    break;
                default:
                    int d = 0; bool esSoloNumeros = int.TryParse((string)valor, out d);
                    if (esSoloNumeros)
                    {
                        celda.DataType = CellValues.Number;
                        celda.StyleIndex = ObtenerEstiloDeCelda(idFormatoDeCeldaEntero, formatosDeCelda, indiceFormatoFuente, indiceFormatoRelleno, indiceFormatoBordes);
                        celda.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        celda.DataType =  CellValues.String;
                        celda.CellValue = new CellValue(valor == null ? "" : LimpiarCaracteresNoValidos(valor.ToString()));
                    }
                    break;
            }
            return celda;
        }
        public string LimpiarCaracteresNoValidos(string text)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(text, r, "", RegexOptions.Compiled);
        }
        private uint ObtenerEstiloDeCelda(OpenXml_NumberFormatId idFormatoDeCelda, IList<CellFormat> formatosDeCelda, uint? indiceFormatoFuente, uint? indiceFormatoRelleno, uint? indiceFormatoBordes)
        {
            int idEstiloCelda = 0;
            for (var i = 0; i < formatosDeCelda.Count; i++)
            {
                if (formatosDeCelda[i].NumberFormatId == null && !idFormatoDeCelda.Equals(OpenXml_NumberFormatId.Ninguno)) continue;
                if (formatosDeCelda[i].NumberFormatId != null && idFormatoDeCelda.Equals(OpenXml_NumberFormatId.Ninguno)) continue;
                if (formatosDeCelda[i].NumberFormatId != null && !formatosDeCelda[i].NumberFormatId.Value.Equals((uint)idFormatoDeCelda)) continue;

                if (formatosDeCelda[i].FontId == null && indiceFormatoFuente.HasValue) continue;
                if (formatosDeCelda[i].FontId != null && !indiceFormatoFuente.HasValue) continue;
                if (formatosDeCelda[i].FontId != null && indiceFormatoFuente.HasValue && !formatosDeCelda[i].FontId.Value.Equals(indiceFormatoFuente.Value)) continue;

                if (formatosDeCelda[i].FillId == null && indiceFormatoRelleno.HasValue) continue;
                if (formatosDeCelda[i].FillId != null && !indiceFormatoRelleno.HasValue) continue;
                if (formatosDeCelda[i].FillId != null && indiceFormatoRelleno.HasValue && !formatosDeCelda[i].FillId.Value.Equals(indiceFormatoRelleno.Value)) continue;

                if (formatosDeCelda[i].BorderId == null && indiceFormatoBordes.HasValue) continue;
                if (formatosDeCelda[i].BorderId != null && !indiceFormatoBordes.HasValue) continue;
                if (formatosDeCelda[i].BorderId != null && indiceFormatoBordes.HasValue && !formatosDeCelda[i].BorderId.Value.Equals(indiceFormatoBordes.Value)) continue;

                idEstiloCelda = i;
                break;
            }
            return (uint)idEstiloCelda;
        }
        private Columns AjustarTamañoColumnas(Dictionary<string, string> nombresDeColumna, SheetData sheetData, bool usaTitulo)
        {
            var rows = sheetData.Elements<Row>().ToList();
            if (rows.Any() && usaTitulo)
                rows.RemoveAt(0);

            var largoMaximoDeColumnas = new List<int>();
            if (rows.Any())
                for (var i = 0; i < nombresDeColumna.Count; i++)
                    largoMaximoDeColumnas.Add(rows.Max(r => { var celda = r.Elements<Cell>().ToArray()[i]; return celda == null ? 0 : celda.InnerText.Length; }));

            var columns = new Columns();
            for (var i = 0; i < largoMaximoDeColumnas.Count(); i++)
            {
                double width = ((largoMaximoDeColumnas[i] * 5 + 5) / 5 * 256) / 256;
                var largoDeColumna = (double)decimal.Round((decimal)width + 0.2M, 2);
                columns.AppendChild(new Column()
                {
                    BestFit = true,
                    Min = (uint)(i + 1),
                    Max = (uint)(i + 1),
                    CustomWidth = true,
                    Width = largoDeColumna
                });
            }
            return columns;
        }
        private string DeducirNombreDeHoja(string nombreDeArchivo)
        {
            if (!string.IsNullOrWhiteSpace(nombreDeArchivo))
                return nombreDeArchivo.Length > LargoMaximoDeHoja ? nombreDeArchivo.Substring(0, LargoMaximoDeHoja) : nombreDeArchivo;
            else
                return NombreDeHojaPorDefecto;
        }
        public string DeterminarLetraDeColumna(int index)
        {
            var letras = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            if (index < letras.Length)
                return letras[index];
            var entero = (index / letras.Length) - 1;
            var decimales = index % letras.Length;
            if (entero < letras.Length)
                return letras[entero] + DeterminarLetraDeColumna(decimales);
            else
                return DeterminarLetraDeColumna(entero) + DeterminarLetraDeColumna(decimales);
        }

        #endregion

        #region >>>> Metodos de Aplicación de estilos

        private Stylesheet CrearHojaDeEstilos(Dictionary<string, OpenXml_NumberFormatId> formatos, bool usarTitulo)
        {
            var hojaDeEstilos = new Stylesheet();
            hojaDeEstilos.Fonts = CrearFuentes(usarTitulo);
            hojaDeEstilos.Fills = CrearRellenos(usarTitulo);
            hojaDeEstilos.Borders = CrearBordes();
            hojaDeEstilos.CellStyleFormats = CrearFormatosDeEstiloDeCelda();
            hojaDeEstilos.CellFormats = CrearFormatosDeCelda(formatos, usarTitulo);
            hojaDeEstilos.NumberingFormats = AgregarFormatosNumericos();
            return hojaDeEstilos;
        }



        private Fonts CrearFuentes(bool usarTitulo = false)
        {
            var fuentes = new Fonts();
            fuentes.AppendChild(CrearFuente()); //0 - Requerido, reservado por Excel
            fuentes.AppendChild(CrearFuente(11)); //1 - Detalle
            fuentes.AppendChild(CrearFuente(11, true, "FFFFFF")); //2 - Encabezado
            fuentes.Count = 3;
            if (!usarTitulo)
                return fuentes;
            fuentes.Append(CrearFuente(14, true, "000080")); //3 - Titulo (opcional) //FF8000
            fuentes.Count = 4;
            return fuentes;
        }
        private Font CrearFuente(int? tamaño = null, bool negrita = false, string rgbColor = "")
        {
            var fuente = new Font();
            if (tamaño.HasValue)
                fuente.FontSize = new FontSize { Val = tamaño.Value };
            if (negrita)
                fuente.Bold = new Bold();
            if (rgbColor != null || rgbColor != string.Empty)
                fuente.Color = new Color { Rgb = rgbColor };
            return fuente;
        }
        private Fills CrearRellenos(bool usarTitulo = false)
        {
            var rellenos = new Fills();
            rellenos.AppendChild(CrearRelleno()); //0 - Requerido, reservado por Excel
            rellenos.AppendChild(CrearRelleno(PatternValues.Gray125)); //1 - Requerido, reservado por Excel 
            rellenos.AppendChild(CrearRelleno(PatternValues.Solid, "63A7EB")); //2 - Encabezado
            rellenos.Count = 3;
            if (!usarTitulo)
                return rellenos;
            rellenos.AppendChild(CrearRelleno(PatternValues.Solid, "e6e6e6")); //3 - Titulo (opcional)
            rellenos.Count = 4;
            return rellenos;
        }
        private Fill CrearRelleno(PatternValues pattern = PatternValues.None, string rgbColor = "")
        {
            var relleno = new Fill();
            relleno.PatternFill = new PatternFill();
            relleno.PatternFill.PatternType = pattern;
            if (!string.IsNullOrWhiteSpace(rgbColor))
                relleno.PatternFill.ForegroundColor = new ForegroundColor { Rgb = new HexBinaryValue { Value = rgbColor } };
            return relleno;
        }
        private Borders CrearBordes()
        {
            var bordes = new Borders();
            bordes.AppendChild(CrearBorde()); //0 - Requerido, reservado por Excel
            bordes.AppendChild(CrearBorde(BorderStyleValues.Thin, BorderStyleValues.Thin, BorderStyleValues.Thin, BorderStyleValues.Thin)); //1 - Todos los bordes
            bordes.Count = 2;
            return bordes;
        }
        private Border CrearBorde(BorderStyleValues arriba = BorderStyleValues.None, BorderStyleValues derecha = BorderStyleValues.None, BorderStyleValues abajo = BorderStyleValues.None, BorderStyleValues izquierda = BorderStyleValues.None)
        {
            var borde = new Border();
            if (!arriba.Equals(BorderStyleValues.None))
                borde.TopBorder = new TopBorder(new Color { Auto = true }) { Style = arriba };
            if (!derecha.Equals(BorderStyleValues.None))
                borde.RightBorder = new RightBorder(new Color { Auto = true }) { Style = derecha };
            if (!abajo.Equals(BorderStyleValues.None))
                borde.BottomBorder = new BottomBorder(new Color { Auto = true }) { Style = abajo };
            if (!izquierda.Equals(BorderStyleValues.None))
                borde.LeftBorder = new LeftBorder(new Color { Auto = true }) { Style = izquierda };
            return borde;
        }
        private CellStyleFormats CrearFormatosDeEstiloDeCelda()
        {
            var formatosDeEstiloDeCelda = new CellStyleFormats();
            formatosDeEstiloDeCelda.AppendChild(new CellFormat());
            formatosDeEstiloDeCelda.Count = 1;
            return formatosDeEstiloDeCelda;
        }
        private CellFormats CrearFormatosDeCelda(Dictionary<string, OpenXml_NumberFormatId> formatos, bool usarTitulo)
        {
            var formatosDeCelda = new CellFormats();
            formatosDeCelda.Count = 0;
            var formatosDeCeldaPorDefecto = CrearFormatosDeCelda(formatos);
            foreach (var formatoDeCeldaPorDefecto in formatosDeCeldaPorDefecto)
            {
                formatosDeCelda.AppendChild(formatoDeCeldaPorDefecto);
                formatosDeCelda.Count += 1;
            }
            var formatosDeCeldaEncabezado = CrearFormatosDeCelda(formatos, FontIndexEncabezado, FillIndexEncabezado, BorderIndexTodo);
            foreach (var formatoDeCeldaEncabezado in formatosDeCeldaEncabezado)
            {
                formatosDeCelda.AppendChild(formatoDeCeldaEncabezado);
                formatosDeCelda.Count += 1;
            }
            var formatosDeCeldaDetalle = CrearFormatosDeCelda(formatos, FontIndexDetalle, FillIndexDetalle, BorderIndexTodo);
            foreach (var formatoDeCeldaDetalle in formatosDeCeldaDetalle)
            {
                formatosDeCelda.AppendChild(formatoDeCeldaDetalle);
                formatosDeCelda.Count += 1;
            }
            if (!usarTitulo)
                return formatosDeCelda;
            var formatosDeCeldaTitulo = CrearFormatosDeCelda(formatos, FontIndexTitulo, FillIndexTitulo, BorderIndexTodo);
            foreach (var formatoDeCeldaTitulo in formatosDeCeldaTitulo)
            {
                formatosDeCelda.AppendChild(formatoDeCeldaTitulo);
                formatosDeCelda.Count += 1;
            }
            return formatosDeCelda;
        }
        private List<CellFormat> CrearFormatosDeCelda(Dictionary<string, OpenXml_NumberFormatId> formatos, uint? fontIndex = null, uint? fillIndex = null, uint? borderIndex = null)
        {
            var formatosDeCelda = new List<CellFormat>();
            foreach (var formato in formatos)
                formatosDeCelda.Add(CrearFormatoDeCelda(formato, fontIndex, fillIndex, borderIndex));
            return formatosDeCelda;
        }
        private CellFormat CrearFormatoDeCelda(KeyValuePair<string, OpenXml_NumberFormatId> formato, uint? fontIndex = null, uint? fillIndex = null, uint? borderIndex = null)
        {
            var formatoDeCelda = new CellFormat();
            if (formato.Value != OpenXml_NumberFormatId.Ninguno)
            {
                formatoDeCelda.NumberFormatId = (uint)formato.Value;
                formatoDeCelda.ApplyNumberFormat = true;
            }
            else
            {
                formatoDeCelda.NumberFormatId = null;
                formatoDeCelda.ApplyNumberFormat = false;

            }

            if (fontIndex.HasValue) { formatoDeCelda.FontId = fontIndex.Value; formatoDeCelda.ApplyFont = true; }
            if (fillIndex.HasValue) { formatoDeCelda.FillId = fillIndex.Value; formatoDeCelda.ApplyFill = true; }
            if (borderIndex.HasValue) { formatoDeCelda.BorderId = borderIndex.Value; formatoDeCelda.ApplyBorder = true; }

            return formatoDeCelda;
        }
        private NumberingFormats AgregarFormatosNumericos()
        {
            var formats = new NumberingFormats();
            var nformat0Decimal = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(165),
                FormatCode = StringValue.FromString("##,##0")
            };
            formats.Append(nformat0Decimal);

            formats.Count = UInt32Value.FromUInt32((uint)formats.ChildElements.Count);
            return formats;
        }
        private Dictionary<string, OpenXml_NumberFormatId> ObtenerFormatosDeCeldaPorDefecto()
        {
            var formatos = new Dictionary<string, OpenXml_NumberFormatId>();
            formatos.Add("System.String", OpenXml_NumberFormatId.Ninguno);
            formatos.Add("System.Int32", OpenXml_NumberFormatId.Numero);
            formatos.Add("System.Decimal", OpenXml_NumberFormatId.Moneda_Decimales);
            formatos.Add("System.DateTime", OpenXml_NumberFormatId.Fecha_DiaMesAño);
            formatos.Add("SinDecimal", OpenXml_NumberFormatId.Moneda_SinDecimales);

            return formatos;
        }

        #endregion

        #region >>>> Metodos de Trabajo con archivos

        private string DeducirNombreDeArchivo(ref IEnumerable<object> objeto, string nombreDeArchivo)
        {
            if (!string.IsNullOrWhiteSpace(nombreDeArchivo))
                return nombreDeArchivo;
            return NombreDeArchivoPorDefecto;
        }
        private FileInfo DeducirNombreDeArchivoFisico(string nombreDeArchivoFisicoDeducido)
        {
            if (!nombreDeArchivoFisicoDeducido.ToLower().EndsWith(Extension))
                nombreDeArchivoFisicoDeducido += Extension;
            var file = new FileInfo(DirectorioTemporal + nombreDeArchivoFisicoDeducido);
            var i = 1;
            while (file.Exists)
            {
                file = new FileInfo(DirectorioTemporal + nombreDeArchivoFisicoDeducido.Split(new[] { Extension }, StringSplitOptions.None)[0] + "(" + i.ToString() + ")" + Extension);
                i++;
            }
            return file;
        }

        #endregion
    }

    public enum OpenXml_NumberFormatId
    {
        Ninguno = 0,
        ///<summary>"0"</summary>
        Numero = 1,
        ///<summary>"0.00"</summary>
        Numero_Decimales = 2,
        ///<summary>"#.##0"</summary>
        Moneda = 3,
        ///<summary>"#.##0.00"</summary>
        Moneda_Decimales = 4,
        ///<summary>"0%"</summary>
        Porcentaje = 9,
        ///<summary>"0.00%"</summary>
        Porcentaje_Decimales = 10,
        ///<summary>"0.00E+00"</summary>
        Cientifica = 11,
        ///<summary>"# ?/?"</summary>
        Fraccion = 12,
        ///<summary>"# ??/??"</summary>
        Fraccion_Digitos = 13,
        ///<summary>"d/m/yyyy"</summary>
        Fecha_DiaMesAño = 14,
        ///<summary>"d-mmm-yy"</summary>
        Fecha_DiaNombreMesAño = 15,
        ///<summary>"d-mmm"</summary>
        Fecha_DiaNombreMes = 16,
        ///<summary>"mmm-yy"</summary>
        Fecha_NombreMesYAño = 17,
        ///<summary>"h:mm tt"</summary>
        Tiempo_HoraMinutosAMPM = 18,
        ///<summary>"h:mm:ss tt"</summary>
        Tiempo_HoraMinutosSegundosAMPM = 19,
        ///<summary>"H:mm"</summary>
        Tiempo_HoraMinutos = 20,
        ///<summary>"H:mm:ss"</summary>
        Tiempo_HoraMinutosSegundos = 21,
        ///<summary>"m/d/yyyy H:mm"</summary>
        Fecha_MesDiaAñoHoraMinutos = 22,
        ///<summary>"#,##0 ;(#,##0)"</summary>
        Moneda_NegativoParentesis = 37,
        ///<summary>"#,##0 ;[Red](#,##0)"</summary>
        Modeda_NegativoParenesisRojo = 38,
        ///<summary>"#,##0.00;(#,##0.00)"</summary>
        Moneda_Decimales_NegativoParentesis = 39,
        ///<summary>"#,##0.00;[Red](#,##0.00)"</summary>
        Moneda_Decimales_NegativoParentesisRojo = 40,
        ///<summary>"mm:ss"</summary>
        Tiempo_MinutosSegundos = 45,
        ///<summary>"[h]:mm:ss"</summary>
        Tiempo_HoraMinutosSegundos_Sumables = 46,
        ///<summary>"mmss.0"</summary>
        Tiempo_MinutosSegundos_Decimal = 47,
        ///<summary>"##0.0E+0"</summary>
        Cientifica_3Digitos = 48,
        ///<summary>"@"</summary>
        Arroba = 49,
        ///<summary>"#,#,,"</summary>
        Moneda_SinDecimales = 165,
    }
}
