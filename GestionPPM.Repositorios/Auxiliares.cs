//using NLog;
using OfficeToPDF;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Repositorios
{
    public static class Auxiliares
    {
        //private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public partial class MenuParcial
        {
            public int MenuID { get; set; }
            public int AccionID { get; set; }
            public string NombreAccionCatalogo { get; set; }
            public string Controlador { get; set; }
            public string Metodo { get; set; }
        }
        public partial class UsuarioRolMenuPermisoR
        {
            public int IDRolMenuPermiso { get; set; }
            public int RolID { get; set; }
            public string NombreRol { get; set; }
            public int PerfilID { get; set; }
            public string NombrePerfil { get; set; }
            public int MenuID { get; set; }
            public string NombreMenu { get; set; }
            public string EnlaceMenu { get; set; }
            public string MenuPadre { get; set; }
            public int IDCatalogo { get; set; }
            public string CodigoCatalogo { get; set; }
            public string TextoCatalogoAccion { get; set; }
            public int? CreadoPorID { get; set; }
            public string CreadoPor { get; set; }
            public int? ActualizadoPorID { get; set; }
            public string ActualizadoPor { get; set; }
            public DateTime? CreatedAt { get; set; }
            public DateTime? UpdatedAt { get; set; }
            public bool Estado { get; set; }
            public string MetodoControlador { get; set; }
            public string NombreControlador { get; set; }
            public string AccionEnlace { get; set; }
        }

        public static List<string> GetNombreCamposObjeto(List<object> collection)
        {
            var bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            List<string> listNames = collection.FirstOrDefault().GetType().GetFields(bindingFlags).Select(field => field.Name).ToList();

            return GetListaNombreColumnas(listNames).ConvertAll(d => d.ToUpper());
        }

        public static List<string> GetListaNombreColumnas(List<string> listado)
        {
            List<string> columnas = new List<string>();

            foreach (var item in listado)
            {
                int posicionInicial = item.IndexOf("<") + 1;
                int posicionFinal = item.IndexOf(">") - 1;
                string nombre = posicionInicial != -1 && posicionFinal != -1 ? item.Substring(posicionInicial, posicionFinal) : "";
                columnas.Add(nombre);
            }

            return columnas;
        }

        public static List<object> GetValoresCamposObjeto(object collection)
        {
            var bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            List<object> listValues = collection.GetType().GetFields(bindingFlags).Select(field => field.GetValue(collection)).Where(value => value != null).ToList();
            return listValues;
        }

        public static string FormatearValorMilesDecimales(decimal? valor)
        {
            try
            {
                string specifier;
                CultureInfo culture;

                specifier = "N6";
                culture = CultureInfo.CreateSpecificCulture("es-ES");

                var conversion = valor.Value.ToString(specifier, culture);
                return conversion;
            }
            catch (Exception ex)
            {
                return "0";
            }
        }

        public static string ConvertToListHtml(string prefacturas, string ids, string enlace)
        {
            string resultado = string.Empty;
            try
            {

                if (!string.IsNullOrEmpty(prefacturas) && !string.IsNullOrEmpty(ids))
                {
                    var itemsIDS = ids.Split(',');//listado.Split(',');
                    var listadoPrefacturas = prefacturas.Split(',');

                    int indice = 0;
                    foreach (string idPrefactura in itemsIDS)
                    {
                        string prefactura = listadoPrefacturas.ElementAt(indice);
                        string enlaceDescarga = enlace + "?listadoIDs=" + idPrefactura + "&descargaDirecta=" + true;
                        resultado += "<ul><li><a title='Descargar prefactura.' href='" + enlaceDescarga + "'> #" + prefactura + "</a></li></ul>";
                        //resultado += "<li> <a href='" + enlace + "'" + @service + "</a></li>";
                        indice++;
                    }

                    //<a href="index.html">valid</a>
                }

                return resultado;
            }
            catch (Exception ex)
            {
                return resultado;
            }

        }

        public static string ConvertToListHtmlRespuestas(string prefacturas, string ids, string enlace)
        {
            string resultado = string.Empty;
            try
            {

                if (!string.IsNullOrEmpty(prefacturas) && !string.IsNullOrEmpty(ids))
                {
                    var itemsIDS = ids.Split(',');//listado.Split(',');
                    var listadoPrefacturas = prefacturas.Split(',');

                    int indice = 0;
                    foreach (string idPrefactura in itemsIDS)
                    {
                        
                        string prefactura = listadoPrefacturas.ElementAt(indice);
                        string enlaceDescarga = enlace + "?listadoIDs=" + idPrefactura + "&descargaDirecta=" + true;
                        resultado += "  <li type='circle' style='color:"+ idPrefactura + "' >" + prefactura + "</li>";
                        //resultado += "<li> <a href='" + enlace + "'" + @service + "</a></li>";
                        indice++;
                    }

                    //<a href="index.html">valid</a>
                }

                return resultado;
            }
            catch (Exception ex)
            {
                return resultado;
            }
        }



        public static bool GetValorBooleano(string valor)
        {
            valor = valor.ToUpper();
            bool valorBooleano;
            switch (valor)
            {
                case "FALSE":
                    valorBooleano = false;
                    break;
                case "NO":
                    valorBooleano = false;
                    break;
                case "TRUE":
                    valorBooleano = true;
                    break;
                case "SI":
                    valorBooleano = true;
                    break;
                default:
                    valorBooleano = false;
                    break;
            }
            return valorBooleano;
        }

        public static string ConvertToLinkDescargaHtml(string id, string nombreArchivo, string enlace)
        {
            string resultado = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(nombreArchivo) && !string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(enlace))
                {
                    string enlaceDescarga = enlace + "?id=" + id;
                    resultado += "<ul><li><a title='Descargar Archivo.' href='" + enlaceDescarga + "'>" + nombreArchivo + "</a></li></ul>";
                }
                return resultado;
            }
            catch (Exception ex)
            {
                return resultado;
            }
        }


        public static IHtmlString GetCheckBoxTitle(string titulo, string idElemento, string claseElemento)
        {
            IHtmlString final = new HtmlString("");
            try
            {
                string elemento = "<input class='" + claseElemento + "' id='" + idElemento + "' name='CheckAll' title='Seleccionar todos los elementos mostrados en la página.' type='checkbox' value='false'>" + "<label>" + titulo + "</label>";
                final = new HtmlString(elemento);
                return final;
            }
            catch (Exception ex)
            {
                return final;
            }
        }

        //Para Font Awesome Icons
        public static string GetIconoExtension(string valor)
        {
            string icono;
            switch (valor)
            {
                case ".xls":
                    icono = "fa fa-file-excel-o";
                    break;
                case ".xlsx":
                    icono = "fa fa-file-excel-o";
                    break;
                case ".xlsm":
                    icono = "fa fa-file-excel-o";
                    break;
                case ".pdf":
                    icono = "fa fa-file-pdf-o";
                    break;
                case ".jpg":
                    icono = "fa fa-file-image-o";
                    break;
                case ".png":
                    icono = "fa fa-file-image-o";
                    break;
                case ".jpeg":
                    icono = "fa fa-file-image-o";
                    break;
                default:
                    icono = "";
                    break;
            }
            return icono;
        }

        public static string CrearCarpetasDirectorio(string carpetaInicial, List<string> carpetas)
        {
            string camino = carpetaInicial;

            foreach (string ele in carpetas)
            {
                if (!Directory.Exists(Path.Combine(camino, ele)))
                {
                    Directory.CreateDirectory(Path.Combine(camino, ele));
                }
                camino = Path.Combine(camino, ele);
            }

            return camino;

        }

        //Para ejecutar varios procesos en la línea de comandos
        public static void EjecutarProcesosCMD(List<string> comandos, List<string> archivos, string rutaExe = "")
        {
            try
            {
                Process cmd = new Process();
                cmd.StartInfo.FileName = rutaExe;// "cmd.exe";
                cmd.StartInfo.Arguments = archivos[0] + " " + archivos[1];
                cmd.StartInfo.RedirectStandardInput = true;

                cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();

                foreach (var item in comandos)
                {
                    cmd.StandardInput.WriteLine(item);
                }

                cmd.StandardInput.Flush();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                Console.WriteLine(cmd.StandardOutput.ReadToEnd());
            }
            catch (Exception ex)
            {
                //Log.Error(ex, "Excepcion");
            }
        }

        //Para objetos con propiedades en común
        public static T Casting<T>(this Object myobj)
        {
            Type objectType = myobj.GetType();
            Type target = typeof(T);
            var x = Activator.CreateInstance(target, false);
            var z = from source in objectType.GetMembers().ToList()
                    where source.MemberType == MemberTypes.Property
                    select source;
            var d = from source in target.GetMembers().ToList()
                    where source.MemberType == MemberTypes.Property
                    select source;
            List<MemberInfo> members = d.Where(memberInfo => d.Select(c => c.Name)
               .ToList().Contains(memberInfo.Name)).ToList();
            PropertyInfo propertyInfo;
            object value;
            foreach (var memberInfo in members)
            {
                propertyInfo = typeof(T).GetProperty(memberInfo.Name);
                value = myobj.GetType().GetProperty(memberInfo.Name).GetValue(myobj, null);

                propertyInfo.SetValue(x, value, null);
            }
            return (T)x;
        }


        // Add any bookmarks returned by the conversion process
        public static void AddPDFBookmarks(String generatedFile, List<PDFBookmark> bookmarks, Hashtable options, PdfOutline parent)
        {
            var hasParent = parent != null;
            if ((Boolean)options["verbose"])
            {
                Console.WriteLine("Adding {0} bookmarks {1}", bookmarks.Count, (hasParent ? "as a sub-bookmark" : "to the PDF"));
            }

            var srcPdf = hasParent ? parent.Owner : OpenPDFFile(generatedFile, options);
            if (srcPdf != null)
            {
                foreach (var bookmark in bookmarks)
                {
                    var page = srcPdf.Pages[bookmark.page - 1];
                    // Work out what level to add the bookmark
                    var outline = hasParent ? parent.Outlines.Add(bookmark.title, page) : srcPdf.Outlines.Add(bookmark.title, page);
                    if (bookmark.children != null && bookmark.children.Count > 0)
                    {
                        AddPDFBookmarks(generatedFile, bookmark.children, options, outline);
                    }
                }
                if (!hasParent)
                {
                    srcPdf.Save(generatedFile);
                }
            }
        }

        public static PdfDocument OpenPDFFile(string file, Hashtable options, PdfDocumentOpenMode mode = PdfDocumentOpenMode.Modify, string password = null)
        {
            int tries = 10;
            while (tries-- > 0)
            {
                try
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Opening {0} in PDF Reader", file);
                    }
                    if (password == null)
                    {
                        return PdfReader.Open(file, mode);
                    }
                    else
                    {
                        return PdfReader.Open(file, password, mode);
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Re-trying PDF open of {0}", file);
                    }
                    Thread.Sleep(500);
                }
            }
            return null;
        }

        // Perform some post-processing on the generated PDF
        public static void PostProcessPDFFile(String generatedFile, String finalFile, Hashtable options, Boolean postProcessPDFSecurity)
        {
            // Handle PDF merging
            if ((MergeMode)options["pdf_merge"] != MergeMode.None)
            {
                if ((Boolean)options["verbose"])
                {
                    Console.WriteLine("Merging with existing PDF");
                }
                PdfDocument srcDoc;
                PdfDocument dstDoc = null;
                if ((MergeMode)options["pdf_merge"] == MergeMode.Append)
                {
                    srcDoc = OpenPDFFile(generatedFile, options, PdfDocumentOpenMode.Import);
                    dstDoc = ReadExistingPDFDocument(finalFile, generatedFile, ((string)options["pdf_owner_pass"]).Trim(), PdfDocumentOpenMode.Modify, options);
                }
                else
                {
                    dstDoc = OpenPDFFile(generatedFile, options);
                    srcDoc = ReadExistingPDFDocument(finalFile, generatedFile, ((string)options["pdf_owner_pass"]).Trim(), PdfDocumentOpenMode.Import, options);
                }
                int pages = srcDoc.PageCount;
                for (int pi = 0; pi < pages; pi++)
                {
                    PdfPage page = srcDoc.Pages[pi];
                    dstDoc.AddPage(page);
                }
                dstDoc.Save(finalFile);
                System.IO.File.Delete(generatedFile);

            }

            if (options["pdf_page_mode"] != null || options["pdf_layout"] != null ||
                (MetaClean)options["pdf_clean_meta"] != MetaClean.None || postProcessPDFSecurity)
            {

                PdfDocument pdf = OpenPDFFile(finalFile, options);

                if (options["pdf_page_mode"] != null)
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Setting PDF Page mode");
                    }
                    pdf.PageMode = (PdfPageMode)options["pdf_page_mode"];
                }
                if (options["pdf_layout"] != null)
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Setting PDF layout");
                    }
                    pdf.PageLayout = (PdfPageLayout)options["pdf_layout"];
                }

                if ((MetaClean)options["pdf_clean_meta"] != MetaClean.None)
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Cleaning PDF meta-data");
                    }
                    pdf.Info.Creator = "";
                    pdf.Info.Keywords = "";
                    pdf.Info.Author = "";
                    pdf.Info.Subject = "";
                    //pdf.Info.Producer = "";
                    if ((MetaClean)options["pdf_clean_meta"] == MetaClean.Full)
                    {
                        pdf.Info.Title = "";
                        pdf.Info.CreationDate = System.DateTime.Today;
                        pdf.Info.ModificationDate = System.DateTime.Today;
                    }
                }

                // See if there are security changes needed
                if (postProcessPDFSecurity)
                {
                    PdfSecuritySettings secSettings = pdf.SecuritySettings;
                    if (((string)options["pdf_owner_pass"]).Trim().Length != 0)
                    {

                        // Set the owner password
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Setting PDF owner password");
                        }
                        secSettings.OwnerPassword = ((string)options["pdf_owner_pass"]).Trim();
                    }
                    if (((string)options["pdf_user_pass"]).Trim().Length != 0)
                    {
                        // Set the user password
                        // Set the owner password
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Setting PDF user password");
                        }
                        secSettings.UserPassword = ((string)options["pdf_user_pass"]).Trim();
                    }

                    secSettings.PermitAccessibilityExtractContent = !(Boolean)options["pdf_restrict_accessibility_extraction"];
                    secSettings.PermitAnnotations = !(Boolean)options["pdf_restrict_annotation"];
                    secSettings.PermitAssembleDocument = !(Boolean)options["pdf_restrict_assembly"];
                    secSettings.PermitExtractContent = !(Boolean)options["pdf_restrict_extraction"];
                    secSettings.PermitFormsFill = !(Boolean)options["pdf_restrict_forms"];
                    secSettings.PermitModifyDocument = !(Boolean)options["pdf_restrict_modify"];
                    secSettings.PermitPrint = !(Boolean)options["pdf_restrict_print"];
                    secSettings.PermitFullQualityPrint = !(Boolean)options["pdf_restrict_full_quality"];
                }
                pdf.Save(finalFile);
                pdf.Close();
            }
        }

        static PdfDocument ReadExistingPDFDocument(String filename, String generatedFilename, String password, PdfDocumentOpenMode mode, Hashtable options)
        {
            PdfDocument dstDoc = null;
            try
            {

                dstDoc = OpenPDFFile(filename, options, mode);
            }
            catch (PdfReaderException)
            {
                if (password.Length != 0)
                {
                    try
                    {
                        dstDoc = OpenPDFFile(filename, options, mode, password);
                    }
                    catch (PdfReaderException)
                    {
                        if (System.IO.File.Exists(generatedFilename))
                        {
                            System.IO.File.Delete(generatedFilename);
                        }
                        Console.WriteLine("Unable to modify a protected PDF. Invalid owner password given.");
                        // Return the general failure code and the specific failure code
                        Environment.Exit((int)ExitCode.PDFProtectedDocument);
                    }
                }
                else
                {
                    // There is no owner password, we can not open this
                    if (System.IO.File.Exists(generatedFilename))
                    {
                        System.IO.File.Delete(generatedFilename);
                    }
                    Console.WriteLine("Unable to modify a protected PDF. Use the /pdf_owner_pass option to specify an owner password.");
                    // Return the general failure code and the specific failure code
                    Environment.Exit((int)ExitCode.PDFProtectedDocument);
                }
            }
            return dstDoc;
        }

        public static Hashtable GetOptionsConversion()
        {
            Hashtable options = new Hashtable
            {
                // Loop through the input, grabbing switches off the command line
                ["hidden"] = false,
                ["markup"] = false,
                ["readonly"] = false,
                ["bookmarks"] = false,
                ["print"] = true,
                ["screen"] = false,
                ["pdfa"] = false,
                ["verbose"] = false,
                ["excludeprops"] = false,
                ["excludetags"] = false,
                ["noquit"] = false,
                ["merge"] = false,
                ["template"] = "",
                ["password"] = "",
                ["printer"] = "",
                ["fallback_printer"] = "",
                ["working_dir"] = "",
                ["has_working_dir"] = false,
                ["excel_show_formulas"] = false,
                ["excel_show_headings"] = false,
                ["excel_auto_macros"] = false,
                ["excel_template_macros"] = false,
                ["excel_active_sheet"] = false,
                ["excel_no_link_update"] = false,
                ["excel_no_recalculate"] = false,
                ["excel_max_rows"] = (int)0,
                ["excel_worksheet"] = (int)0,
                ["excel_delay"] = (int)0,
                ["word_field_quick_update"] = false,
                ["word_field_quick_update_safe"] = false,
                ["word_no_field_update"] = false,
                ["word_header_dist"] = (float)-1,
                ["word_footer_dist"] = (float)-1,
                ["word_max_pages"] = (int)0,
                ["word_ref_fonts"] = false,
                ["word_keep_history"] = false,
                ["word_no_repair"] = false,
                ["word_show_comments"] = false,
                ["word_show_revs_comments"] = false,
                ["word_show_format_changes"] = false,
                ["word_show_ink_annot"] = false,
                ["word_show_ins_del"] = false,
                ["word_markup_balloon"] = false,
                ["word_show_all_markup"] = false,
                ["word_fix_table_columns"] = false,
                ["original_filename"] = "",
                ["original_basename"] = "",
                ["powerpoint_output"] = "",
                ["pdf_page_mode"] = null,
                ["pdf_layout"] = null,
                ["pdf_merge"] = (int)MergeMode.None,
                ["pdf_clean_meta"] = (int)MetaClean.None,
                ["pdf_owner_pass"] = "",
                ["pdf_user_pass"] = "",
                ["pdf_restrict_annotation"] = false,
                ["pdf_restrict_extraction"] = false,
                ["pdf_restrict_assembly"] = false,
                ["pdf_restrict_forms"] = false,
                ["pdf_restrict_modify"] = false,
                ["pdf_restrict_print"] = false,
                ["pdf_restrict_annotation"] = false,
                ["pdf_restrict_accessibility_extraction"] = false,
                ["pdf_restrict_full_quality"] = false
            };
            return options;
        }

        public static IHtmlString GetBotones(List<UsuarioRolMenuPermisoR> listado, List<string> accionesControladorSeleccionado = null, string nombreControlador = null, string gridID = null)
        {
            var cadena = "";
            try
            {
                //if (listado.Any() && accionesControladorSeleccionado != null)
                if (accionesControladorSeleccionado != null)
                {
                    listado = listado.Where(s => accionesControladorSeleccionado.Contains(s.MetodoControlador)).ToList();

                    string textoPlaceHolderBusqueda = "Búsqueda";
                    string maximoCaracteresInputBusqueda = "50";
                    string functionalidadbusqueda = "<input class='filtro-general-busqueda-grid' type='text' maxlength='" + maximoCaracteresInputBusqueda + "' name='name' id='GridSearch' placeholder='" + textoPlaceHolderBusqueda + "' " + "style = 'width:200px; padding:8px 15px; margin:0px 0  display:inline-block; border:1px solid #ccc; background-color:#FBF6F3; border-radius:4px; background-color:#FBF6F3; border-radius:4px; box-sizing: border-box; color:black;' > ";

                    string textoToolTipBotonAgregar = "Agregar nuevo registro.";
                    string clsIconoAgregar = "fa fa-plus";
                    string colorBotonAgregar = "#00B5E6";
                    string functionalidadAgregarRegistro = "<button title='" + textoToolTipBotonAgregar + "' style='background-color:" + colorBotonAgregar + "; border-color: " + colorBotonAgregar + ";' class='btn btn-primary' id='nuevo'><i class='" + clsIconoAgregar + "' aria-hidden='true'></i></button>";

                    string textoToolTipBotonReporte = "Descargar información.";
                    string clsIconoReporte = "fa fa-download";
                    string colorBotonReporte = "#FFC52D";
                    string functionalidadDescargaReportes = "<button title='" + textoToolTipBotonReporte + "' style='background-color:" + colorBotonReporte + "; border-color: " + colorBotonReporte + ";' data-toggle='modal' data-target='#export-modal' class='btn btn-outline-info'  id='reporte'><i style='color: black;' class='" + clsIconoReporte + "' aria-hidden='true'></i></button>";

                    string textoToolTipBotonRecargarGrid = "Recargar información.";
                    string clsIconoRecargarGrid = "glyphicon glyphicon-refresh";
                    string colorBotonRecargarGrid = "#00AD8E";
                    string functionalidadRecargarGrid = "<button title='" + textoToolTipBotonRecargarGrid + "' style='background-color:" + colorBotonRecargarGrid + "; border-color: " + colorBotonRecargarGrid + ";' class='btn btn-success' id='recargar'><i class='" + clsIconoRecargarGrid + "' aria-hidden='true'></i></button>";

                    string textoToolTipBotonManual = "Manual Usuario.";
                    string clsIconoManual = "fa fa-question-circle";
                    string colorBotonManual = "#585858";
                    string functionalidadManualUsuario = "<button title='" + textoToolTipBotonManual + "' style='background-color:" + colorBotonManual + "; border-color: " + colorBotonManual + ";' data-toggle='modal' data-target='#help-modal' class='btn btn-outline-info' id='reporte'><i style='color:white;' class='" + clsIconoManual + "' aria-hidden='true'></i></button>";

                    string clsIconoCargaMasiva = "glyphicon glyphicon-open";
                    string colorBotonCargaMasiva = "#FE5000";
                    string textoToolTipBotonCargaMasiva = "Cargar información masiva.";
                    string funcionalidadCargaMasiva = "<button title='" + textoToolTipBotonCargaMasiva + "' style='background-color:" + colorBotonCargaMasiva + ";border-color: " + colorBotonCargaMasiva + ";' class='btn btn-danger' data-toggle='modal' data-backdrop='static' data-target='#cargas-masivas-modal' id='cargar-data'><span class='" + clsIconoCargaMasiva + "'></span> </button>";

                    string clsIconoAprobarPrefactura = "fa fa-thumbs-o-up";
                    string colorBotonAprobarPrefactura = "#ccc";
                    string textoToolTipBotonAprobarPrefactura = "Aprobar.";
                    string funcionalidadAprobarPrefactura = "<button title='" + textoToolTipBotonAprobarPrefactura + "' style='background-color:" + colorBotonAprobarPrefactura + "; border-color: " + colorBotonAprobarPrefactura + ";' class='btn btn-primary' id='aprobar-seleccion'><i style='color: black;' class='" + clsIconoAprobarPrefactura + "' aria-hidden='true'></i></button>";

                    string clsIconoConsolidar = "fa fa-check-square-o";
                    string colorBotonConsolidar = "#fff";
                    string textoToolTipBotonConsolidar = "Consolidar.";
                    string funcionalidadConsolidar = "<button title='" + textoToolTipBotonConsolidar + "' style='background-color:" + colorBotonConsolidar + "; border-color: " + colorBotonConsolidar + ";' class='btn btn-primary' id='consolidar-seleccion'><i style='color: black;'  class='" + clsIconoConsolidar + "' aria-hidden='true'></i></button>";

                    string clsIconoImpresionMasiva = "glyphicon glyphicon-print";
                    string colorImpresionMasiva = "#FE5000";
                    string textoToolTipBotonImpresionMasiva = "Impresión masiva.";
                    string funcionalidadImpresionMasiva = "<button title='" + textoToolTipBotonImpresionMasiva + "' style='background-color:" + colorImpresionMasiva + "; border-color: " + colorImpresionMasiva + ";' class='btn btn-primary' id='impresion-masiva'><i style='color: black;' class='" + clsIconoImpresionMasiva + "' aria-hidden='true'></i></button>";

                    string clsIconoEnvioSistemaContable = "glyphicon glyphicon-send";
                    string colorEnvioSistemaContable = "#00AD8E";
                    string textoToolTipEnvioSistemaContable = "Enviar al Sistema Contable.";
                    string funcionalidadEnvioSistemaContable = "<button title='" + textoToolTipEnvioSistemaContable + "' style='background-color:" + colorEnvioSistemaContable + "; border-color: " + colorEnvioSistemaContable + ";' class='btn btn-success' id='presupuestos-data'><i style = 'color: black;' class='" + clsIconoEnvioSistemaContable + "' aria-hidden='true'></i></button>";


                    if (nombreControlador != "ManejoPermisos")
                    {
                        foreach (var item in listado)
                        {
                            switch (item.CodigoCatalogo)
                            {
                                case "ACCIONES-SIST-01-BUSQUEDA":

                                    cadena += functionalidadbusqueda + " ";
                                    break;
                                case "ACCIONES-SIST-01-CREAR":
                                case "ACCIONES-SIST-01-CREAR-PARCIAL":
                                    if (gridID != "grid-Subcatalogo")
                                    {
                                        cadena += functionalidadAgregarRegistro + " ";
                                    }
                                    else
                                    {
                                        cadena += " " + " ";
                                    }
                                    break;
                                case "ACCIONES-SIST-01-CREAR-COSTO":
                                    cadena += functionalidadAgregarRegistro + " ";
                                    break;
                                case "ACCIONES-SIST-01-REPOR-BASICOS":
                                    if (gridID != "grid-Subcatalogo")
                                    {

                                        cadena += functionalidadDescargaReportes + " ";
                                    }
                                    else
                                    {
                                        cadena += " " + " ";
                                    }
                                    break;
                                case "ACCIONES-SIST-01-RECARGAR":
                                    cadena += functionalidadRecargarGrid + " ";
                                    break;
                                case "ACCIONES-SIST-01-MANUAL-US":
                                    cadena += functionalidadManualUsuario + " ";
                                    break;
                                case "ACCIONES-SIST-01-CARGAMASIVA":
                                    cadena += funcionalidadCargaMasiva + " ";
                                    break;
                                case "ACCIONES-SIST-01-APRO-PRE":
                                    cadena += funcionalidadAprobarPrefactura + " ";
                                    break;
                                case "ACCIONES-SIST-01-CONSOLIDAR":
                                    cadena += funcionalidadConsolidar + " ";
                                    break;
                                case "ACCIONES-SIST-01-IMP-MASIVA":
                                    if (nombreControlador != "SolicitudesClienteExterno" || nombreControlador != "Prefactura")
                                    {
                                        cadena += funcionalidadImpresionMasiva + " ";
                                    }
                                    cadena += " ";
                                    break;
                                case "ACCIONES-SIST-01-ENVIO-SIS-CONT":
                                    cadena += funcionalidadEnvioSistemaContable + " ";
                                    break;

                                default:
                                    cadena += String.Empty;
                                    break;
                            }
                        }
                    }
                    else
                    {
                        cadena = functionalidadbusqueda + " " + functionalidadDescargaReportes + " " + functionalidadRecargarGrid + " " + functionalidadManualUsuario;
                    }
                }

                return new HtmlString(cadena);
            }
            catch (Exception)
            {

                return new HtmlString(cadena);
            }

        }

        public static string GestionBontonesGrid(List<string> listado, string comando)
        {
            string clase = "ocultar-accion-catalogo";
            try
            {
                clase = listado.Any(s => s == comando) ? "mostrar-accion-catalogo" : "ocultar-accion-catalogo";
                return clase;
            }
            catch (Exception ex)
            {
                return clase;
            }
        }
        
        public static string VerificarPermisoAccionControlador(DataRow fila, int indiceColumna, List<SelectListItem> ListadoMenusAplicacion = null, List<SelectListItem> ListadoAccionesAplicacion = null)
        {
            string clase = "Error-Metodo-VerificarPermisoAccionControlador";
            try
            {
                string columnName = fila.Table.Columns[indiceColumna].ColumnName;

                switch (columnName)
                {
                    case "N":
                    case "NOMBREMENU":
                    case "MenuID":
                    case "PerfilID":
                    case "RolID":
                        clase = "Centrar";
                        break;
                    default:
                        bool valorColumna = Convert.ToBoolean(fila.ItemArray[indiceColumna]); // Valor del Check

                        var access = ListadoAccionesAplicacion.Where(s => columnName == s.Text.Split('|')[0]).FirstOrDefault();
                        var method = string.Empty;

                        if (access != null)
                            method = access.Text.Split('|')[1];

                        int menuID = int.Parse((fila["MenuID"] ?? "0").ToString());

                        var menu = ListadoMenusAplicacion.FirstOrDefault(s => s.Value == menuID.ToString());
                        string nombreControlador = menu != null ? menu.Text.Split('/')[0] : string.Empty;

                        var acciones = ActionNames(nombreControlador);

                        bool ok = acciones.Any(s => s == method);

                        if (!ok)
                            clase = "Centrar bloquear-accion-catalogo";
                        else
                            clase = "Centrar";
                        break;
                }
                return clase;
            }
            catch (Exception ex)
            {
                return clase;
            }
        }

        public static bool AccionPermitidaControlador(DataRow fila, int indiceColumna, List<SelectListItem> ListadoMenusAplicacion = null, List<SelectListItem> ListadoAccionesAplicacion = null)
        {
            bool permitido = false;
            try
            {
                string columnName = fila.Table.Columns[indiceColumna].ColumnName;

                switch (columnName)
                {
                    case "N":
                    case "NOMBREMENU":
                    case "MenuID":
                    case "PerfilID":
                    case "RolID":
                        permitido = true;
                        break;
                    default:
                        bool valorColumna = Convert.ToBoolean(fila.ItemArray[indiceColumna]); // Valor del Check

                        var access = ListadoAccionesAplicacion.Where(s => columnName == s.Text.Split('|')[0]).FirstOrDefault();
                        var method = string.Empty;

                        if (access != null)
                            method = access.Text.Split('|')[1];

                        int menuID = int.Parse((fila["MenuID"] ?? "0").ToString());

                        var menu = ListadoMenusAplicacion.FirstOrDefault(s => s.Value == menuID.ToString());
                        string nombreControlador = menu != null ? menu.Text.Split('/')[0] : string.Empty;

                        var acciones = ActionNames(nombreControlador);

                        bool ok = acciones.Any(s => s.Contains(method));

                        if (!ok)
                            permitido = false;
                        else
                            permitido = true;

                        break;
                }
                return permitido;
            }
            catch (Exception ex)
            {
                return permitido;
            }
        }

        public static string SetFilaTemporal(DataRow fila)
        {
            HttpContext.Current.Session["numeroColumna"] = fila;
            return "fila-almacenada";
        }

        public static DataRow GetFilaTemporal()
        {
            DataTable table = new DataTable();
            DataRow row = table.NewRow();
            try
            {
                row = HttpContext.Current.Session["Fila"] != null ? (DataRow)HttpContext.Current.Session["Fila"] : null;
                return row;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<string> ActionNames(string controllerName)
        {
            var types =
                from a in AppDomain.CurrentDomain.GetAssemblies()
                from t in a.GetTypes()
                where typeof(IController).IsAssignableFrom(t) &&
                    string.Equals(controllerName + "Controller", t.Name, StringComparison.OrdinalIgnoreCase)
                select t;

            var controllerType = types.FirstOrDefault();

            if (controllerType == null)
            {
                return Enumerable.Empty<string>().ToList();
            }
            return new ReflectedControllerDescriptor(controllerType)
               .GetCanonicalActions().Select(x => x.ActionName).ToList();
        }

        public static bool Base64Decode(string base64EncodedData, string rutaCompleta)
        {
            try
            {
                var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
                File.WriteAllBytes(rutaCompleta/*"pdf.pdf"*/, base64EncodedBytes);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
}