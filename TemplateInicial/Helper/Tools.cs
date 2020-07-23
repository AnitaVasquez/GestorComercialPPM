using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Collections; 
using System.IO;
using Newtonsoft.Json;
using NLog;

namespace TemplateInicial.Helper
{
    public static class Tools
    { 
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public static IEnumerable<SelectListItem> ConvertToSelectListItem<T>(this IEnumerable<T> elements, string selectedItem = "", string selectedText = null, string value = "Id", string name = "Nombre") where T : class
        {
            var list = new List<SelectListItem>();
            if (elements != null && elements.Any())
            {
                foreach (var item in elements)
                {
                    var type = item.GetType();
                    var valueProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(value.ToLower()));
                    var nameProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(name.ToLower()));
                    if (valueProp != null && nameProp != null)
                    {
                        var valueV = valueProp.GetValue(item, null);
                        var nameV = nameProp.GetValue(item, null);
                        var temp = new SelectListItem()
                        {
                            Value = valueV.ToString().Trim(),
                            Text = nameV.ToString().Trim(),
                            Selected = valueV.ToString().Equals(selectedItem)
                        };
                        list.Add(temp);
                    }
                }
                if (!string.IsNullOrEmpty(selectedText))
                {
                    list.Insert(0, new SelectListItem()
                    {
                        Selected = false,
                        Text = selectedText,
                        Value = ""
                    });
                }
            }
            else
            {
                list.Insert(0, new SelectListItem()
                {
                    Selected = true,
                    Text = "No Existen...",
                    Value = "-1"
                });
            }

            return list;
        }

        public static bool IsInRange(this double number, double edadMaxima, double edadMinima)
        {

            var text = number.ToString(CultureInfo.InvariantCulture);
            var array = text.Split(",.".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            //if (array.Length > 1)
            //    return int.Parse(array[1]);
            //return int.Parse(array[0]);
            return edadMinima <= number && number <= edadMaxima;
        }

        public static IEnumerable<SelectListItem> ConvertToSelectListItem<T>(this IEnumerable<T> elements, IEnumerable<T> validos, string selectedItem = "", string selectedText = null, string value = "Id", string name = "Nombre") where T : class
        {
            var list = new List<SelectListItem>();
            var invalids = elements.Except(validos);
            foreach (var item in validos)
            {
                var type = item.GetType();
                var valueProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(value.ToLower()));
                var nameProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(name.ToLower()));
                if (valueProp != null && nameProp != null)
                {
                    var valueV = valueProp.GetValue(item, null);
                    var nameV = nameProp.GetValue(item, null);
                    var temp = new SelectListItem()
                    {
                        Value = valueV.ToString().Trim(),
                        Text = nameV.ToString().Trim(),
                        Selected = valueV.ToString().Equals(selectedItem)
                    };
                    list.Add(temp);
                }
            }
            foreach (var item in invalids)
            {
                var type = item.GetType();
                var valueProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(value.ToLower()));
                var nameProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(name.ToLower()));
                if (valueProp != null && nameProp != null)
                {
                    var valueV = valueProp.GetValue(item, null);
                    var nameV = nameProp.GetValue(item, null);
                    var temp = new SelectListItem()
                    {
                        Value = valueV.ToString(),
                        Text = "☺" + nameV.ToString(),
                        Selected = valueV.ToString().Equals(selectedItem)
                    };
                    list.Add(temp);
                }
            }
            if (!string.IsNullOrEmpty(selectedText))
            {
                list.Insert(0, new SelectListItem()
                {
                    Selected = false,
                    Text = selectedText,
                    Value = ""
                });
            }
            return list;
        }

        public static IEnumerable<SelectListItem> ConvertToSelectListItem<T>(this IEnumerable<T> elements, IEnumerable<string> selectedItems, string selectedText = null, string value = "Id", string name = "Nombre") where T : class
        {
            var list = new List<SelectListItem>();
            foreach (var item in elements)
            {
                var type = item.GetType();
                var valueProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(value.ToLower()));
                var nameProp = type.GetProperties().FirstOrDefault(t => t.Name.ToLower().Equals(name.ToLower()));
                if (valueProp != null && nameProp != null)
                {
                    var valueV = valueProp.GetValue(item, null);
                    var nameV = nameProp.GetValue(item, null);
                    var temp = new SelectListItem()
                    {
                        Value = valueV.ToString().Trim(),
                        Text = nameV.ToString().Trim(),
                        Selected = selectedItems.Contains(valueV.ToString())
                    };
                    list.Add(temp);
                }
            }
            if (selectedText != null)
            {
                list.Insert(0, new SelectListItem()
                {
                    Selected = false,
                    Text = selectedText,
                    Value = ""
                });
            }
            return list;
        }

        public static long GetIdFromPath(this string path)
        {
            var regex = new Regex("[0-9]+");
            var collection = regex.Matches(path);
            if (collection.Count > 0)
                return long.Parse(collection[collection.Count - 1].Value);
            return 0;
        }

        public static string NormalizeForSearch(this string text)
        {
            var temp =
                text.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u").Replace(
                    " ", "");
            return temp;
        }

        /// <summary>
        ///  A partir de un ModelState de una action POST, crea una lista resumida de los errores que existen
        /// </summary>
        /// <param name="actual">ModelState actual</param>
        /// <returns></returns>
        public static List<string> ErroresValidacionModelState(ModelStateDictionary actual)
        {
            return (from valor in actual.Values
                    where valor.Errors.Any()
                    from error in valor.Errors
                    select error.ErrorMessage).ToList();
        }

        public static string CrearCaminos(string carpetaInicial, List<string> carpetas)
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

        public static string GetContentType(string extension)
        {
            var contentType = "";

            switch (extension.ToLower())
            {
                case ".pdf":
                    contentType = "application/pdf";
                    break;
                case ".doc":
                    contentType = "application/doc";
                    break;
                case ".docx":
                    contentType = "application/doc";
                    break;
                case ".txt":
                    contentType = "application/txt";
                    break;
                case ".xls":
                    contentType = "application/xls";
                    //Response.ContentType = "application/ms-excel";
                    break;
                case ".xlsx":
                    contentType = "application/xls";
                    //Response.ContentType = "application/ms-excel";
                    break;
                case ".log":
                    contentType = "application/txt";
                    break;
                case ".rar":
                    contentType = "application/rar";
                    break;
                case ".7z":
                    contentType = "application/7z";
                    break;
                case ".jpg":
                    contentType = "application/jpg";
                    break;
                case ".bmp":
                    contentType = "application/bmp";
                    break;
                case ".png":
                    contentType = "application/png";
                    break;
            }

            return contentType;
        }

        public static string GetNameMonth(int mont)
        {
            var res = "";

            switch (mont)
            {
                case 1:
                    res = "Enero";
                    break;
                case 2:
                    res = "Febrero";
                    break;
                case 3:
                    res = "Marzo";
                    break;
                case 4:
                    res = "Abril";
                    break;
                case 5:
                    res = "Mayo";
                    break;
                case 6:
                    res = "Junio";
                    break;
                case 7:
                    res = "Julio";
                    break;
                case 8:
                    res = "Agosto";
                    break;
                case 9:
                    res = "Septiembre";
                    break;
                case 10:
                    res = "Octubre";
                    break;
                case 11:
                    res = "Noviembre";
                    break;
                case 12:
                    res = "Diciembre";
                    break;
            }
            return res;
        }
        public static string GetNameDay(int day)
        {
            var res = "";

            switch (day)
            {
                case 1:
                    res = "Lunes";
                    break;
                case 2:
                    res = "Martes";
                    break;
                case 3:
                    res = "Miércoles";
                    break;
                case 4:
                    res = "Jueves";
                    break;
                case 5:
                    res = "Viernes";
                    break;
                case 6:
                    res = "Sábado";
                    break;
                case 7:
                    res = "Domingo";
                    break;
            }
            return res;
        }

        public static string GetNombreArchivoPlantilla(string plantilla)
        {
            var archivo = string.Empty;
            try
            {
                switch (plantilla)
                {
                    case "PREFACTURA":
                        archivo = "prefactura.pdf";
                        break;
                    default:
                        archivo = string.Empty;
                        break;
                }
                return archivo;
            }
            catch (Exception ex)
            {
                return archivo;
            }
        }
    }
}