using NonFactors.Mvc.Grid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TemplateInicial.Models
{
    public static class Tools
    {
        public static IGridColumn<T, TValue> RawNamed<T, TValue>(this IGridColumn<T, TValue> column, String name)
        {
            column.Name = name;
            return column;
        }
    }
}