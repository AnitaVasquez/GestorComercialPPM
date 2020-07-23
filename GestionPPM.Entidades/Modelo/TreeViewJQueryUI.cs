using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{
    public class TreeViewJQueryUI
    {
        public TreeViewJQueryUI()
        {
        }

        public TreeViewJQueryUI(string _text) {
            text = _text;
            children = new List<TreeViewJQueryUI>();
            state = "closed";
            iconCls = null;
            esCarpeta = true;
        }
        public Guid id { get; set; }
        public string text { get; set; }
        public string desc { get; set; }
        public string state { get; set; }
        public string iconCls { get; set; }
        public bool esCarpeta { get; set; }
        public List<TreeViewJQueryUI> children { get; set; } 
    }
    //public class Children {
    //    public Children() {
    //        children = new List<Children>();
    //    }
    //    public long id { get; set; }
    //    public string text { get; set; }
    //    public string state { get; set; }
    //    public List<Children> children { get; set; }
    //}
}