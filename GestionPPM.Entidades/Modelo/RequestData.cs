using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Modelo
{ 
    public class RequestData
    {
        public string Version { get; set; }
        public int StatusCode { get; set; }
        public string Data { get; set; }
        public List<MessagesData> Messages { get; set; }
    }

    public class MessagesData
    {
        public string ID { get; set; }
        public string Message { get; set; }
        public string InfoAdditional { get; set; }
        public string Type { get; set; }
    }
     
}