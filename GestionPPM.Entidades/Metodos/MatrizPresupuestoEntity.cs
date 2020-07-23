using GestionPPM.Entidades.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public class MatrizPresupuestoEntity
    {
        private static GestionPPMEntities db = new GestionPPMEntities();

        public static List<MatrizPresupuestoInfo> ConsultarMatrizPresupuestoByCodigoCotizacion(int idCodigoCotizacion, DateTime fechaInicial, DateTime fechaFinal)
        {
            List<MatrizPresupuestoInfo> listado = new List<MatrizPresupuestoInfo>();
            try
            {
                listado = db.ConsultarMatrizPresupuesto(idCodigoCotizacion, fechaInicial, fechaFinal).ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }
    }
}