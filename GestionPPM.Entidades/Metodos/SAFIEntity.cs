using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web.Mvc;
using GestionPPM.Entidades.Modelo.SistemaContable;
using Newtonsoft.Json;
using NLog;
using RestSharp;
using System.Configuration;
using System.Net.Http;

namespace GestionPPM.Entidades.Metodos
{
    public class SAFIEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();


        public static List<ConsultarPresupuestoFacturados> ConsultarPresupuestosFaturados(int idEjecutivo, DateTime fechaInicial, DateTime fechaFinal)
        {
            List<ConsultarPresupuestoFacturados> listadoPresupuestosFacturados = new List<ConsultarPresupuestoFacturados>();
            try
            {
                listadoPresupuestosFacturados = db.ConsultarPresupuestoFacturados(idEjecutivo, fechaInicial, fechaFinal).ToList();
                return listadoPresupuestosFacturados;
            }
            catch (Exception ex)
            {
                return listadoPresupuestosFacturados;
            }
        }

        public static SAFIGeneral consultarDatosPresupuesto (int id)
        {
            SAFIGeneral safi = db.SAFIGeneral.Find(id);

            return safi;
        }

    }
}