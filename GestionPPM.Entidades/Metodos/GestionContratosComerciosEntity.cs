using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using static System.Net.Mime.MediaTypeNames;

namespace GestionPPM.Entidades.Metodos
{
    public static class GestionContratosComerciosEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearGestionContratosComercios(GestionContratosComercios gestion)
        {

            //manejar transaccionalidad
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {

                    db.GestionContratosComercios.Add(gestion);
                    db.SaveChanges();

                    Comercio comercio = db.Comercio.Find(gestion.id_comercio);
                    comercio.id_estatus_contrato = gestion.id_estatus_contrato_actual;

                    db.Entry(comercio).State = EntityState.Modified;
                    db.SaveChanges();

                    transaction.Commit();
                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };

                }
            }
        }
    }
}