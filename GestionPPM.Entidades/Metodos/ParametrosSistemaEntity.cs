using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ParametrosSistemaEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearParametrosSistema(ParametrosSistema parametros)
        {
            try
            {
                parametros.nombre = parametros.nombre.ToUpper();
                parametros.estado = true;                
                db.ParametrosSistema.Add(parametros);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarParametrosSistema(ParametrosSistema parametros)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.ParametrosSistema.FirstOrDefault(f => f.id_parametro == parametros.id_parametro);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                parametros.nombre = parametros.nombre.ToUpper();
                db.Entry(parametros).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarParametrosSistema(int id)
        {
            try
            {
                var parametros = db.ParametrosSistema.Find(id);

                if(parametros.estado ==true)
                {
                    parametros.estado = false;
                }
                else
                {
                    parametros.estado = true;
                }
                 
                db.Entry(parametros).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoParametrosSistemas> ListarParametros()
        {
            try
            {
                return db.ListadoParametrosSistemas().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
         
        public static ParametrosSistema ConsultarParametros(int id)
        {
            try
            {
                ParametrosSistema parametros = db.ParametrosSistema.Find(id);
                return parametros;
            }
            catch (Exception)
            {
                throw;
            }
        } 

    }
}