using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq; 
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ANSClienteEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearANSCliente(ANSCliente ansCliente)
        {
            try
            { 
                ansCliente.ans_estado = true;                
                db.ANSCliente.Add(ansCliente);
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarANSCliente(ANSCliente ansCliente)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.ANSCliente.FirstOrDefault(f => f.id_ans_sla == ansCliente.id_ans_sla);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                } 

                db.Entry(ansCliente).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarANSCliente(int id)
        {
            try
            {
                var ANSCliente = db.ANSCliente.Find(id);

                if(ANSCliente.ans_estado == true)
                {
                    ANSCliente.ans_estado = false;
                }
                else
                {
                    ANSCliente.ans_estado = true;
                }

                db.Entry(ANSCliente).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoANSCliente> ListarANSCliente()
        {
            try
            {
                return db.ListadoANSCliente().ToList();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
         
        public static ANSCliente ConsultarANSCliente(int id)
        {
            try
            {
                ANSCliente ansCliente = db.ANSCliente.Find(id);
                return ansCliente;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<SelectListItem> ListarANSdelCliente(string id, string id_tipo_solicitud, string seleccionado = null)
        {  
            List<SelectListItem> ListadoSLA = new List<SelectListItem>();
            try
            {
                ListadoSLA = db.ListadoANSDelCliente(Convert.ToInt32(id), Convert.ToInt32(id_tipo_solicitud)).Select(c => new SelectListItem
                {
                    Text = c.tipo_requerimiento,
                    Value = c.id_ans_sla.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoSLA.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoSLA.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoSLA;
            }
            catch (Exception ex)
            {
                return ListadoSLA;
            }
        }

        
    }
}