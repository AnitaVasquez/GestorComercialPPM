using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public static class ContactoClienteEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearContactosClientes(ContactosInfo Contactos, int? idCliente = null)
        {
            using (var transaction = db.Database.BeginTransaction())
            {

                try
                {
                    Contactos contacto = new Contactos
                    {
                        nombre_contacto = Contactos.nombre_contacto.ToUpper(),
                        apellido_contacto = Contactos.apellido_contacto.ToUpper(),
                        cargo_contacto = string.IsNullOrEmpty(Contactos.CodigoCatalogoCargoContacto) ? Contactos.cargo_contacto : Contactos.CodigoCatalogoCargoContacto,
                        mail_contacto = Contactos.mail_contacto,
                        telefono_contacto = Contactos.telefono_contacto,
                        extension_contacto = Contactos.extension_contacto,
                        prefijo_pais = Contactos.prefijo_pais,
                        celular_contacto = Contactos.celular_contacto,
                        tipo_contacto = Contactos.CodigoCatalogoTipoContacto,
                        estado_contacto = true,
                    };

                    db.Contactos.Add(contacto);
                    db.SaveChanges();

                    if (idCliente.HasValue)
                    {
                        ContactosClientes contactoClienteObj = new ContactosClientes
                        {
                            idCliente = idCliente.Value,
                            idContacto = contacto.id_contacto,
                        };
                        db.ContactosClientes.Add(contactoClienteObj);
                        db.SaveChanges();
                    }

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

        public static RespuestaTransaccion ActualizarContactosClientes(Contactos Contactos)
        {
            try
            {
                // Por si queda el Attach de la entidad y no deja actualizar
                var local = db.Contactos.FirstOrDefault(f => f.id_contacto == Contactos.id_contacto);
                if (local != null)
                {
                    db.Entry(local).State = EntityState.Detached;
                }

                Contactos.nombre_contacto = Contactos.nombre_contacto.ToUpper();
                Contactos.apellido_contacto = Contactos.apellido_contacto.ToUpper();
                db.Entry(Contactos).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarContactosClientes(int id)
        {
            try
            {
                var Contactos = db.Contactos.Find(id);

                if(Contactos.estado_contacto == true)
                {
                    Contactos.estado_contacto = false;
                }
                else
                { 
                    Contactos.estado_contacto = true;
                }

                db.Entry(Contactos).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CrearContactosCargaMasiva(List<Contactos> listado)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in listado)
                    {
                        item.estado_contacto = true;

                        db.Contactos.Add(item);
                        db.SaveChanges();
                    }

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


        public static List<ContactosInfo> ListarContactos()
        {
            try
            {
                return db.ListadoContactos().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ListadoContacto> ListadoContactosClientes()
        {
            try
            {
                return db.ListadoContacto().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ContactosClientesInfo> ListarContactosFacturacion()
        {
            try
            {
                return db.ListadoContactosClientes().Where(s=> s.TipoContacto == "TCFACT-01").ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<ContactosClientesInfo> ListarContactosDeCliente()
        {
            try
            {
                return db.ListadoContactosClientes().Where(s => s.TipoContacto == "TCCLI-01").ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }


        //Verificar si el contacto es de facturacion
        public static bool VerificarContactoFacturacion(int idContacto)
        {
            try
            {
                var contacto = db.ListadoContactosClientes().Where(s => s.TipoContacto == "TCFACT-01" && s.id_contacto == idContacto).FirstOrDefault();
                if (contacto == null)
                    return false;
                else
                    return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static Contactos ConsultarContactoCliente(int id)
        {
            try
            {
                Contactos contacto = db.Contactos.Find(id);
                return contacto;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<int> ListadIdsContactosClientesByCliente(int idCliente)
        {
            try
            {
                var listadoPerfiles = db.ContactosClientes.Where(s => s.idCliente == idCliente).Select(s => s.idContacto.Value).ToList();
                return listadoPerfiles;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<int> ListadIdsContactosCodigoCotizacion(int idCodigoCotizacion)
        {
            try
            {
                var listado = db.ContactosCodigoCotizacion.Where(s => s.idCodigoCotizacion == idCodigoCotizacion).Select(s => s.idContacto.Value).ToList();
                return listado;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool VerificarCorreoContactoExistente(string correo, int tipoContacto)
        {
            try
            {
                var validacion = db.Contactos.Where(s => s.mail_contacto == correo && s.estado_contacto.Value && s.tipo_contacto == tipoContacto).ToList();
                if (validacion.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public static bool VerificarCorreoContactoExistenteEdit(string correo, int tipoContacto, int idContacto)
        {
            try
            {
                var validacion = db.Contactos.Where(s => s.mail_contacto == correo && s.estado_contacto.Value && s.tipo_contacto == tipoContacto && s.id_contacto != idContacto ).ToList();
                if (validacion.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public static List<ContactosInfo> ListarContactosClientesByCliente(int idCliente)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {

                    var idsClientes = context.ContactosClientes.Where(s => s.idCliente == idCliente).Select(s => s.idContacto.Value).ToList();//ListadIdsContactosClientesByCliente(idCliente);

                    var listado = context.ListadoContactos().Where(s => idsClientes.Contains(s.id_contacto) && s.TipoContacto == "TCCLI-01").ToList();
                    return listado;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ListarContactosCliente(int? idCliente, int? seleccionado = null)
        {
            var listado = new List<SelectListItem>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var idsClientes = context.ContactosClientes.Where(s => s.idCliente == idCliente).Select(s => s.idContacto.Value).ToList();//ListadIdsContactosClientesByCliente(idCliente);

                    listado = context.ListadoContactos().Where(s => idsClientes.Contains(s.id_contacto) && s.TipoContacto == "TCCLI-01").OrderBy(s => s.nombre_contacto).Select(c => new SelectListItem
                    {
                        Text = c.nombre_contacto + " " + c.apellido_contacto,
                        Value = c.id_contacto.ToString()
                    }).ToList();

                    if (seleccionado.HasValue)
                    {
                        if (listado.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                            listado.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                    }

                    return listado;
                }

            }
            catch (Exception)
            {
                return listado;
            }
        }

    }
}