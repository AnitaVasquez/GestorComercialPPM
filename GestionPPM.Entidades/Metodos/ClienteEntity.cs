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
    public static class ClienteEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        //Metodo para listar clientes activos en lista desplegable
        public static IEnumerable<SelectListItem> ObtenerListadoClientes(int? id = null, bool filtroLineaNegocioGestor = true)
        {
            //Filtro de Clientes de la Línea de negocio del gestor
            var clientesGestor = db.ClienteLineaNegocio.Where(s => s.CatalogoLineaNegocioID == 710).Select(s => s.ClienteID).ToList();

            var listadoClientes = db.Cliente.Where(c => c.estado_cliente == true && clientesGestor.Contains(c.id_cliente)).OrderBy(c => c.nombre_comercial_cliente).Select(x => new SelectListItem
            {
                Text = x.nombre_comercial_cliente,
                Value = x.id_cliente.ToString()
            }).ToList();

            if (id.HasValue)
            {
                if (listadoClientes.FirstOrDefault(s => s.Value == id.Value.ToString()) != null)
                    listadoClientes.FirstOrDefault(s => s.Value == id.Value.ToString()).Selected = true;
            }

            return listadoClientes;
        }

        public static IEnumerable<SelectListItem> ObtenerListadoEjecutivos(int? idCliente = null, string seleccionado = null)
        {
            List<SelectListItem> ListadoCatalogo = new List<SelectListItem>();
            try
            {
                var items = ContactoClienteEntity.ListarContactosClientesByCliente(idCliente.Value);

                var listadoEjecutivos = items.OrderBy(c => c.nombre_contacto).Select(c => new SelectListItem
                {
                    Text = c.nombre_contacto + " " + c.apellido_contacto,
                    Value = c.id_contacto.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listadoEjecutivos.FirstOrDefault(s => s.Text == seleccionado.ToString()) != null)
                        listadoEjecutivos.FirstOrDefault(s => s.Text == seleccionado.ToString()).Selected = true;
                }

                return listadoEjecutivos;
            }
            catch (Exception ex)
            {
                return ListadoCatalogo;
            }

        }

        public static RespuestaTransaccion CrearCliente(Cliente Cliente, List<int> contactosCliente, List<int> lineasNegocioCliente)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    if (Cliente.razon_social_cliente != null)
                    {
                        Cliente.razon_social_cliente = Cliente.razon_social_cliente.ToUpper();
                    }
                    Cliente.nombre_comercial_cliente = Cliente.nombre_comercial_cliente.ToUpper();
                    Cliente.estado_cliente = true;

                    db.Cliente.Add(Cliente);
                    db.SaveChanges();

                    if (contactosCliente != null)
                    {
                        foreach (var item in contactosCliente)
                        {
                            db.ContactosClientes.Add(new ContactosClientes
                            {
                                idCliente = Cliente.id_cliente,
                                idContacto = item,
                            });
                            db.SaveChanges();
                        }
                    }

                    ClientesInfo cliente = db.ListadoClientes().FirstOrDefault(s => s.id_cliente == Cliente.id_cliente);
                    var prefijo = cliente.Prefijo;

                    if (contactosCliente != null)
                    {
                        var contactos = db.Contactos.Where(c => contactosCliente.Contains(c.id_contacto)).ToList();

                        foreach (var item in contactos)
                        {
                            item.prefijo_pais = prefijo;
                            db.Entry(item).State = EntityState.Modified;
                            db.SaveChanges();
                        }
                    }

                    if (lineasNegocioCliente != null)
                    {
                        //Limpiar los contactos anteriores
                        var ClienteLineaNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == Cliente.id_cliente).ToList();
                        foreach (var item in ClienteLineaNegocioLimpiar)
                        {
                            db.ClienteLineaNegocio.Remove(item);
                            db.SaveChanges();
                        }

                        foreach (var item in lineasNegocioCliente)
                        {
                            db.ClienteLineaNegocio.Add(new ClienteLineaNegocio
                            {
                                ClienteID = Cliente.id_cliente,
                                CatalogoLineaNegocioID = item,
                            });
                            db.SaveChanges();
                        }

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

        public static RespuestaTransaccion CrearActualizarClienteCargaMasiva(List<ClienteExcel> Clientes)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    foreach (var item in Clientes)
                    {
                        if (item.razon_social_cliente != null)
                        {
                            item.razon_social_cliente = item.razon_social_cliente.ToUpper();
                        }

                        var existeCliente = db.Cliente.Where(s => s.ruc_ci_cliente == item.ruc_ci_cliente).FirstOrDefault();

                        if (existeCliente != null) // Actualiza el cliente
                        {

                            var entidad = db.Cliente.Find(existeCliente.id_cliente);

                            entidad.nombre_comercial_cliente = item.nombre_comercial_cliente.ToUpper();
                            entidad.estado_cliente = true;

                            entidad.referido = item.referido;
                            entidad.intermediario = item.intermediario;
                            entidad.tipo_zoho = item.tipo_zoho;
                            entidad.tipo_cliente = item.tipo_cliente;
                            entidad.tamanio_empresa = item.tamanio_empresa;
                            entidad.etapa_cliente = item.etapa_cliente;
                            entidad.potencial_crecimiento = item.potencial_crecimiento;
                            entidad.categorizacion_cliente = item.categorizacion_cliente;
                            entidad.ruc_ci_cliente = item.ruc_ci_cliente;
                            entidad.razon_social_cliente = item.razon_social_cliente;
                            entidad.nombre_comercial_cliente = item.nombre_comercial_cliente;
                            entidad.ingresos_anuales_cliente = item.ingresos_anuales_cliente;
                            entidad.sector = item.sector;
                            entidad.pais = item.pais;
                            entidad.ciudad = item.ciudad;
                            entidad.direccion_cliente = item.direccion_cliente;
                            //entidad.estado_cliente = item.estado_cliente;
                            entidad.usuario_id = item.usuario_id;

                            db.Entry(entidad).State = EntityState.Modified;
                            db.SaveChanges();

                            var contactosCliente = item.contactos;
                            // Solo si se agregan contactos al cliente
                            if (contactosCliente != null)
                            {
                                //Limpiar los contactos anteriores
                                var contactosLimpiar = db.ContactosClientes.Where(s => s.idCliente == entidad.id_cliente).ToList();
                                foreach (var objLimpiar in contactosLimpiar)
                                {
                                    db.ContactosClientes.Remove(objLimpiar);
                                    db.SaveChanges();
                                }

                                foreach (var objNuevo in contactosCliente)
                                {
                                    var contactoNuevo = new ContactosClientes
                                    {
                                        idCliente = entidad.id_cliente,
                                        idContacto = objNuevo.idContacto,
                                    };

                                    db.ContactosClientes.Add(contactoNuevo);
                                    db.SaveChanges();

                                    ClientesInfo cliente = db.ListadoClientes().FirstOrDefault(s => s.id_cliente == entidad.id_cliente);
                                    var prefijo = cliente.Prefijo;

                                    var contactoEditarPrefijo = db.Contactos.Find(contactoNuevo.idContacto);

                                    if (contactoEditarPrefijo == null)
                                        return new RespuestaTransaccion { Estado = false, Respuesta = "No se pueden agregar contactos vacíos." };

                                    contactoEditarPrefijo.prefijo_pais = prefijo;

                                    db.Entry(contactoEditarPrefijo).State = EntityState.Modified;
                                    db.SaveChanges();

                                }
                            }
                            else
                            { // Se eliminaron todos los contactos del cliente
                              //Limpiar los contactos anteriores
                                var contactosLimpiar = db.ContactosClientes.Where(s => s.idCliente == entidad.id_cliente).ToList();
                                foreach (var objL in contactosLimpiar)
                                {
                                    db.ContactosClientes.Remove(objL);
                                    db.SaveChanges();
                                }
                            }


                            var lineasNegocio = item.lineasNegocio;
                            // Solo si se agregan LineaNegocio al cliente
                            if (lineasNegocio != null)
                            {
                                //Limpiar los ClienteLineaNegocio anteriores
                                var LineaNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == entidad.id_cliente).ToList();
                                foreach (var objLimpiar in LineaNegocioLimpiar)
                                {
                                    db.ClienteLineaNegocio.Remove(objLimpiar);
                                    db.SaveChanges();
                                }

                                foreach (var objNuevo in lineasNegocio)
                                {
                                    var lineasNegocioNuevo = new ClienteLineaNegocio
                                    {
                                        ClienteID = entidad.id_cliente,
                                        CatalogoLineaNegocioID = objNuevo.CatalogoLineaNegocioID,
                                    };

                                    db.ClienteLineaNegocio.Add(lineasNegocioNuevo);
                                    db.SaveChanges();
                                }
                            }
                            else
                            { // Se eliminaron todos los contactos del cliente
                              //Limpiar los contactos anteriores
                                var lineasNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == entidad.id_cliente).ToList();
                                foreach (var objL in lineasNegocioLimpiar)
                                {
                                    db.ClienteLineaNegocio.Remove(objL);
                                    db.SaveChanges();
                                }
                            }

                        }
                        else
                        { // Guarda el cliente

                            item.nombre_comercial_cliente = item.nombre_comercial_cliente.ToUpper();
                            item.estado_cliente = true;

                            var registro = new Cliente
                            {
                                nombre_comercial_cliente = item.nombre_comercial_cliente.ToUpper(),
                                estado_cliente = true,

                                referido = item.referido,
                                intermediario = item.intermediario,
                                tipo_zoho = item.tipo_zoho,
                                tipo_cliente = item.tipo_cliente,
                                tamanio_empresa = item.tamanio_empresa,
                                etapa_cliente = item.etapa_cliente,
                                potencial_crecimiento = item.potencial_crecimiento,
                                categorizacion_cliente = item.categorizacion_cliente,
                                ruc_ci_cliente = item.ruc_ci_cliente,
                                razon_social_cliente = item.razon_social_cliente,
                                ingresos_anuales_cliente = item.ingresos_anuales_cliente,
                                sector = item.sector,
                                pais = item.pais,
                                ciudad = item.ciudad,
                                direccion_cliente = item.direccion_cliente,
                                usuario_id = item.usuario_id,
                            };

                            db.Cliente.Add(registro);
                            db.SaveChanges();

                            var contactos = item.contactos;
                            if (contactos != null)
                            {
                                foreach (var obj in contactos)
                                {
                                    db.ContactosClientes.Add(new ContactosClientes
                                    {
                                        idCliente = registro.id_cliente,
                                        idContacto = obj.idContacto,
                                    });
                                    db.SaveChanges();
                                }
                            }

                            var lineasNegocio = item.lineasNegocio;
                            // Solo si se agregan LineaNegocio al cliente
                            if (lineasNegocio != null)
                            {
                                //Limpiar los ClienteLineaNegocio anteriores
                                var LineaNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == registro.id_cliente).ToList();
                                foreach (var objLimpiar in LineaNegocioLimpiar)
                                {
                                    db.ClienteLineaNegocio.Remove(objLimpiar);
                                    db.SaveChanges();
                                }

                                foreach (var objNuevo in lineasNegocio)
                                {
                                    var lineasNegocioNuevo = new ClienteLineaNegocio
                                    {
                                        ClienteID = registro.id_cliente,
                                        CatalogoLineaNegocioID = objNuevo.CatalogoLineaNegocioID,
                                    };

                                    db.ClienteLineaNegocio.Add(lineasNegocioNuevo);
                                    db.SaveChanges();
                                }
                            }


                        }
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

        public static RespuestaTransaccion ActualizarCliente(Cliente Cliente, List<int> contactosCliente, List<int> lineasNegocioCliente)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.Cliente.FirstOrDefault(f => f.id_cliente == Cliente.id_cliente);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    if (Cliente.razon_social_cliente != null)
                    {
                        Cliente.razon_social_cliente = Cliente.razon_social_cliente.ToUpper();
                    }

                    Cliente.nombre_comercial_cliente = Cliente.nombre_comercial_cliente.ToUpper();
                    db.Entry(Cliente).State = EntityState.Modified;
                    db.SaveChanges();

                    // Solo si se agregan contactos al cliente
                    if (contactosCliente != null)
                    {

                        //Limpiar los contactos anteriores
                        var contactosLimpiar = db.ContactosClientes.Where(s => s.idCliente == Cliente.id_cliente).ToList();
                        foreach (var item in contactosLimpiar)
                        {
                            db.ContactosClientes.Remove(item);
                            db.SaveChanges();
                        }

                        foreach (var item in contactosCliente)
                        {
                            db.ContactosClientes.Add(new ContactosClientes
                            {
                                idCliente = Cliente.id_cliente,
                                idContacto = item,
                            });
                            db.SaveChanges();
                        }

                        ClientesInfo cliente = db.ListadoClientes().FirstOrDefault(s => s.id_cliente == Cliente.id_cliente);
                        var prefijo = cliente.Prefijo;

                        var contactos = db.Contactos.Where(c => contactosCliente.Contains(c.id_contacto)).ToList();

                        foreach (var item in contactos)
                        {
                            item.prefijo_pais = prefijo;
                            db.Entry(item).State = EntityState.Modified;
                            db.SaveChanges();
                        }

                    }
                    else
                    { // Se eliminaron todos los contactos del cliente
                      //Limpiar los contactos anteriores
                        var contactosLimpiar = db.ContactosClientes.Where(s => s.idCliente == Cliente.id_cliente).ToList();
                        foreach (var item in contactosLimpiar)
                        {
                            db.ContactosClientes.Remove(item);
                            db.SaveChanges();
                        }
                    }

                    if (lineasNegocioCliente != null)
                    {
                        //Limpiar los contactos anteriores
                        var ClienteLineaNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == Cliente.id_cliente).ToList();
                        foreach (var item in ClienteLineaNegocioLimpiar)
                        {
                            db.ClienteLineaNegocio.Remove(item);
                            db.SaveChanges();
                        }

                        foreach (var item in lineasNegocioCliente)
                        {
                            db.ClienteLineaNegocio.Add(new ClienteLineaNegocio
                            {
                                ClienteID = Cliente.id_cliente,
                                CatalogoLineaNegocioID = item,
                            });
                            db.SaveChanges();
                        }

                    }
                    else
                    { // Se eliminaron todos los contactos del cliente
                      //Limpiar los contactos anteriores
                        var ClienteLineaNegocioLimpiar = db.ClienteLineaNegocio.Where(s => s.ClienteID == Cliente.id_cliente).ToList();
                        foreach (var item in ClienteLineaNegocioLimpiar)
                        {
                            db.ClienteLineaNegocio.Remove(item);
                            db.SaveChanges();
                        }
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

        //Eliminación Lógica
        public static RespuestaTransaccion EliminarCliente(int id)
        {
            try
            {
                var Cliente = db.Cliente.Find(id);

                if (Cliente.estado_cliente == true)
                {
                    Cliente.estado_cliente = false;
                }
                else
                {
                    Cliente.estado_cliente = true;
                }

                db.Entry(Cliente).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoCliente> ListarCliente()
        {
            try
            {
                return db.ListadoCliente().ToList();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool VerificarRUCCedulaExistente(string ruc, int? idCliente = null)
        {
            bool validacion = true;
            try
            {
                if (idCliente.HasValue)
                {
                    if (db.Cliente.Where(s => s.ruc_ci_cliente == ruc && s.id_cliente != idCliente).FirstOrDefault() != null)
                        validacion = false;
                }
                else
                {
                    if (db.Cliente.Where(s => s.ruc_ci_cliente == ruc).FirstOrDefault() != null)
                        validacion = false;
                }
                return validacion;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static Cliente ConsultarCliente(int id)
        {
            try
            {
                Cliente cliente = db.Cliente.Find(id);
                return cliente;
            }
            catch (Exception ex)
            {
                return new Cliente();
            }
        }

        public static ClientesInfo ConsultarClienteInformacionCompleta(int? id)
        {
            try
            {
                ClientesInfo cliente = db.ListadoClientes().FirstOrDefault(s => s.id_cliente == id);
                return cliente;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static List<long> ListarIdsLineasNegocioCliente(int idCliente)
        {
            try
            {
                var listado = db.ClienteLineaNegocio.Where(s => s.ClienteID == idCliente).Select(s => s.CatalogoLineaNegocioID).ToList();
                return listado;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static ListadoClienteLineaNegocio ConsultarClienteLineaNegocio(string linea, string ruc)
        {
            try
            {
                ListadoClienteLineaNegocio cliente = db.ListadoClienteLineaNegocio(linea).FirstOrDefault(s => s.ruc_ci_cliente == ruc);

                if (cliente != null)
                {
                    return cliente;
                }
                else
                {
                    ListadoClienteLineaNegocio clienteAuxiliar = new ListadoClienteLineaNegocio();
                    return cliente;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static IEnumerable<SelectListItem> ListarClientePorLineaNegocio(string LineaNegocio, string seleccionado = null)
        {
            List<SelectListItem> ListadoClienteLineaNegocio = new List<SelectListItem>();
            try
            {
                ListadoClienteLineaNegocio = db.ListadoClienteLineaNegocio(LineaNegocio).Select(c => new SelectListItem
                {
                    Text = c.ruc_ci_cliente + " - " + c.razon_social_cliente,
                    Value = c.id_cliente.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoClienteLineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoClienteLineaNegocio.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoClienteLineaNegocio;
            }
            catch (Exception ex)
            {
                return ListadoClienteLineaNegocio;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoCliente(string seleccionado = null)
        {
            List<SelectListItem> ListadoClientes = new List<SelectListItem>();
            try
            {
                ListadoClientes = db.ListadoCliente().Where(c => c.Estado == "Activo").OrderBy(c => c.Id).Select(c => new SelectListItem
                {
                    Text = c.Nombre_Comercial,
                    Value = c.Id.ToString(),
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (ListadoClientes.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        ListadoClientes.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return ListadoClientes;
            }
            catch (Exception ex)
            {
                return ListadoClientes;
            }

        }
    }
}