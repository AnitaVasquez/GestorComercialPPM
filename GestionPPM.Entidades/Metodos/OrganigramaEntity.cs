using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GestionPPM.Entidades.Metodos
{
    public class TreeOrganigrama
    {
        public int id { get; set; }
        public string usuario { get; set; }
        public string cargo { get; set; }
        public string mail { get; set; }
        public string departamento { get; set; }
        public string codigo { get; set; }
        public int? parent { get; set; }
        public List<TreeOrganigrama> Children { get; set; }
    }

    public class OrganigramaEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearOrganigrama(Organigrama organigrama)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    db.Organigrama.Add(organigrama);
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

        public static RespuestaTransaccion EditarOrganigrama(Organigrama organigrama)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // assume Entity base class have an Id property for all items
                    var entity = db.Organigrama.Find(organigrama.IDOrganigrama);
                    if (entity == null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }

                    db.Entry(entity).CurrentValues.SetValues(organigrama);
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

        public static RespuestaTransaccion EliminarOrganigrama(int id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var organigrama = db.Organigrama.Find(id);
                    organigrama.EstructuraOrganigrama = string.Empty;
                    db.Entry(organigrama).State = EntityState.Modified;
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

        public static bool EsOrganigramaNuevo(int? id = null, bool esTipoGenerico = false)
        {
            try
            {
                if (esTipoGenerico)
                    return true;

                var organigrama = db.Organigrama.FirstOrDefault(s => s.TipoOrganigramaID == id.Value);
                if (organigrama != null)
                    return false;
                else
                    return true; // Es uno nuevo
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static Organigrama ConsultarOrganigramaPrincipal(int? id = null)
        {
            Organigrama organigrama = id.HasValue ? new Organigrama() : new Organigrama { Estado = true, EmpresaID = 1, Codigo = "ORG-1" + Guid.NewGuid() };
            try
            {
                if (id.HasValue)
                    organigrama = db.Organigrama.FirstOrDefault(s => s.TipoOrganigramaID == id.Value);

                return organigrama;
            }
            catch (Exception ex)
            {
                return organigrama;
            }
        }

        public static TipoOrganigrama ConsultarTipoOrganigrama(int id)
        {
            TipoOrganigrama tipo = new TipoOrganigrama();
            try
            {
                tipo = db.TipoOrganigrama.FirstOrDefault(s => s.IDTipoOrganigrama == id);

                return tipo;
            }
            catch (Exception ex)
            {
                return tipo;
            }
        }

        public static string GetEstructuraOrganigrama(int id)
        {
            string organigrama = string.Empty;
            try
            {
                var elemento = db.Organigrama.FirstOrDefault(s=> s.TipoOrganigramaID == id);

                if (organigrama != null)
                    organigrama = elemento.EstructuraOrganigrama;

                return organigrama;
            }
            catch (Exception ex)
            {
                return organigrama;
            }
        }

        public static string GetOrganigrama()
        {
            try
            {
                var organigrama = db.Organigrama.FirstOrDefault();
                return organigrama.EstructuraOrganigrama;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }

        public static IEnumerable<SelectListItem> ObtenerListadoTiposOrganigrama(string seleccionado = null)
        {
            List<SelectListItem> listado = new List<SelectListItem>();
            try
            {
                listado = db.TipoOrganigrama.OrderBy(s => s.Nombre).Select(c => new SelectListItem
                {
                    Text = c.Nombre,
                    Value = c.IDTipoOrganigrama.ToString()
                }).ToList();

                if (!string.IsNullOrEmpty(seleccionado))
                {
                    if (listado.FirstOrDefault(s => s.Value == seleccionado.ToString()) != null)
                        listado.FirstOrDefault(s => s.Value == seleccionado.ToString()).Selected = true;
                }

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }

        }

        private static List<TreeOrganigrama> GetHijosOrganigramaInternoByUsuarioID(int id)
        {
            var org = ConsultarOrganigramaPrincipal(1).EstructuraOrganigrama;

            var estr = JsonConvert.DeserializeObject<List<TreeOrganigrama>>(org);

            var locations = estr.Where(x => x.parent == id || x.id == id).ToList();

            var child = locations.AsEnumerable().Union(
                                        estr.AsEnumerable().Where(x => x.parent == id).SelectMany(y => GetHijosOrganigramaInternoByUsuarioID(y.id))).ToList();
            return child;
        }

        private static List<TreeOrganigrama> GetHijosOrganigramaExternoByUsuarioID(int id)
        {
            var org = ConsultarOrganigramaPrincipal(2).EstructuraOrganigrama;

            var estr = JsonConvert.DeserializeObject<List<TreeOrganigrama>>(org);

            var locations = estr.Where(x => x.parent == id || x.id == id).ToList();

            var child = locations.AsEnumerable().Union(
                                        estr.AsEnumerable().Where(x => x.parent == id).SelectMany(y => GetHijosOrganigramaExternoByUsuarioID(y.id))).ToList();
            return child;
        }

        public static List<int> GetHijosOrganigramaByUsuarioID(int id, int organigramaID)
        {
            List<int> hijos = new List<int>();

            try
            {
                switch (organigramaID)
                {
                    case 1:
                        hijos = GetHijosOrganigramaInternoByUsuarioID(id).Where(s => s.id != id).Select(s => s.id).Distinct().ToList();
                        break;
                    case 2:
                        hijos = GetHijosOrganigramaExternoByUsuarioID(id).Where(s => s.id != id).Select(s => s.id).Distinct().ToList();
                        break;
                    default:
                        return hijos;
                }

                return hijos;
            }
            catch (Exception ex)
            {
                return hijos;
            }
        }
        public static NivelOrganigramaInfo ConsultarNivelOrganigrama(int idOrganigrama, int idEmpresa)
        {
            NivelOrganigramaInfo nivel = new NivelOrganigramaInfo();
            try
            {
                nivel = db.ConsultarNivelOrganigrama(idOrganigrama, idEmpresa).FirstOrDefault();

                return nivel;
            }
            catch (Exception ex)
            {
                return nivel;
            }
        }

        public static List<EstructuraOrganigramaInfo> ConsultarHijosOrganigramaPorUsuarioId(int idOrganigrama, int idEmpresa, int idUsuario, bool dependientes)
        {
            List<EstructuraOrganigramaInfo> hijos = new List<EstructuraOrganigramaInfo>();
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    hijos = context.ConsultarHijosOrganigramaPorUsuarioID(idOrganigrama, idEmpresa, idUsuario, dependientes).ToList();

                    return hijos;
                }
            }
            catch (Exception ex)
            {
                return hijos;
            }
        }

        public static List<EstructuraOrganigramaInfo> ConsultarEstructuraOrganigrama(int idOrganigrama, int idEmpresa, int idUsuario)
        {
            List<EstructuraOrganigramaInfo> estructura = new List<EstructuraOrganigramaInfo>();
            try
            {
                estructura = db.ConsultarEstructuraOrganigrama(idOrganigrama, idEmpresa, idUsuario).ToList();

                return estructura;
            }
            catch (Exception ex)
            {
                return estructura;
            }
        }

        public static int ConsultarNivelUsuario(int idOrganigrama, int idEmpresa, int idUsuario)
        {
            int nivel = 0;
            try
            {
                var usuario = db.ConsultarEstructuraOrganigrama(idOrganigrama, idEmpresa, idUsuario).FirstOrDefault();

                if (usuario != null)
                    nivel = usuario.nivel.Value;

                return nivel;

            }
            catch (Exception ex)
            {
                return nivel;
            }
        }

        public static List<EstructuraOrganigramaInfo> ConsultarEstructuraOrganigramaUsuarioIDByRangoNivel(int idOrganigrama, int idEmpresa, int idUsuario, int inicio, int fin)
        {
            List<EstructuraOrganigramaInfo> estructura = new List<EstructuraOrganigramaInfo>();
            try
            {
                estructura = db.ConsultarEstructuraOrganigramaByRangoNivel(idOrganigrama, idEmpresa, inicio, fin).ToList();

                return estructura;
            }
            catch (Exception ex)
            {
                return estructura;
            }
        }




    }
}