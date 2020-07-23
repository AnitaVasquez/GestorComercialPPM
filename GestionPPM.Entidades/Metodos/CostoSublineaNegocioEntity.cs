using GestionPPM.Entidades.Modelo;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace GestionPPM.Entidades.Metodos
{
    public class CostoSublineaNegocioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearCostoSublineaNegocio(CostoSublineaNegocio costoSublineaNegocio)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    costoSublineaNegocio.Estado = true;
                    db.CostoSublineaNegocio.Add(costoSublineaNegocio);
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

        public static RespuestaTransaccion EditarCostoSublineaNegocio(CostoSublineaNegocio costoSublineaNegocio)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    // assume Entity base class have an Id property for all items
                    var entity = db.CostoSublineaNegocio.Find(costoSublineaNegocio.IDCostoSublineaNegocio);
                    if (entity == null)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }

                    db.Entry(entity).CurrentValues.SetValues(costoSublineaNegocio);
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

        public static RespuestaTransaccion EliminarCostoSublineaNegocio(int id)
        {
            using (var transaction = db.Database.BeginTransaction())
            {
                try
                {
                    var costoSublineaNegocio = db.CostoSublineaNegocio.Find(id);
                    costoSublineaNegocio.Estado = false;
                    db.Entry(costoSublineaNegocio).State = EntityState.Modified;
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

        public static List<CostosSublineaNegocioInfo> ListadoCostosSublineaNegocio()
        {
            List<CostosSublineaNegocioInfo> listado = new List<CostosSublineaNegocioInfo>();
            try
            {
                listado = db.ListadoCostosSublineaNegocio().ToList();
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static List<SublineaNegocioCodigoCotizacionInfo> ConsultarCostosSublineaNegocioBySublineaNegocio(int tipoSolicitudID)
        {
            List<SublineaNegocioCodigoCotizacionInfo> listado = new List<SublineaNegocioCodigoCotizacionInfo>();
            try
            {
                var sublinea = db.ListadoCostosSublineaNegocio().FirstOrDefault(s => s.CodigoCatalogoSubTipoSolicitud == tipoSolicitudID);

                listado.Add(new SublineaNegocioCodigoCotizacionInfo
                {
                    //IdSublineaNegocioCotizacion = sublinea.CodigoCatalogoSublineaNegocio,
                    TextoCatalogoSublineaNegocio = sublinea.TextoCatalogoSublineaNegocio,
                    CodigoCatalogoSublineaNegocio = sublinea.CodigoCatalogoSublineaNegocio,
                    Valor = sublinea.Valor,
                    Estado = "Activo",
                });

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static CostosSublineaNegocioInfo ConsultarCostosSublineaNegocio(int id)
        {
            CostosSublineaNegocioInfo objeto = new CostosSublineaNegocioInfo();
            try
            {
                objeto = db.ConsultarCostosSublineaNegocio(id).FirstOrDefault();
                return objeto;
            }
            catch (Exception ex)
            {
                return objeto;
            }
        }

        public static bool SublineaNegocioExistente(int id, int catalogoSublineaNegocioID)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var objeto = context.CostoSublineaNegocio.FirstOrDefault(s => s.IDCostoSublineaNegocio != id && s.CatalogoSublineaNegocioID == catalogoSublineaNegocioID); //db.ConsultarCostosSublineaNegocio(id).FirstOrDefault(s => s.CodigoCatalogoSublineaNegocio == catalogoSublineaNegocioID);

                    if (objeto != null)
                        return true;
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                return true;
            }
        }

        public static bool TipoRequerimientoExistente(int id, int catalogoTipoID, int catalogoSubTipoID)
        {
            try
            {
                using (var context = new GestionPPMEntities())
                {
                    var objeto = context.CostoSublineaNegocio.FirstOrDefault(s => s.IDCostoSublineaNegocio != id && s.CatalogoTipoSolicitudID == catalogoTipoID && s.CatalogoSubTipoSolicitudID == catalogoSubTipoID && s.Estado); //db.ConsultarCostosSublineaNegocio(id).FirstOrDefault(s => s.CodigoCatalogoSublineaNegocio == catalogoSublineaNegocioID);

                    if (objeto != null)
                        return true;
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                return true;
            }
        }


    }
}