using GestionPPM.Entidades.Modelo.PlaceToPay;
using GestionPPM.Repositorios;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace GestionPPM.Entidades.Metodos
{
    public class ComercioPlaceToPayEntity
    {
        private static readonly PlaceToPay db = new PlaceToPay();

        public static async Task<RespuestaTransaccion> CrearComercioPlaceToPay(ComercioPlaceToPay objeto)
        {
            db.ComercioPlaceToPay.Add(objeto);
            try
            {
                await db.SaveChangesAsync();
                return new RespuestaTransaccion
                {
                    Estado = true,
                    Respuesta = Mensajes.MensajeTransaccionExitosa
                };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion
                {
                    Estado = false,
                    Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() + " ;" + ex.InnerException.Message
                };
            }
        }

        public static async Task<RespuestaTransaccion> ActualizarComercioPlaceToPay(ComercioPlaceToPay objeto)
        {
            try
            {
                // assume Entity base class have an Id property for all items
                var entity = db.ComercioPlaceToPay.Find(objeto.IDComercioPlaceToPay);

                if (entity == null)
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };

                db.Entry(entity).CurrentValues.SetValues(objeto);

                return await db.SaveChangesAsync() > 0 ? new RespuestaTransaccion
                {
                    Estado = true,
                    Respuesta = Mensajes.MensajeTransaccionExitosa
                } : new RespuestaTransaccion
                {
                    Estado = true,
                    Respuesta = Mensajes.MensajeTransaccionFallida
                };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion
                {
                    Estado = false,
                    Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString()
                };
            }
        }

        public static async Task<List<ComercioPlaceToPayInfo>> ListadoComercioPlaceToPayAsync(long? pagina = null, string textoBusqueda = null, string filtro = null, int? id = null)
        {
            List<ComercioPlaceToPayInfo> listado = new List<ComercioPlaceToPayInfo>();
            try
            {
                if (!id.HasValue)
                    listado = await db.ListadoComercioPlaceToPay(pagina, textoBusqueda, filtro).AsQueryable().ToListAsync(); // Listado Completo
                else
                {
                    filtro = " WHERE IDComercioPlaceToPay = '{0}' ";

                    filtro = id.HasValue ? string.Format(filtro, id) : null;
                    listado = await db.ListadoComercioPlaceToPay(null, null, filtro).AsQueryable().ToListAsync(); // Consulta por ID
                }

                return listado;
            }
            catch (Exception)
            {
                return listado;
            }
        }

        public static  List<ComercioPlaceToPayInfo> ListadoComercioPlaceToPay(long? pagina = null, string textoBusqueda = null, string filtro = null, int? id = null)
        {
            List<ComercioPlaceToPayInfo> listado = new List<ComercioPlaceToPayInfo>();
            try
            {
                if (!id.HasValue)
                    listado =  db.ListadoComercioPlaceToPay(pagina, textoBusqueda, filtro).ToList(); // Listado Completo
                else
                {
                    filtro = " WHERE IDComercioPlaceToPay = '{0}' ";

                    filtro = id.HasValue ? string.Format(filtro, id) : null;
                    listado =  db.ListadoComercioPlaceToPay(null, null, filtro).ToList(); // Consulta por ID
                }

                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }


        public static async Task<ComercioPlaceToPay> GetComercioPlaceToPayAsync(int id)
        {
            try
            {
                var entity = await db.ComercioPlaceToPay.SingleOrDefaultAsync(s => s.IDComercioPlaceToPay == id);
                return entity;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static async Task<ComercioPlaceToPay> GetInformacionComercioPlaceToPayAsync(string ruc)
        {
            try
            {
                var entity = await db.ComercioPlaceToPay.SingleOrDefaultAsync(s => s.RUC.Contains(ruc));
                return entity;
            }
            catch (Exception ex)
            {
                return new ComercioPlaceToPay();
            }
        }

        public static List<ComercioPlaceToPayInfo> ConsultarComercioPlaceToPayPorRUC( string ruc)
        {
            List<ComercioPlaceToPayInfo> listado = new List<ComercioPlaceToPayInfo>();
            try
            {
                string filtro = " WHERE RUC LIKE '%{0}%' ";

                filtro = string.Format(filtro, ruc);
                listado = db.ListadoComercioPlaceToPay(null, null, filtro).ToList(); // Consulta por ID
                return listado;
            }
            catch (Exception ex)
            {
                return listado;
            }
        }

        public static int ObtenerTotalRegistrosListadoComercioPlaceToPay()
        {
            int total = 0;
            try
            {
                total = db.Database.SqlQuery<int>("SELECT [dbo].[ObtenerTotalRegistrosListadoComercioPlaceToPay]()").Single();
                return total;
            }
            catch (Exception ex)
            {
                return total;
            }
        }
    }
}