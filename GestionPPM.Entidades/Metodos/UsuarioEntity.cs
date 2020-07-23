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
    public static class UsuarioEntity
    {
        private static readonly GestionPPMEntities db = new GestionPPMEntities();

        public static RespuestaTransaccion CrearUsuario(UsuarioCE usuario) 
        {
            bool todoOk = true;
            try
            {
                //validar si existe el registro por el correo
                var validarUsuarios = db.Usuario.Where(u => u.mail_usuario == usuario.mail_usuario).ToList();

                if (validarUsuarios.Count() > 0)
                {
                    todoOk = false;
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                }
                else
                {
                    //validar si existe el registro de tipo interno con ese codigo vendedor
                    var validarUsuariosCodigoVendeor = db.Usuario.Where(u => u.secu_usua == usuario.secu_usua && u.tipo_usuario ==109).ToList();
                    if (validarUsuariosCodigoVendeor.Count() > 0)
                    {
                        todoOk = false;
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCodigoVendedor };
                    }
                    else
                    {
                        //validar que todos los campos esten completos
                        if (usuario.nombre_usuario == null || usuario.apellido_usuario == null || usuario.pais == null || usuario.ciudad == null ||
                        usuario.direccion_usuario == null || usuario.mail_usuario == null || usuario.telefono_usuario == null || usuario.id_rol == 0)
                        {
                            todoOk = false;
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                        }
                        else
                        {
                            //Validar Campos por Usuario Interno
                            if (usuario.tipo_usuario == 109)
                            {
                                if (usuario.area_departamento == null || usuario.cargo_usuario == null || usuario.secu_usua == 0 || usuario.link_firma == null)
                                {
                                    todoOk = false;
                                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                                } 
                            }
                            else
                            {
                                if (usuario.area_departamento_texto == null || usuario.cargo_usuario_texto == null || usuario.tipo_usuario == null || usuario.cliente_asociado == null)
                                {
                                    todoOk = false;
                                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                                }
                            }
                        }
                    }
                }

                //Si no hay eerores se continua con la creacion del usuario
                if (todoOk == true)
                {
                    //validar el correo 
                    if (Validaciones.ValidarMail(usuario.mail_usuario) == false)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto };
                    }
                    else
                    {
                        //manejar transaccionalidad
                        using (var transaction = db.Database.BeginTransaction())
                        {
                            try
                            {
                                //Crear el usuario Final
                                Usuario usuarioFinal = new Usuario();

                                usuarioFinal.nombre_usuario = usuario.nombre_usuario.ToUpper();
                                usuarioFinal.apellido_usuario = usuario.apellido_usuario.ToUpper();
                                usuarioFinal.tipo_usuario = usuario.tipo_usuario;
                                usuarioFinal.cliente_asociado = usuario.cliente_asociado;
                                usuarioFinal.pais = usuario.pais;
                                usuarioFinal.ciudad = usuario.ciudad;
                                usuarioFinal.direccion_usuario = usuario.direccion_usuario;
                                usuarioFinal.mail_usuario = usuario.mail_usuario;
                                usuarioFinal.telefono_usuario = usuario.telefono_usuario;
                                usuarioFinal.celular_usuario = usuario.celular_usuario;
                                usuarioFinal.rol_id = usuario.id_rol;

                                //validacion tipo usuario
                                //usuario interno
                                if (usuario.tipo_usuario == 109)
                                {
                                    usuarioFinal.area_departamento = usuario.area_departamento;
                                    usuarioFinal.cargo_usuario = usuario.cargo_usuario;
                                    usuarioFinal.secu_usua = usuario.secu_usua;
                                    usuarioFinal.link_firma = usuario.link_firma;
                                    usuarioFinal.validacion_correo = usuario.validacion_correo;
                                    usuarioFinal.cliente_asociado = null;
                                    usuarioFinal.reset_clave = false;
                                }
                                //usuario externo
                                else
                                {
                                    usuarioFinal.area_departamento = usuario.area_departamento_texto.ToUpper();
                                    usuarioFinal.cargo_usuario = usuario.cargo_usuario_texto.ToUpper();
                                    usuarioFinal.reset_clave = false;
                                    usuarioFinal.validacion_correo = false;
                                }

                                //generar la clave random de 8
                                Random random = new Random();
                                String numero = "";
                                for (int i = 0; i < 9; i++)
                                {
                                    numero += Convert.ToString(random.Next(0, 9));
                                }

                                //encriptar la clave 
                                var clave = Validaciones.GetMD5_1(numero);

                                usuarioFinal.clave_usuario = clave;
                                usuarioFinal.estado_usuario = true;
                                usuarioFinal.activo_usuario = false;
                                usuarioFinal.id_empresa = 1;

                                //obtener codigo usuario
                                var codigo = db.usp_g_codigo_usuario(usuarioFinal.tipo_usuario).First().codigoUsuario;
                                usuarioFinal.codigo_usuario = codigo.ToString();

                                db.Usuario.Add(usuarioFinal);
                                db.SaveChanges();
                                int id_usuario = usuarioFinal.id_usuario;

                                //enviar correo
                                db.usp_envia_correo_usuario(id_usuario, numero);
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
                else
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion ActualizarUsuario(UsuarioCE usuario)
        {
            bool todoOk = true;
            try
            {
                //validar que todos los campos esten completos
                if (usuario.nombre_usuario == null || usuario.apellido_usuario == null || usuario.pais == null || usuario.ciudad == null ||
                usuario.direccion_usuario == null || usuario.mail_usuario == null || usuario.telefono_usuario == null ||
                usuario.id_rol == 0)
                {
                    todoOk = false;
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                }
                else
                {
                    //Validar Campos por Usuario Interno
                    if (usuario.tipo_usuario == 109)
                    {
                        if (usuario.area_departamento == null || usuario.cargo_usuario == null || usuario.secu_usua == 0 || usuario.link_firma == null)
                        {
                            todoOk = false;
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                        }
                    }
                    else
                    {
                        if (usuario.area_departamento_texto == null || usuario.cargo_usuario_texto == null || usuario.cliente_asociado == null)
                        {
                            todoOk = false;
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                        }
                    }

                    //si todos los campos estan llenos se procede a validar datos

                    if (todoOk == true)
                    {
                        //validar el correo 
                        if (Validaciones.ValidarMail(usuario.mail_usuario) == false)
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto };
                        }
                        else
                        {
                            //manejar transaccionalidad
                            using (var transaction = db.Database.BeginTransaction())
                            {
                                try
                                {

                                    //Crear el usuario Final
                                    Usuario usuarioFinal = new Usuario();

                                    usuarioFinal.id_usuario = usuario.id_usuario;
                                    usuarioFinal.nombre_usuario = usuario.nombre_usuario.ToUpper();
                                    usuarioFinal.apellido_usuario = usuario.apellido_usuario.ToUpper();
                                    usuarioFinal.tipo_usuario = usuario.tipo_usuario;
                                    usuarioFinal.cliente_asociado = usuario.cliente_asociado;
                                    usuarioFinal.pais = usuario.pais;
                                    usuarioFinal.ciudad = usuario.ciudad;
                                    usuarioFinal.direccion_usuario = usuario.direccion_usuario;
                                    usuarioFinal.mail_usuario = usuario.mail_usuario;
                                    usuarioFinal.telefono_usuario = usuario.telefono_usuario;
                                    usuarioFinal.celular_usuario = usuario.celular_usuario;
                                    usuarioFinal.rol_id = usuario.id_rol;
                                    if (usuario.tipo_usuario == 109)
                                    {
                                        usuarioFinal.area_departamento = usuario.area_departamento;
                                        usuarioFinal.cargo_usuario = usuario.cargo_usuario;
                                        usuarioFinal.secu_usua = usuario.secu_usua;
                                        usuarioFinal.link_firma = usuario.link_firma;
                                        usuarioFinal.validacion_correo = usuario.validacion_correo;
                                        usuarioFinal.cliente_asociado = null;
                                    }
                                    else
                                    {
                                        usuarioFinal.area_departamento = usuario.area_departamento_texto.ToUpper();
                                        usuarioFinal.cargo_usuario = usuario.cargo_usuario_texto.ToUpper();
                                        usuarioFinal.validacion_correo = usuario.validacion_correo;
                                    }

                                    usuarioFinal.clave_usuario = usuario.clave_usuario;
                                    usuarioFinal.estado_usuario = usuario.estado_usuario;
                                    usuarioFinal.activo_usuario = usuario.activo_usuario;
                                    usuarioFinal.codigo_usuario = usuario.codigo_usuario;
                                    usuarioFinal.id_empresa = usuario.id_empresa;
                                    usuarioFinal.reset_clave = usuario.reset_clave;

                                    // Por si queda el Attach de la entidad y no deja actualizar
                                    var local = db.Usuario.FirstOrDefault(f => f.id_usuario == usuarioFinal.id_usuario);
                                    if (local != null)
                                    {
                                        db.Entry(local).State = EntityState.Detached;
                                    }

                                    db.Entry(usuarioFinal).State = EntityState.Modified;
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
                    else
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida };
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion EliminarUsuario(int id)
        {
            try
            {
                var usuario = db.Usuario.Find(id);

                if (usuario.estado_usuario == true)
                {
                    usuario.estado_usuario = false;
                }
                else
                {
                    usuario.estado_usuario = true;
                }

                db.Entry(usuario).State = EntityState.Modified;
                db.SaveChanges();

                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static List<ListadoUsuarios> ListarUsuarios()
        {
            try
            {
                return db.ListadoUsuarios().ToList();
            }
            catch (Exception e)
            {
                e.ToString();
                throw;
            }
        }

        public static UsuarioCE ConsultarUsuario(int id)
        {
            try
            {
                //Buscar el usuario a Editar
                Usuario usuario = db.Usuario.Find(id);

                //Setear al Usuario General que se creo
                UsuarioCE usuarioGenerico = new UsuarioCE(); 

                usuarioGenerico.id_usuario = usuario.id_usuario;
                usuarioGenerico.nombre_usuario = usuario.nombre_usuario;
                usuarioGenerico.apellido_usuario = usuario.apellido_usuario;
                usuarioGenerico.tipo_usuario = usuario.tipo_usuario;
                usuarioGenerico.pais = usuario.pais;
                usuarioGenerico.ciudad = usuario.ciudad;
                usuarioGenerico.direccion_usuario = usuario.direccion_usuario;
                usuarioGenerico.mail_usuario = usuario.mail_usuario;
                usuarioGenerico.telefono_usuario = usuario.telefono_usuario;
                usuarioGenerico.celular_usuario = usuario.celular_usuario;
                usuarioGenerico.id_rol = Convert.ToInt32(usuario.rol_id);
                usuarioGenerico.id_empresa = Convert.ToInt32(usuario.id_empresa);

                usuarioGenerico.clave_usuario = usuario.clave_usuario;
                usuarioGenerico.estado_usuario = usuario.estado_usuario;
                usuarioGenerico.activo_usuario = usuario.activo_usuario; 
                usuarioGenerico.codigo_usuario = usuario.codigo_usuario;
                usuarioGenerico.validacion_correo = usuario.validacion_correo.Value;
                


                if (usuario.tipo_usuario == 109)
                {
                    usuarioGenerico.area_departamento = usuario.area_departamento;
                    usuarioGenerico.cargo_usuario = usuario.cargo_usuario;
                    usuarioGenerico.secu_usua = usuario.secu_usua.Value;
                    usuarioGenerico.link_firma = usuario.link_firma;
                }
                else
                {
                    usuarioGenerico.area_departamento_texto = usuario.area_departamento;
                    usuarioGenerico.cargo_usuario_texto = usuario.cargo_usuario;
                    if(usuario.cliente_asociado != null)
                    {
                        usuarioGenerico.cliente_asociado = usuario.cliente_asociado.Value;
                    }
                }

                return usuarioGenerico;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static usp_b_usuario ValidarUsuario(string usuario, string clave, string clavesinMD5)
        {
            try
            {
                var UsuarioLogin = db.usp_b_usuario(usuario, clave).FirstOrDefault();

                if (UsuarioLogin != null)
                {
                    //correo electronico solo para usuarios internos
                    if (UsuarioLogin.CodigoTipoUsuario == 109 && UsuarioLogin.validacion_correo == true && UsuarioLogin.reset_clave == false)
                    {
                        //validar el envio de correo
                        Boolean credenciales = validacionCorreo(UsuarioLogin.mail_usuario, clavesinMD5, "Usted ha ingresado al sistema Gestión_PPM");

                        if (credenciales == true)
                        {
                            return UsuarioLogin;
                        }
                        else
                        {
                            var UsuarioCorreo = db.usp_b_usuario("abc", "123").FirstOrDefault();
                            return UsuarioCorreo;
                        }
                    }
                    else
                    {
                        return UsuarioLogin;
                    }
                } else
                {
                    return UsuarioLogin;
                }
            }
            catch (Exception e)
            { 
                throw;
            }
        } 

        public static RespuestaTransaccion CrearUsuarioGenerico(RegisterViewModel usuario)
        {
            bool todoOk = true;
            try
            {
                try
                {
                    //validar si existe el registro
                    Usuario usuarioActual = db.Usuario.Where(u => u.mail_usuario == usuario.Email).First();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeResgistroExistente };
                }
                catch
                {
                    //validar que todos los campos esten completos
                    if (usuario.Nombre == null || usuario.Apellido == null || usuario.Email == null || usuario.Password == null || usuario.ConfirmPassword == null || (usuario.telefono_usuario == null && usuario.celular_usuario == null))
                    {
                        if (usuario.telefono_usuario == null && usuario.celular_usuario == null)
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatoriosTelefono };
                        }
                        else
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                        }
                    }
                    else
                    {
                        //Validar minimo 8 caracteres en la contraseña
                        if (usuario.Password.Length < 8)
                        {
                            todoOk = false;
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeLongitudContrasenia };
                        }

                        //si todos los campos estan llenos se procede a validar datos
                        if (todoOk == true)
                        {
                            //validar el correo 
                            if (Validaciones.ValidarMail(usuario.Email) == false)
                            {
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCorreoIncorrecto };
                            }
                            else
                            {
                                //Crear el usuario Final
                                Usuario usuarioFinal = new Usuario(); 

                                usuarioFinal.nombre_usuario = usuario.Nombre.ToUpper();
                                usuarioFinal.apellido_usuario = usuario.Apellido.ToUpper();
                                usuarioFinal.mail_usuario = usuario.Email;
                                //110 para usuario externo
                                usuarioFinal.tipo_usuario = 110;

                                //obtener la clave 
                                var clave = Validaciones.GetMD5_1(usuario.Password);

                                usuarioFinal.clave_usuario = clave;
                                usuarioFinal.estado_usuario = true;
                                usuarioFinal.activo_usuario = false;
                                usuarioFinal.id_empresa = 1;
                                usuarioFinal.reset_clave = false;
                                usuarioFinal.validacion_correo = false;
                                usuarioFinal.telefono_usuario = usuario.telefono_usuario;
                                usuarioFinal.celular_usuario = usuario.celular_usuario;

                                //obtener codigo usuario
                                var codigo = db.usp_g_codigo_usuario(usuarioFinal.tipo_usuario).First().codigoUsuario;
                                usuarioFinal.codigo_usuario = codigo.ToString();

                                //Obtener el Rol generico
                                Rol rol = db.Rol.Where(r => r.nombre_rol == "ROL GENÉRICO" || r.nombre_rol=="ROL GENERICO").First();
                                usuarioFinal.rol_id = rol.id_rol;

                                db.Usuario.Add(usuarioFinal);
                                db.SaveChanges();
                                int id_usuario = usuarioFinal.id_usuario;
                                  
                                //enviar correo
                                db.usp_envia_correo_usuario_generico(id_usuario);
                                db.SaveChanges();

                                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                            }
                        }
                        else
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion RecuperarClave(ForgotViewModel recuperar)
        {
            bool todoOk = true;
            try
            {
                try
                {
                    //validar si existe el registro
                    Usuario usuarioActual = db.Usuario.Where(u => u.mail_usuario == recuperar.Login).First();

                    //obtener la clave 
                    Random random = new Random();
                    String numero = "";
                    for (int i = 0; i < 9; i++)
                    {
                        numero += Convert.ToString(random.Next(0, 9));
                    }

                    var clave = Validaciones.GetMD5_1(numero);

                    usuarioActual.clave_usuario = clave;
                    usuarioActual.reset_clave = true;

                    // Por si queda el Attach de la entidad y no deja actualizar
                    var local = db.Usuario.FirstOrDefault(f => f.id_usuario == usuarioActual.id_usuario);
                    if (local != null)
                    {
                        db.Entry(local).State = EntityState.Detached;
                    }

                    db.Entry(usuarioActual).State = EntityState.Modified;
                    db.SaveChanges();

                    //enviar correo
                    db.usp_envia_correo_usuario_reset_clave(usuarioActual.id_usuario, numero);
                    db.SaveChanges();

                    return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeRecuperacionClave };

                }
                catch(Exception ex)
                {
                    ex.ToString();
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeUsuarioNoExiste };
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " ;" + ex.Message.ToString() };
            }
        }

        public static RespuestaTransaccion CambiarClave(CambiarClave usuario)
        {
            try
            {
                //validar datos obligatorios
                if (usuario.ContraseniaActual == null || usuario.ContraseniaNueva == null || usuario.ConfirmarContrasenia == null)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeDatosObligatorios };
                }
                else
                { 
                    //validar las contraseña nuevas coincidan
                    if (usuario.ContraseniaNueva != usuario.ConfirmarContrasenia)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionContrasenias };
                    }
                    else
                    {
                        if (usuario.ContraseniaNueva.Length < 8)
                        {
                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeLongitudContrasenia };
                        }
                        else
                        { 
                            //Datos actuales del usuario
                            var user = System.Web.HttpContext.Current.Session["usuario"] as String;

                            Usuario usuarioActual = new Usuario();
                            usuarioActual = db.Usuario.Find(Convert.ToInt32(user));

                            var actual = db.usp_b_datos_usuario(user.ToString());

                            //hacer md5 a la clave actual
                            var claveActual = Validaciones.GetMD5_1(usuario.ContraseniaActual);
                            var claveNueva = Validaciones.GetMD5_1(usuario.ContraseniaNueva);

                            //Validar que la calve nueva sea diferente a la anterior
                            if (usuario.ContraseniaActual == usuario.ContraseniaNueva)
                            {
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeCambioContaseña };
                            }
                            else
                            {
                                //Validar las claves si coinciden
                                if (claveActual == actual.Single().clave_usuario)
                                {
                                    //Se procede actualizar la clave
                                    usuarioActual.clave_usuario = claveNueva;

                                    //si es usuario interno valida credenciales del correo
                                    if (usuarioActual.tipo_usuario == 109 && usuarioActual.validacion_correo == true)
                                    {
                                        //validar el envio de correo
                                        Boolean credenciales = validacionCorreo(usuarioActual.mail_usuario, usuario.ConfirmarContrasenia, "Usted ha cambiado la clave del sistema Gestión_PPM");
                                        if (credenciales == true)
                                        {

                                            // Por si queda el Attach de la entidad y no deja actualizar
                                            var local = db.Usuario.FirstOrDefault(f => f.id_usuario == usuarioActual.id_usuario);
                                            if (local != null)
                                            {
                                                db.Entry(local).State = EntityState.Detached;
                                            }

                                            db.Entry(usuarioActual).State = EntityState.Modified;
                                            db.SaveChanges();

                                            db.usp_envia_correo_usuario_cambio_clave(usuarioActual.id_usuario);
                                            db.SaveChanges();

                                            return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                                        }
                                        else
                                        {
                                            return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionContraseniaActualOffice };
                                        }
                                    }
                                    else
                                    {
                                        // Por si queda el Attach de la entidad y no deja actualizar
                                        var local = db.Usuario.FirstOrDefault(f => f.id_usuario == usuarioActual.id_usuario);
                                        if (local != null)
                                        {
                                            db.Entry(local).State = EntityState.Detached;
                                        }

                                        db.Entry(usuarioActual).State = EntityState.Modified;
                                        db.SaveChanges();

                                        db.usp_envia_correo_usuario_cambio_clave(usuarioActual.id_usuario);
                                        db.SaveChanges();

                                        return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                                    }
                                }
                                else
                                {
                                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionContraseniaActual };
                                }
                            }
                        }
                    }
                } 
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida +" "+ ex.ToString() };
                throw;
            }
        } 

        public static List<usp_b_opci_menu_usua> OpcionesMenuUsuario(int id_usuario)
        {
            try
            {
                return db.usp_b_opci_menu_usua(id_usuario).ToList();
            }
            catch (Exception e)
            {
                e.ToString();
                throw;
            }
        }
         
        public static bool VerificarCorreoUsuarioExistente(string correo)
        {
            try
            {
                var usuario = db.Usuario.Where(s => s.mail_usuario == correo && s.estado_usuario.Value).ToList();
                if (usuario.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public static RespuestaTransaccion CambiarClaveReset(CambiarClave usuario)
        {
            try
            { 
                //Obtener los datos del Usuario
                var datos = db.Usuario.Where(u => u.id_usuario == usuario.UsuaCodi).First();

                //validar las contraseña nuevas coincidan
                if (usuario.ContraseniaNueva != usuario.ConfirmarContrasenia)
                {
                    return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionContrasenias };
                }
                else
                {
                    if (usuario.ContraseniaNueva.Length < 8)
                    {
                        return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeLongitudContrasenia };
                    }
                    else
                    {
                        //solo si es interno valida correo
                        if (datos.tipo_usuario == 109 && datos.validacion_correo == true)
                        {
                            //validar el envio de correo
                            bool credenciales = validacionCorreo(datos.mail_usuario, usuario.ConfirmarContrasenia, "Se ha cambiado la clave temporal asignada en el sistema Gestión_PPM");

                            if (credenciales == true)
                            {
                                var claveNueva = Validaciones.GetMD5_1(usuario.ContraseniaNueva);

                                //Datos actuales del usuario  
                                Usuario usuarioActual = new Usuario();
                                usuarioActual = db.Usuario.Find(usuario.UsuaCodi);
                                usuarioActual.reset_clave = false;
                                usuarioActual.clave_usuario = claveNueva;

                                db.Entry(usuarioActual).State = EntityState.Modified;
                                db.SaveChanges();

                                db.usp_envia_correo_usuario_cambio_clave(usuarioActual.id_usuario);
                                db.SaveChanges();

                                return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                            }
                            else
                            {
                                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeValidacionContraseniaActualOffice };
                            }
                        }
                        else
                        {
                            var claveNueva = Validaciones.GetMD5_1(usuario.ContraseniaNueva);

                            //Datos actuales del usuario  
                            Usuario usuarioActual = new Usuario();
                            usuarioActual = db.Usuario.Find(usuario.UsuaCodi);
                            usuarioActual.reset_clave = false;
                            usuarioActual.clave_usuario = claveNueva;

                            db.Entry(usuarioActual).State = EntityState.Modified;
                            db.SaveChanges();

                            db.usp_envia_correo_usuario_cambio_clave(usuarioActual.id_usuario);
                            db.SaveChanges();

                            return new RespuestaTransaccion { Estado = true, Respuesta = Mensajes.MensajeTransaccionExitosa };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return new RespuestaTransaccion { Estado = false, Respuesta = Mensajes.MensajeTransaccionFallida + " " + ex.ToString() };
                throw;
            }
        }

        public static Boolean validacionCorreo(string mailDirection, string clave, string texto)
        {
            try
            {
                MailMessage mailMessage = new MailMessage();

                mailMessage.From = new MailAddress(mailDirection);
                mailMessage.To.Add(new MailAddress("notificacionesgestor@ppm.com.ec"));
                mailMessage.Subject = "Ingreso al Sistema Gestion PPM";
                mailMessage.Body = texto;
                 
                SmtpClient client = new SmtpClient("smtp-mail.outlook.com");
                client.Port = 587;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false; 
                client.EnableSsl = true;  
                client.Credentials = new System.Net.NetworkCredential(mailDirection, clave);

                client.Send(mailMessage);

            }
            catch (Exception ex)
            {
                ex.ToString();
                return false;
            }
            finally
            {
            }
            return true;
        }

        public static List<ListadoUsuarios> ListarUsuariosInternos()
        {
            try
            {
                return db.ListadoUsuarios().Where(u => u.Tipo_Usuario == "INTERNO").ToList();
            }
            catch (Exception e)
            {
                e.ToString();
                throw;
            }
        }

        public static List<OrganigramaParcial> GetUsuarioAdministradorPrincipalOrganigrama(int id)
        {
            List<OrganigramaParcial> usuario = new List<OrganigramaParcial>();
            try
            {
                var administrador = db.ListadoUsuarios().Where(s => s.Id == id).FirstOrDefault(); // ID DEL ADMINISTRADOR
                usuario.Add(new OrganigramaParcial
                {
                    id = administrador.Id,
                    cargo = administrador.Cargo,
                    //parent = null,
                    codigo = administrador.Codigo,
                    departamento = administrador.Area_o_Departamento,
                    usuario = administrador.Nombres_Completos,
                    mail = administrador.Mail,
                });
                return usuario;
            }
            catch (Exception e)
            {
                e.ToString();
                return usuario;
            }
        }

        public static List<ListadoUsuarios> ListarUsuariosOrganigrama(List<int> ids, string tipo = null)
        {
            try
            {
                if (string.IsNullOrEmpty(tipo))
                    return db.ListadoUsuarios().Where(s => !ids.Contains(s.Id)).ToList();
                else
                    return db.ListadoUsuarios().Where(s => !ids.Contains(s.Id) && s.Tipo_Usuario == tipo).ToList();
            }
            catch (Exception e)
            {
                e.ToString();
                throw;
            }
        }

        public static UsuarioCE ConsultarInformacionPrincipalUsuario(int id)
        {
            UsuarioCE Usuario = new UsuarioCE();
            try
            {
                //Buscar el usuario a Editar
                var _usuario = db.ListadoUsuarios().Where(s => s.Id == id).FirstOrDefault();
                Usuario = new UsuarioCE
                {
                    nombre_usuario = _usuario.Nombres_Completos,
                    NombreEmpresa = _usuario.NombreEmpresa,
                    cargo_usuario = _usuario.Cargo,
                    id_usuario = _usuario.Id,
                    usar_organigrama = _usuario.usar_organigrama.Value,
                    cliente_asociado = _usuario.id_cliente,
                };

                return Usuario;
            }
            catch (Exception ex)
            {
                return Usuario;
            }
        }

    }
}