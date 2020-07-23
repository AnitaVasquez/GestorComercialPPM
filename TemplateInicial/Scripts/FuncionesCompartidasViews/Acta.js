function CalculoEstadisticasParticipantes(tablaID, mensajeActaSuspendida = "") {
    
    var listadoParticipantes = GetListadoTablaDinamica(tablaID);

    debugger

    let totalListado = listadoParticipantes.length;

    if (totalListado > 0) {
        let contadorEleccionSI = 0;
        let contadorEleccionNO = 0;
        for (var i = 0; i < listadoParticipantes.length; i++) {
            let presente = listadoParticipantes[i].Presente
            if (presente == "true")
                contadorEleccionSI++;
            else
                contadorEleccionNO++;
        }

        let porcentajeSI = parseFloat((contadorEleccionSI / totalListado) * 100);
        let porcentajeNO = parseFloat((contadorEleccionNO / totalListado) * 100);

        if (porcentajeNO > porcentajeSI)
            $("#ResolucionParticipantes").text(mensajeActaSuspendida);

        if (porcentajeNO == porcentajeSI)
            $("#ResolucionParticipantes").text("");

        porcentajeSI = porcentajeSI.toFixed(2);
        porcentajeNO = porcentajeNO.toFixed(2);

        $("#PorcentajeParticipantes-EleccionSI").fadeOut(300);
        $("#PorcentajeParticipantes-EleccionSI").text("SI " + porcentajeSI + "%");
        $("#PorcentajeParticipantes-EleccionSI").fadeIn(300);

        $("#PorcentajeParticipantes-EleccionNO").fadeOut(300);
        $("#PorcentajeParticipantes-EleccionNO").text("NO " + porcentajeNO + "%");
        $("#PorcentajeParticipantes-EleccionNO").fadeIn(300);
    } else {
        $("#PorcentajeParticipantes-EleccionSI").fadeOut(300);
        $("#PorcentajeParticipantes-EleccionSI").text("SI 0%");
        $("#PorcentajeParticipantes-EleccionSI").fadeIn(300);

        $("#PorcentajeParticipantes-EleccionNO").fadeOut(300);
        $("#PorcentajeParticipantes-EleccionNO").text("NO 0%");
        $("#PorcentajeParticipantes-EleccionNO").fadeIn(300);

        $("#ResolucionParticipantes").text("");
    }
}

function validacionDetallesActaInicioProyecto() {
    let validacion1 = TablaDinamicaVacia("tbl-ResponsablesCliente")
    let validacion2 = TablaDinamicaVacia("tbl-Entregables")
    let validacion3 = TablaDinamicaVacia("tbl-CondicionesGenerales")

    if (validacion1 || validacion2 || validacion3)
        return false;
    else
        return true;
}

function validacionDetallesActaCierreProyecto() {
    let validacion1 = TablaDinamicaVacia("tbl-ResponsablesCliente")
    let validacion2 = TablaDinamicaVacia("tbl-Entregables")
    if (validacion1 || validacion2)
        return false;
    else
        return true;
}

function validacionDetallesActaReunion() {
    let validacion1 = TablaDinamicaVacia("tbl-Participantes")
    let validacion2 = TablaDinamicaVacia("tbl-Temas")
    let validacion3 = TablaDinamicaVacia("tbl-Acuerdos")

    if (validacion1 || validacion2 || validacion3)
        return false;
    else
        return true;
}

function validacionDetallesActaCliente() {
    let validacion = TablaDinamicaVacia("tbl-Cliente")
    if (validacion)
        return false;
    else
        return true;
}

function validacionDetallesActaContabilidad() {
    let validacion = TablaDinamicaVacia("tbl-Contabilidad")
    if (validacion)
        return false;
    else
        return true;
}

function firmasVacias() {
    debugger
    let firmas = $('.firma').editable('getValue');
    let validacion = CompararObjetos(valorInicialFirmas, firmas);
    return validacion;
}

function GetFirmasActa() {
    let firmas = [];
    var firma1 = $('.firma-persona-general').editable('getValue');
    var firma2 = $('.firma-usuario').editable('getValue');
    firmas.push(firma1);
    firmas.push(firma2);
    return JSON.stringify(firmas);
}


function cargarFirmaUsuario(firmas) {
    $('#cargo').editable('setValue', firmas[0].cargo);
    $('#empresa').editable('setValue', firmas[0].empresa);
    $('#nombre').editable('setValue', firmas[0].nombre);

    $('#usuarioCargo').editable('setValue', firmas[1].usuarioCargo);
    $('#usuarioEmpresa').editable('setValue', firmas[1].usuarioEmpresa);
    $('#usuarioNombre').editable('setValue', firmas[1].usuarioNombre);
}

function limpiarCamposDependientesCodigoCotizacion() {
    $('.campo-dependiente-busqueda-codigo-cotizacion').val("");
}
