var listado = [];

$(document).ready(function () {
    console.log("Script tabla dinámica inicializado.")
});

function agregarFila(idTemplate, idTabla, contador) {
    try {
        debugger
        //$('#' + idTemplate)
        //    .clone()                        // Clona elemento del DOM
        //    .attr('id', 'row' + (contador))    // Setea ID unico
        //    .appendTo($('#' + idTabla + ' tbody'))  // Agrega la fila
        //    .show();

        var valor = $('#' + idTemplate).find('select').val();

        var clonar = $('#' + idTemplate).clone();

        clonar.find('select').val(valor);

        clonar.attr('id', 'row' + (contador))    // Setea ID unico
            .appendTo($('#' + idTabla + ' tbody'))  // Agrega la fila
            .show();

        var fila = $('#row' + contador); // Captura la fila
        var celdaAcciones = $(fila).find("#accion"); // Captura la celda de las acciones
        celdaAcciones.html('<button type="button" id="eliminar' + contador + '" class="eliminarFila btn btn-danger"><i class="fa fa-trash-o" aria-hidden="true"></i></button>'); // Reemplaza la celda por el botón eliminar

        limpiarControles(idTemplate); // Limpiar Valores

        $('#eliminar' + contador).click(function (e) {
            e.preventDefault();
            var idFila = 'eliminar' + contador;
            eliminarFila(idFila)

            //if (idTabla == "tbl-Participantes")
            //    CalculoEstadisticasParticipantes(idTabla)

        });

        $(".informacion-prefactura").click(function (e) {
            debugger
            let elemento = $(e.currentTarget);

            var fila = elemento.closest("tr"); // fila
            var ddlPrefactura = fila.find("select"); // dropdownlist
            var valor = parseInt($(ddlPrefactura).val()); // valor seleccionado

            if (!valor) {
                //toastr.error('Los campos con * son de ingreso obligatorio.')
                return;
            }

            _GetCreate({ flag: null, id: valor }, urlAccionInformacionPrefactura);
            $('#contenido-modal').modal({
                'show': 'true',
                'backdrop': 'static',
                'keyboard': false
            });
            return;
        })

    }
    catch (error) {
        console.log(error);
    }
}

function limpiarControles(idTemplate) {
    $("#" + idTemplate).find("input[type=text],select").not(".idObjeto").val(""); // Inputs textbox
    $("#" + idTemplate + " input[type='checkbox']").attr('checked', false)
    //$("select option:first-child").attr("selected", "selected");

    //$('#' + idTemplate + ' .chk-presente').iCheck('uncheck');
}

function eliminarFila(idElemento) {
    debugger
    try {
        var idTabla = $("#" + idElemento).closest('table').attr('id');

        $("#" + idElemento).closest("tr").remove();

        if (idTabla == "tbl-Participantes") {
            CalculoEstadisticasParticipantes(idTabla)
        }

    } catch (error) {
        console.log(error)
    }
}

function GetListadoTablaDinamica(tablaID) {
    try {
        debugger
        var tabla = $("#" + tablaID);

        listado = tabla.tableToJSON({
            extractor: function (cellIndex, $cell) {
                debugger
                // get text from the span inside table cells;
                // if empty or non-existant, get the cell text
                return $cell.find('input,select').val() || $cell.text() /*|| $('input[type=checkbox]').val()*/
            }
        });
        debugger
        //Eliminando el primer elemento #template(fila vacía)
        if (listado.length > 0)
            listado = listado.slice(1)

        return listado;
    } catch (e) {
        console.log(e);
        return listado;
    }
}

//Valida que todos los inputs sean obligatorios
function TablaDinamicaVacia(tablaID) {
    try {
        debugger
        let elementoVacio = false;
        var tabla = $("#" + tablaID);

        listado = tabla.tableToJSON({
            extractor: function (cellIndex, $cell) {
                debugger
                var numeroFila = $cell.closest("tr").index();

                if ($cell.find('input,select').val() == "" && $cell.text() == "" && numeroFila > 0) {
                    //console.log($cell.find('input'))
                    //console.log($cell.text());
                    elementoVacio = true;
                }

                // get text from the span inside table cells;
                // if empty or non-existant, get the cell text
                return $cell.find('input').val() || $cell.text()
            }
        });

        if (elementoVacio)
            return true;

        //Eliminando el primer elemento #template(fila vacía)
        if (listado.length > 0)
            listado = listado.slice(1)

        //Eliminando el primer elemento #template(fila vacía)
        if (listado.length > 0)
            return false;
        else
            return true;

    } catch (e) {
        console.log(e);
        return true;
    }
}


function SetListadoTablaDinamica(lista) {
    listado = lista;
}

function filaVacia(idTemplate) {
    debugger
    var objetos = $("#" + idTemplate).find("input,select");
    let flag = false;
    for (var i = 0; i < objetos.length; i++) {
        let elemento = $(objetos[i]);
        let clase = elemento.attr('class');
        //if (!elemento.val() && clase !== 'chosen-search-input') // --> PARA COMPONENTE CON BUSQUEDA
        if (!elemento.val())
            flag = true;
    }
    return flag;
}

function DetallesPendientes(ids) {
    debugger
    let flag = false;
    $.each(ids, function (index, value) {
        debugger
        var objetos = $("#" + value).find("input,select");
        for (var i = 0; i < objetos.length; i++) {
            let elemento = $(objetos[i]);
            if (elemento.val() && !$(elemento[0]).hasClass('idObjeto') && !$(elemento[0]).hasClass('tipoFecha') && !$(elemento[0]).hasClass('chk-presente')) {
                flag = true;
                console.log("seccion: " + value)
                console.log(elemento)
            }
        }
    })
    return flag;
}
