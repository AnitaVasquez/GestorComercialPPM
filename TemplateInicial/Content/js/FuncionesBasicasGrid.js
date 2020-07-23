//Recarga todos los grids del documento
function recargarGrids() {
    debugger
    var enDocumento = document.querySelector('.mvc-grid');
    if (enDocumento !== null) {
        var grid = new MvcGrid(document.querySelector('.mvc-grid'));
        grid.query.set('search', this.value);
        //Do things
        [].forEach.call(document.getElementsByClassName('mvc-grid'), function (element) {
            debugger
            new MvcGrid(element).reload();
        });
    }
}

// Recarga el grid por el parametro ID
function recargarGridByID(idGrid) {
    debugger
    var grid = new MvcGrid(document.querySelector("#" + idGrid));
    grid.query["parameters"] = [];
    grid.query.set('search', '');
    grid.reload();
    //Do things
    //[].forEach.call(document.getElementsByClassName('mvc-grid'), function (element) {
    //    debugger
    //    new MvcGrid(element).reload();
    //});
}

// Recarga el grid con varios parametros de catalogo
function recargarGridByCatalogos(idGrid, tipo, subcatalogo, filtro) {
    debugger
    var grid = new MvcGrid(document.querySelector("#" + idGrid));
    grid.query["parameters"] = [];
    grid.query.set('search', '');
    grid.query.set('tipo', tipo);
    grid.query.set('subcatalogo', subcatalogo);
    grid.query.set('filtro', filtro);
    grid.reload();
    //Do things
    //[].forEach.call(document.getElementsByClassName('mvc-grid'), function (element) {
    //    debugger
    //    new MvcGrid(element).reload();
    //});
}


//Búsqueda general en Grid
function busquedaGrid(ID) {
    debugger
    grid = document.getElementById('GridSearch');
    if (grid !== null) {
        document.getElementById('GridSearch').addEventListener('input', function () {
            debugger
            var grid = new MvcGrid(document.querySelector('#' + ID));
            grid.query.set('search', this.value);
            grid.reload();
        });
    } else {
        return;
        //toastr.error("Ocurrió un error al cargar los datos.")
    }

}

// Recarga el grid con varios parametros de catalogo
function recargarGridByPermisos(idGrid, rol, filtro) {
    debugger
    var grid = new MvcGrid(document.querySelector("#" + idGrid));
    grid.query["parameters"] = [];
    grid.query.set('search', '');
    grid.query.set('rol', rol);
    grid.query.set('filtro', filtro);
    grid.reload();
    //Do things
    //[].forEach.call(document.getElementsByClassName('mvc-grid'), function (element) {
    //    debugger
    //    new MvcGrid(element).reload();
    //});
}

function reporteGridPDF(urlAccion) {
    debugger
    $.ajax({
        type: "POST",
        url: urlAccion,
        content: "application/json; charset=utf-8",
        dataType: "json",
        //data: JSON.stringify(data),
        success: function (data) {
            debugger
            var atributos = Object.keys(data[0])
            var propiedades = [];

            for (var i = 0; i < atributos.length; i++) {
                propiedades.push({
                    'field': atributos[i],
                    'displayName': atributos[i],
                })
            }

            //JSON.parse(resultado)
            //console.log(resultado.Data)
            printJS({
                printable: data,
                properties: propiedades,
                type: 'json'
            })
        },
        error: function (xhr, textStatus, errorThrown) {
            debugger
            toastr.error("Ocurrió un error.")
        }
    });
}

function buscarGrid(valor) {
    var grid = new MvcGrid(document.querySelector('.mvc-grid'));
    grid.query.set('search', valor);
    grid.reload();
}