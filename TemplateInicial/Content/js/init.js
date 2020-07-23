//Alertas Adicionales de boostrap.
var _ALERT = {
    SUCCESS: 'alert-success',
    INFO: 'alert-info',
    WARNING: 'alert-warning',
    DANGER: 'alet-danger',
};

//Alertas TOASTR para validaciones y respuestas de operaciones asíncronas.
var _ALERT_TOASTR = {
    SUCCESS: {
        Type: 1,
        Title: 'ÉXITO',
        Message: 'Sin definir',
        Options: {
            closeButton: true,
            progressBar: true,
            positionClass: "toast-top-right",
            timeOut: 1500,
            extendedTimeOut: 3500,
            onHidden: function () {
                //Cuando termina la animación principal se remueven todas las restantes.
                toastr.clear();
            },
            // Define a callback for when the toast is shown/hidden/clicked
            onShown: function () {
                console.log('onShown');
            },
            onclick: function () {
                console.log('onclick');
            },
            onCloseClick: function () {
                console.log('onCloseClick');
            },
        }
    },
    INFO: {
        Type: 2,
        Title: 'INFORMACIÓN',
        Message: 'Sin definir',
        Options: {
            closeButton: true,
            progressBar: false,
            positionClass: "toast-top-right",
            timeOut: 0,
            extendedTimeOut: 0,
            onHidden: function () {
                //Cuando termina la animación principal se remueven todas las restantes.
                console.log('onHidden');
            },
            // Define a callback for when the toast is shown/hidden/clicked
            onShown: function () {
                console.log('onShown');
            },
            onclick: function () {
                console.log('onclick');
            },
            onCloseClick: function () {
                console.log('onCloseClick');
            },
        }
    },
    WARNING: {
        Type: 3,
        Title: 'ADVERTENCIA',
        Message: 'Sin definir',
        Options: {
            closeButton: true,
            progressBar: false,
            positionClass: "toast-top-right",
            timeOut: 0,
            extendedTimeOut: 0,
            onHidden: function () {
                //Cuando termina la animación principal se remueven todas las restantes.
                console.log('onHidden');
            },
            // Define a callback for when the toast is shown/hidden/clicked
            onShown: function () {
                console.log('onShown');
            },
            onclick: function () {
                console.log('onclick');
            },
            onCloseClick: function () {
                console.log('onCloseClick');
            },
        }
    },
    DANGER: {
        Type: 4,
        Title: 'ERROR',
        Message: 'Sin definir',
        Options: {
            closeButton: true,
            progressBar: true,
            positionClass: "toast-top-right",
            timeOut: 4000,
            extendedTimeOut: 4000,
            preventDuplicates: true,
            onHidden: function () {
                //Cuando termina la animación principal se remueven todas las restantes.
                console.log('onHidden');
            },
            // Define a callback for when the toast is shown/hidden/clicked
            onShown: function () {
                console.log('onShown');
            },
            onclick: function () {
                console.log('onclick');
            },
            onCloseClick: function () {
                console.log('onCloseClick');
            },
        }
    },

    ShowAlert: function (TYPE, message, title) {

        if (typeof TYPE.Type == "undefined") {
            console.log(TYPE)
            alert("TIPO DE ALERTA NO DEFINIDA");
            return;
        }

        if (title)
            TYPE.Title = title;
        if (message)
            TYPE.Message = message;

        switch (TYPE.Type) {
            case 1:
                toastr.success(TYPE.Message, TYPE.Title, TYPE.Options)
                break;
            case 2:
                toastr.info(TYPE.Message, TYPE.Title, TYPE.Options)
                break;
            case 3:
                toastr.warning(TYPE.Message, TYPE.Title, TYPE.Options)
                break;
            case 4:
                toastr.error(TYPE.Message, TYPE.Title, TYPE.Options)
                break;
            default:
                console.log(TYPE)
                alert("TIPO DE ALERTA NO DEFINIDA");
                break;
        }
    },
};

var _HTTP_METHOD = {
    POST: 'POST',
    GET: 'GET',
};

var _TYPE_REQUEST = {
    JSON: 'json',
};

//overlayAwaitActive -> TRUE : Cuando se use en un formulario que requiera esperar la respuesta del servidor (Esperando ..)
var Common = {
    Ajax: function (httpMethod, url, data, type, successCallBack, overlayAwaitActive, messageAwait, async, cache, contentType) {
        if (typeof async == "undefined") {
            async = true;
        }
        if (typeof cache == "undefined") {
            cache = false;
        }
        if (typeof contentType == "undefined") {
            contentType = 'application/json; charset=utf-8';
        }

        if (typeof messageAwait == "undefined" && overlayAwaitActive) {
            messageAwait = 'Por favor espere ..';
        }

        if (overlayAwaitActive)
            ShowOnlyAwait(messageAwait)

        var ajaxObj = $.ajax({
            type: httpMethod.toUpperCase(),
            url: url,
            data: data,
            dataType: type,
            async: async,
            cache: cache,
            contentType: contentType,
        }).done(successCallBack).fail(function (jqXHR, textStatus, errorThrown) {
            if (jqXHR.status === 0) {
                Common.AjaxFailureCallback(`Not connect: Verify Network. Status Response:  ${jqXHR.status}`);
            } else if (jqXHR.status == 404) {
                Common.AjaxFailureCallback(`Requested page not found [404]. Status Response:  ${jqXHR.status}`);
            } else if (jqXHR.status == 500) {
                Common.AjaxFailureCallback(`Internal Server Error [500]. Status Response:  ${jqXHR.status}`);
            } else if (textStatus === 'parsererror') {
                Common.AjaxFailureCallback(`'Requested JSON parse failed. Response:  ${textStatus}`);
            } else if (textStatus === 'timeout') {
                Common.AjaxFailureCallback(`'Time out error. Response:  ${textStatus}`);
            } else if (textStatus === 'abort') {
                Common.AjaxFailureCallback(`'Ajax request aborted. Response:  ${textStatus}`);
            } else {
                Common.AjaxFailureCallback(`'Uncaught Error:  ${jqXHR.responseText}`);
            }

        }).always(function () {
            HideAwaitTime(250);
        });

        return ajaxObj;
    },

    DisplayAlert: function (type, message, title) {
        //toastr.clear(); // Limpiar alertas previas
        _ALERT_TOASTR.ShowAlert(type, message, title)
    },

    AjaxFailureCallback: function (failureMessage) {
        console.log(failureMessage);
        _ALERT_TOASTR.ShowAlert(_ALERT_TOASTR.DANGER, failureMessage)
    },

    ShowFailSavedMessage: function (messageText) {
        $.blockUI({
            css: {
                border: 'none',
                padding: '15px',
                backgroundColor: '#000',
                '-webkit-border-radius': '10px',
                '-moz-border-radius': '10px',
                opacity: .5,
                color: '#fff'
            }
        });

        HideAwaitTime(1500);
    },
}

function ShowOnlyAwait(messageText) {
    $.blockUI({
        baseZ: 2000,
        message: messageText,
        css: {
            border: 'none',
            padding: '15px',
            backgroundColor: '#000',
            '-webkit-border-radius': '10px',
            '-moz-border-radius': '10px',
            opacity: .5,
            color: '#fff'
        }
    });
}

function HideAwaitTime(time) {
    setTimeout($.unblockUI, time);
}


function DownloadFiles(files, urlAccion, delay = 1000) {
    var contador = delay;
    for (var i = 0; i < files.length; i++) {
        let url = urlAccion + "?path=" + files[i];
        setTimeout(function () {
            location.href = url;
        }, contador);
        contador += delay;
    }
}

function showAlertAditionals(alerttype, header, message, footer, timeOut = 5000) {
    debugger
    $('.content').prepend('<div id="alertdiv" role="alert" class="alert ' + alerttype + '"><button type="button" class="close" data-dismiss="alert" aria-label="Close"> <span aria-hidden="true">&times;</span> </button> <h4 class="alert-heading">' + header + '</h4> <span id="texto-alerta">' + message + '</span> <hr>  <small>' + footer + '</small></div>')
    if (timeOut > 0) {
        setTimeout(function () { // this will automatically close the alert and remove this if the users doesnt close it in 5 secs
            $("#alertdiv").remove();
        }, timeOut);
    }
}

function convertToBase64(element) {
    debugger
    var id = 0;//parseInt($(element).attr('id'));
    var f = $(element)[0].files[0]; // FileList object

    if (f.size > 2097152) {
        toastr.warning('Tamaño de archivo no permitido. Verifique que tenga menos de 2MG.')
        $(element).val("");
        return;
    };

    var reader = new FileReader();
    // Closure to capture the file information.
    reader.onload = (function (theFile) {
        return function (e) {
            var binaryData = e.target.result;
            //Converting Binary Data to base 64
            var base64String = window.btoa(binaryData);
            //showing file converted to base64
            //document.getElementById('base64').value = base64String;
            archivosAdjuntos[id] = base64String
            //alert('File converted to base64 successfuly!\nCheck in Textarea');
        };
    })(f);
    // Read in the image file as a data URL.
    reader.readAsBinaryString(f);
}

$(function () {
    $(".datepicker").datepicker({
        autoclose: true
    });

    //$('input[type="checkbox"], input[type="radio"]').iCheck({
    //    checkboxClass: "icheckbox_minimal-blue",
    //    radioClass: "iradio_minimal-blue"
    //});

    $("#datemask").inputmask("dd/mm/yyyy", { "placeholder": "dd/mm/yyyy" });
    $("#datemask2").inputmask("mm/dd/yyyy", { "placeholder": "mm/dd/yyyy" });
    $("[data-mask]").inputmask();

    $(".campo-decimal-manual-1").inputmask('Regex', { regex: "^[0-9]{1,9}(\\.\\d{1,4})?$" });
    $(".campo-decimal-manual-2").inputmask('Regex', { regex: "^[0-9]{1,9}(\\,\\d{1,4})?$" });

    $(".campo-entero").inputmask('Regex', { regex: "^[0-9]+$" });

    // Fix sidebar white space at bottom of page on resize
    $(window).on("load", function () {
        setTimeout(function () {
            $("body").layout("fix");
            $("body").layout("fixSidebar");
        }, 250);
    });
});