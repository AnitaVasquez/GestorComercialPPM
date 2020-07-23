/*
	By Osvaldas Valutis, www.osvaldas.info
	Available for use under the MIT License
*/

'use strict';

;( function ( document, window, index )
{
    debugger
	var inputs = document.querySelectorAll( '.inputfile' );
	Array.prototype.forEach.call( inputs, function( input )
    {
		var label	 = input.nextElementSibling,
			labelVal = label.innerHTML;

		input.addEventListener( 'change', function( e )
        {
            var flag = true;
            //var ext = this.value.match(/\.([^\.]+)$/)[1];
            //console.log(ext)

            //if (ext === 'rar' || ext === 'zip') {
            //    swal("Error en el tipo de archivo!", "Formato de archivo incorrecto.", "info").then((value) => {
            //        location.reload();
            //    });
            //}
             
            $("#Submit").show();

			var fileName = '';
			if( this.files && this.files.length > 1 )
				fileName = ( this.getAttribute( 'data-multiple-caption' ) || '' ).replace( '{count}', this.files.length );
			else
				fileName = e.target.value.split( '\\' ).pop();

            
            //if (fileName) {
            //    label.querySelector('span').innerHTML = fileName;
            //    console.log('file name' + fileName)
            //}
            //else {
            //    console.log('labelVal' + labelVal)
            //    label.innerHTML = labelVal;
            //}
		});

		// Firefox bug fix
		input.addEventListener( 'focus', function(){ input.classList.add( 'has-focus' ); });
		input.addEventListener( 'blur', function(){ input.classList.remove( 'has-focus' ); });
	});
}( document, window, 0 ));