
"use strict";
$('#clear').click(clearCircuit);
$('#embed').click(insertCircuitFromInput);

(function () { 
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var circuit = Office.context.document.settings.get('circuit');
            if (circuit) {
                insertSavedCircuit(circuit);
            }
            else { 
                showInsertControls();
            }
        });
};
})();

function showInsertControls(){ 
    $('#input-box').removeClass('hidden');
    $('#embed').removeClass('hidden');
    $('#clear').addClass('hidden');
}
// Insert circuit by accepting circuitID from text input
function insertCircuitFromInput(circuitID) {
    var circuitID = $('#circuitID').val();
    var src = 'https://circuitverse.org/simulator/embed/'+circuitID;
    var iframe = '<iframe width="800" height="400" src="'+src+'" scrolling="no" webkitAllowFullScreen mozAllowFullScreen allowFullScreen> </iframe>';
    $('#content-main').html(iframe);
    Office.context.document.settings.set('circuit', circuitID);
    persistChanges();
    $('#input-box').addClass('hidden');
    $('#embed').addClass('hidden');
    $('#clear').removeClass('hidden');
}

// Insert a saved circuit
function insertSavedCircuit(circuitID){
    var src = 'https://circuitverse.org/simulator/embed/'+circuitID;
    var iframe = '<iframe width="800" height="365" src="'+src+'" scrolling="no" webkitAllowFullScreen mozAllowFullScreen allowFullScreen> </iframe>';
    $('#content-main').html(iframe);
    $('#clear').removeClass('hidden');
}

function clearCircuit() { 
    Office.context.document.settings.remove('circuit');
    persistChanges();
    showInsertControls();
    $('#content-main').html(" <img id = 'background' src ='https://circuitverse.org/img/circuitverse_black.svg' >")
}

function persistChanges() { 
    Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log('Settings save failed. Error: ' + asyncResult.error.message);
        } else {
            console.log('Settings saved.');
        }
    });
}