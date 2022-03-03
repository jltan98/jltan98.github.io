function showPDF(filetype) {
    var closeDoc = document.getElementById("close");
    closeDoc.style.display = 'block';

    var pdf = document.getElementById('pdf')
    file_name = 'doc/' + filetype + '.pdf'
    pdf.innerHTML =
        "<embed type='application/pdf' class='display' src='" + file_name + "' width='100%' height='800px' />"
}

function closeDoc() {
    var closeDoc = document.getElementById("close");
    var pdf = document.getElementById('pdf')
    pdf.innerHTML = "";
    closeDoc.style.display = 'none';
}
