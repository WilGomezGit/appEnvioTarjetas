async function intercalarPDFs() {
    const pdfFile1 = document.getElementById('pdfFile1').files[0];
    const pdfFile2 = document.getElementById('pdfFile2').files[0];

    if (!pdfFile1 || !pdfFile2) {
        alert('Por favor, selecciona los dos archivos PDF.');
        return;
    }

    const pdf1Bytes = await pdfFile1.arrayBuffer();
    const pdf2Bytes = await pdfFile2.arrayBuffer();

    const pdfDoc1 = await PDFLib.PDFDocument.load(pdf1Bytes);
    const pdfDoc2 = await PDFLib.PDFDocument.load(pdf2Bytes);

    const pdfDocResult = await PDFLib.PDFDocument.create();

    const numPagesPdf1 = pdfDoc1.getPageCount();
    const numPagesPdf2 = pdfDoc2.getPageCount();

    const maxPages = Math.max(numPagesPdf1, numPagesPdf2);

    for (let i = 0; i < maxPages; i++) {
        if (i < numPagesPdf1) {
            const [pdf1Page] = await pdfDocResult.copyPages(pdfDoc1, [i]);
            pdfDocResult.addPage(pdf1Page);
        }
        if (i < numPagesPdf2) {
            const [pdf2Page] = await pdfDocResult.copyPages(pdfDoc2, [i]);
            pdfDocResult.addPage(pdf2Page);
        }
    }

    const pdfBytes = await pdfDocResult.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    saveAs(blob, 'PDF_Intercalado.pdf');

    // Mostrar mensaje de éxito
    document.getElementById('message').textContent = 'PDFs intercalados y unidos exitosamente.';
}


function clearInputs() {
    document.getElementById('pdfFile1').value = '';
    document.getElementById('pdfFile2').value = '';
    document.getElementById('message').textContent = '';
}
