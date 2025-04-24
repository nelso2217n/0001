<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Sistema de Contaduría Mejorado</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js"></script>
  <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
  <style>
    input, select, button {
      margin: 5px;
    }
  </style>
</head>
<body class="bg-gray-100 p-10">
  <div class="max-w-7xl mx-auto bg-white p-6 rounded-2xl shadow">
    <h1 class="text-3xl font-bold mb-6">Sistema de Contaduría</h1>

    <form id="entryForm" class="grid grid-cols-1 md:grid-cols-8 gap-4 items-end">
      <input type="date" id="fecha" required class="border p-2 rounded">
      <input type="text" id="descripcion" placeholder="Descripción" required class="border p-2 rounded">
      <input type="text" id="rnc" placeholder="RNC" class="border p-2 rounded">
      <input type="text" id="ncf" placeholder="NCF" class="border p-2 rounded">
      <input type="number" id="totalBruto" placeholder="Total Bruto" step="0.01" class="border p-2 rounded">
      <input type="number" id="itbis" placeholder="% ITBIS" step="0.01" class="border p-2 rounded">
      <input type="text" id="nombre" placeholder="Nombre" class="border p-2 rounded">
      <button type="submit" class="bg-blue-500 text-white px-4 py-2 rounded">Agregar</button>
    </form>

    <div class="flex flex-wrap justify-between items-center my-4 gap-2">
      <div>
        <button onclick="exportarExcel()" class="bg-green-500 text-white px-4 py-2 rounded">Exportar a Excel</button>
        <button onclick="exportarPDF()" class="bg-red-500 text-white px-4 py-2 rounded">Exportar a PDF</button>
      </div>
      <input type="file" id="importarExcel" accept=".xlsx,.xls" class="border p-2 rounded">
    </div>

    <div class="mb-4">
      <input type="text" id="buscador" placeholder="Buscar en la tabla..." class="w-full border p-2 rounded">
    </div>

    <table class="w-full mt-6 table-auto border-collapse">
      <thead>
        <tr class="bg-gray-200">
          <th class="border px-2 py-1">Fecha</th>
          <th class="border px-2 py-1">Descripción</th>
          <th class="border px-2 py-1">RNC</th>
          <th class="border px-2 py-1">NCF</th>
          <th class="border px-2 py-1">ITBIS %</th>
          <th class="border px-2 py-1">Total Bruto</th>
          <th class="border px-2 py-1">ITBIS Monto</th>
          <th class="border px-2 py-1">Total Neto</th>
          <th class="border px-2 py-1">Nombre</th>
          <th class="border px-2 py-1">Acción</th>
        </tr>
      </thead>
      <tbody id="tablaBody"></tbody>
      <tfoot class="bg-gray-100">
        <tr>
          <td colspan="5" class="text-right font-bold px-2 py-2">Totales:</td>
          <td id="totalBrutoSum" class="border px-2 py-1 font-bold"></td>
          <td id="itbisSum" class="border px-2 py-1 font-bold"></td>
          <td id="totalNetoSum" class="border px-2 py-1 font-bold"></td>
          <td colspan="2"></td>
        </tr>
      </tfoot>
    </table>
  </div>

  <script>
    let datos = JSON.parse(localStorage.getItem('datosContables')) || [];

    const form = document.getElementById('entryForm');
    const tablaBody = document.getElementById('tablaBody');
    const inputFile = document.getElementById('importarExcel');
    const totalBrutoSum = document.getElementById('totalBrutoSum');
    const itbisSum = document.getElementById('itbisSum');
    const totalNetoSum = document.getElementById('totalNetoSum');
    const buscador = document.getElementById('buscador');

    form?.addEventListener('submit', function (e) {
      e.preventDefault();
      const fecha = document.getElementById('fecha').value;
      const descripcion = document.getElementById('descripcion').value;
      const rnc = document.getElementById('rnc').value;
      const ncf = document.getElementById('ncf').value;
      const totalBruto = parseFloat(document.getElementById('totalBruto').value) || 0;
      const itbisPorcentaje = parseFloat(document.getElementById('itbis').value) || 0;
      const itbisMonto = totalBruto * (itbisPorcentaje / 100);
      const totalNeto = totalBruto - itbisMonto;
      const nombre = document.getElementById('nombre').value;

      datos.push({ fecha, descripcion, rnc, ncf, itbisPorcentaje, totalBruto, itbisMonto, totalNeto, nombre });
      localStorage.setItem('datosContables', JSON.stringify(datos));
      actualizarTabla(buscador.value);
      form.reset();
    });

    function actualizarTabla(filtro = '') {
      tablaBody.innerHTML = '';
      let brutoTotal = 0, itbisTotal = 0, netoTotal = 0;

      datos.forEach((dato, index) => {
        const texto = Object.values(dato).join(' ').toLowerCase();
        if (texto.includes(filtro.toLowerCase())) {
          brutoTotal += parseFloat(dato.totalBruto);
          itbisTotal += parseFloat(dato.itbisMonto);
          netoTotal += parseFloat(dato.totalNeto);

          const fila = document.createElement('tr');
          fila.innerHTML = `
            <td class="border px-2 py-1">${dato.fecha}</td>
            <td class="border px-2 py-1">${dato.descripcion}</td>
            <td class="border px-2 py-1">${dato.rnc}</td>
            <td class="border px-2 py-1">${dato.ncf}</td>
            <td class="border px-2 py-1">${dato.itbisPorcentaje.toFixed(2)}%</td>
            <td class="border px-2 py-1">${dato.totalBruto.toFixed(2)}</td>
            <td class="border px-2 py-1">${dato.itbisMonto.toFixed(2)}</td>
            <td class="border px-2 py-1">${dato.totalNeto.toFixed(2)}</td>
            <td class="border px-2 py-1">${dato.nombre}</td>
            <td class="border px-2 py-1 text-center"><button onclick="eliminarDato(${index})" class="text-red-500 hover:underline">Eliminar</button></td>
          `;
          tablaBody.appendChild(fila);
        }
      });

      totalBrutoSum.textContent = brutoTotal.toFixed(2);
      itbisSum.textContent = itbisTotal.toFixed(2);
      totalNetoSum.textContent = netoTotal.toFixed(2);
    }

    buscador?.addEventListener('input', () => actualizarTabla(buscador.value));

    function eliminarDato(index) {
      if (confirm('¿Estás seguro de eliminar este dato?')) {
        datos.splice(index, 1);
        localStorage.setItem('datosContables', JSON.stringify(datos));
        actualizarTabla(buscador.value);
      }
    }

    function exportarExcel() {
      const ws = XLSX.utils.json_to_sheet(datos);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Datos");
      XLSX.writeFile(wb, "contaduria.xlsx");
    }

    function exportarPDF() {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();

      const headers = [["Fecha", "Descripción", "RNC", "NCF", "ITBIS %", "Total Bruto", "ITBIS Monto", "Total Neto", "Nombre"]];
      const rows = datos.map(dato => [
        dato.fecha,
        dato.descripcion,
        dato.rnc,
        dato.ncf,
        `${dato.itbisPorcentaje.toFixed(2)}%`,
        dato.totalBruto.toFixed(2),
        dato.itbisMonto.toFixed(2),
        dato.totalNeto.toFixed(2),
        dato.nombre
      ]);

      let totalBruto = 0, totalItbis = 0, totalNeto = 0;
      datos.forEach(dato => {
        totalBruto += parseFloat(dato.totalBruto);
        totalItbis += parseFloat(dato.itbisMonto);
        totalNeto += parseFloat(dato.totalNeto);
      });

      doc.text("Reporte de Contaduría", 14, 15);
      doc.autoTable({
        startY: 20,
        head: headers,
        body: rows,
        styles: { fontSize: 8 },
        theme: 'grid'
      });

      const finalY = doc.lastAutoTable.finalY || 30;

      doc.text("Totales:", 14, finalY + 10);
      doc.text(`Total Bruto: ${totalBruto.toFixed(2)}`, 14, finalY + 16);
      doc.text(`Total ITBIS: ${totalItbis.toFixed(2)}`, 14, finalY + 22);
      doc.text(`Total Neto: ${totalNeto.toFixed(2)}`, 14, finalY + 28);

      doc.save("reporte_contaduria.pdf");
    }

    inputFile?.addEventListener('change', function (e) {
      const file = e.target.files[0];
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const importedData = XLSX.utils.sheet_to_json(sheet);
        datos = importedData.map(d => {
          const totalBruto = parseFloat(d.totalBruto) || 0;
          const itbisPorcentaje = parseFloat(d.itbisPorcentaje) || 0;
          const itbisMonto = totalBruto * (itbisPorcentaje / 100);
          const totalNeto = totalBruto - itbisMonto;
          return {
            fecha: d.fecha || '',
            descripcion: d.descripcion || '',
            rnc: d.rnc || '',
            ncf: d.ncf || '',
            itbisPorcentaje,
            totalBruto,
            itbisMonto,
            totalNeto,
            nombre: d.nombre || ''
          };
        });
        localStorage.setItem('datosContables', JSON.stringify(datos));
        actualizarTabla(buscador.value);
      };

      reader.readAsArrayBuffer(file);
    });

    actualizarTabla();
  </script>
</body>
</html>
