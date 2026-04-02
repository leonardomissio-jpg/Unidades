async function processar() {
  const files = document.getElementById('files').files;
  const status = document.getElementById('status');

  status.innerText = "⏳ Processando...";

  let dados = [];

  for (let file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    let texto = "";

    for (let i = 1; i <= pdf.numPages; i++) {
      const pagina = await pdf.getPage(i);
      const conteudo = await pagina.getTextContent();
      texto += conteudo.items.map(i => i.str).join(" ") + " ";
    }

    // 🔍 EXTRAÇÕES MELHORADAS

    let valor = texto.match(/R\$ ?\d+[.,]\d+/i);

    let faturamento = texto.match(
      /(faturamento|emissão)[^\d]*(\d{2}\/\d{2}\/\d{4})/i
    );

    let consumo = texto.match(
      /(consumo|energia)[^\d]*(\d+[.,]?\d*)\s?(kWh|m³|m3)/i
    );

    let demandaContratada = texto.match(
      /(demanda contratada)[^\d]*(\d+[.,]?\d*)/i
    );

    let demandaConsumida = texto.match(
      /(demanda (medida|consumida|faturada))[^\d]*(\d+[.,]?\d*)/i
    );

    dados.push({
      arquivo: file.name,
      valor: valor ? valor[0] : "",
      faturamento: faturamento ? faturamento[2] : "",
      consumo: consumo ? consumo[2] + " " + consumo[3] : "",
      demanda_contratada: demandaContratada ? demandaContratada[2] : "",
      demanda_consumida: demandaConsumida ? demandaConsumida[3] : ""
    });
  }

  gerarExcel(dados);

  status.innerText = "✅ Concluído!";
}

function gerarExcel(dados) {
  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "Faturas");

  XLSX.writeFile(wb, "resultado.xlsx");
}
