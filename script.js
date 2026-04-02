async function processar() {
  const status = document.getElementById('status');
  const files = document.getElementById('files').files;

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

    console.log("TEXTO COMPLETO:", texto);

    // 🔥 NORMALIZA TEXTO
    texto = texto.replace(/\s+/g, " ");

    // =========================
    // 🎯 EXTRAÇÃO INTELIGENTE
    // =========================

    // 💰 VALOR TOTAL (último TOTAL)
    let valores = [...texto.matchAll(/TOTAL\s+([\d.,]+)/gi)];
    let valorTotal = valores.length ? valores[valores.length - 1][1] : "";

    // 🔌 UC (Instalação)
    let uc = texto.match(/Instalação nº\s*(\d+)/i);

    // 📅 Data faturamento
    let data = texto.match(/Data do Documento\s*(\d{2}\/\d{2}\/\d{4})/i);

    // ⚡ CONSUMO REAL (pega tabela correta)
    let consumoTabela = texto.match(/FEV\/\d{2}\s+[\d.,]+\s+([\d.,]+)\s+([\d.,]+)/);

    let consumoPonta = consumoTabela ? consumoTabela[1] : "";
    let consumoForaPonta = consumoTabela ? consumoTabela[2] : "";

    // 📈 Demanda contratada (fixa no topo)
    let demContratada = texto.match(/Demanda - KW\s*([\d.,]+)/i);

    // 📉 Demanda medida REAL (linha correta)
    let demMedida = texto.match(/DEMANDA KW\s+([\d.,]+)/i);

    dados.push({
      arquivo: file.name,
      valor_total: valorTotal,
      unidade_consumidora: uc ? uc[1] : "",
      data_faturamento: data ? data[1] : "",
      consumo_ponta: consumoPonta,
      consumo_fora_ponta: consumoForaPonta,
      demanda_contratada: demContratada ? demContratada[1] : "",
      demanda_medida: demMedida ? demMedida[1] : ""
    });
  }

  gerarExcel(dados);

  status.innerText = "✅ EXTRAÇÃO COMPLETA!";
}

function gerarExcel(dados) {
  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Faturas");
  XLSX.writeFile(wb, "resultado.xlsx");
}
