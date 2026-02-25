if (sessionStorage.getItem("logado") !== "true") {
  window.location.href = "login.html";
}

let chart = null;
let relatorioAtual = null;

function logout() {
  sessionStorage.clear();
  window.location.href = "login.html";
}

function salvarRelatorio() {
  const r = {
    meta: Number(meta.value),
    faturamento: Number(faturamento.value),
    pec: pec.value,
    ipc: ipc.value,
    tkm: tkm.value,
    supra: supra.value
  };

  r.percentual = (r.faturamento / r.meta) * 100;
  relatorioAtual = r;

  cardPrincipal.innerHTML = `
    <h2>${r.percentual.toFixed(1)}%</h2>
    Meta: R$ ${r.meta}<br>
    Faturamento: R$ ${r.faturamento}
  `;

  dPEC.innerText = r.pec;
  dIPC.innerText = r.ipc;
  dTKM.innerText = r.tkm;
  dSUPRA.innerText = r.supra;

  if (chart) chart.destroy();

  chart = new Chart(document.getElementById("grafico"), {
    type: "bar",
    data: {
      labels: ["Meta", "Realizado"],
      datasets: [{
        data: [r.meta, r.faturamento],
        backgroundColor: ["#64748b", "#22c55e"]
      }]
    },
    options: {
      plugins: { legend: { display: false } }
    }
  });
}

function gerarPPT() {
  if (!relatorioAtual) {
    alert("Salve um relatório antes de gerar o PowerPoint.");
    return;
  }

  const ppt = new PptxGenJS();

  // 🎨 PALETA EXECUTIVA PREMIUM
  const CORES = {
    fundo: "0B1220",     // Azul marinho profundo
    destaque: "D4AF37",  // Dourado executivo
    verde: "16A34A",
    vermelho: "B91C1C",
    cinza: "94A3B8",
    branco: "F8FAFC"
  };

  const corResultado =
    relatorioAtual.percentual >= 100 ? CORES.verde : CORES.vermelho;

  /* ======================================================
     SLIDE 1 – CAPA EXECUTIVA
  ====================================================== */
  let slide = ppt.addSlide();
  slide.background = { fill: CORES.fundo };

  slide.addText(
    `RELATÓRIO DE RESULTADOS – ${mes.value}/${ano.value}`,
    {
      x: 1,
      y: 2.6,
      w: 8,
      align: "center",
      fontSize: 30,
      bold: true,
      color: CORES.branco
    }
  );

  slide.addText(
    "Apresentação Executiva",
    {
      x: 1,
      y: 3.6,
      w: 8,
      align: "center",
      fontSize: 16,
      color: CORES.destaque
    }
  );

  slide.addText(
    "Reunião de Diretoria",
    {
      x: 0.6,
      y: 6.9,
      fontSize: 9,
      color: CORES.cinza
    }
  );

  /* ======================================================
     SLIDE 2 – VISÃO GERAL (DASHBOARD EXECUTIVO)
  ====================================================== */
  slide = ppt.addSlide();
  slide.background = { fill: CORES.fundo };

  slide.addText("Visão Geral do Mês", {
    x: 0.6,
    y: 0.5,
    fontSize: 20,
    bold: true,
    color: CORES.branco
  });

  const cards = [
    { titulo: "Resultado do Mês", valor: `${relatorioAtual.percentual.toFixed(1)}%` },
    { titulo: "PEC", valor: relatorioAtual.pec },
    { titulo: "IPC", valor: relatorioAtual.ipc },
    { titulo: "TKM", valor: relatorioAtual.tkm },
    { titulo: "VENDAS DE SUPRACORP", valor: relatorioAtual.supra }
  ];

  cards.forEach((c, i) => {
    const xBase = 0.5 + i * 1.85;

    slide.addShape(ppt.ShapeType.rect, {
      x: xBase,
      y: 1.7,
      w: 1.6,
      h: 1.5,
      fill: { color: "020617" },
      line: { color: CORES.destaque, width: 0.4 }
    });

    slide.addText(c.titulo, {
      x: xBase + 0.1,
      y: 1.85,
      w: 1.4,
      fontSize: 9,
      align: "center",
      color: CORES.cinza
    });

    slide.addText(c.valor, {
      x: xBase + 0.1,
      y: 2.4,
      w: 1.4,
      fontSize: 20,
      bold: true,
      align: "center",
      color: corResultado
    });
  });

  /* ======================================================
     SLIDE 3 – RESULTADO DO MÊS
  ====================================================== */
  slide = ppt.addSlide();
  slide.background = { fill: CORES.fundo };

  slide.addText("Resultado do Mês – Meta vs Realizado", {
    x: 0.6,
    y: 0.5,
    fontSize: 20,
    bold: true,
    color: CORES.branco
  });

  slide.addChart(
    ppt.ChartType.bar,
    [
      { name: "Meta", labels: ["Meta"], values: [relatorioAtual.meta] },
      { name: "Realizado", labels: ["Realizado"], values: [relatorioAtual.faturamento] }
    ],
    {
      x: 0.6,
      y: 1.7,
      w: 4.6,
      h: 3.2,
      barDir: "col",
      showLegend: false,
      chartColors: [CORES.cinza, CORES.verde]
    }
  );

  slide.addText(`${relatorioAtual.percentual.toFixed(1)}%`, {
    x: 5.4,
    y: 2.2,
    fontSize: 40,
    bold: true,
    color: corResultado
  });

  slide.addText(
    "Desempenho sustentado por evolução do ticket médio, maior eficiência comercial e fortalecimento das categorias estratégicas.",
    {
      x: 5.4,
      y: 3.4,
      w: 4,
      fontSize: 13,
      color: CORES.cinza
    }
  );

  /* ======================================================
     SLIDE 4 – INDICADORES OPERACIONAIS
  ====================================================== */
  slide = ppt.addSlide();
  slide.background = { fill: CORES.fundo };

  slide.addText("Indicadores Operacionais", {
    x: 0.6,
    y: 0.5,
    fontSize: 20,
    bold: true,
    color: CORES.branco
  });

  const indicadores = [
    { nome: "PEC", valor: relatorioAtual.pec, desc: "Produtividade por colaborador" },
    { nome: "IPC", valor: relatorioAtual.ipc, desc: "Itens médios por cupom" },
    { nome: "TKM", valor: relatorioAtual.tkm, desc: "Valor médio por transação" }
  ];

  indicadores.forEach((i, idx) => {
    slide.addText(i.nome, {
      x: 0.8,
      y: 1.8 + idx * 1.6,
      fontSize: 13,
      bold: true,
      color: CORES.destaque
    });

    slide.addText(i.valor, {
      x: 2.3,
      y: 1.75 + idx * 1.6,
      fontSize: 26,
      bold: true,
      color: CORES.branco
    });

    slide.addText(i.desc, {
      x: 4.3,
      y: 1.85 + idx * 1.6,
      fontSize: 13,
      color: CORES.cinza
    });
  });

  /* ======================================================
     SLIDE 5 – MARCA ESTRATÉGICA
  ====================================================== */
  slide = ppt.addSlide();
  slide.background = { fill: CORES.fundo };

  slide.addText("Marca Estratégica – Análise Executiva", {
    x: 0.6,
    y: 0.5,
    fontSize: 20,
    bold: true,
    color: CORES.branco
  });

  slide.addText(
    "A marca estratégica apresenta trajetória consistente de crescimento, com oportunidades claras de expansão por meio de ações comerciais direcionadas, capacitação da equipe e aumento de visibilidade no ponto de venda.",
    {
      x: 0.6,
      y: 2,
      w: 9,
      fontSize: 14,
      color: CORES.cinza
    }
  );

  ppt.writeFile(`Relatorio_Executivo_${mes.value}_${ano.value}.pptx`);
}