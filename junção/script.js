
"use strict";

// ======================================
// 1. CONFIGURAÇÕES
// ======================================

const ARQUIVO = "vendas.xlsx";

let pedidos = [];
let filtrados = [];
let graficos = [];

const $ = id => document.getElementById(id);

const dinheiro = valor =>
    Number(valor || 0).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL"
    });

// ======================================
// 2. INICIALIZAÇÃO
// ======================================

document.addEventListener("DOMContentLoaded", () => {

    console.log("SCRIPT.JS ESTÁ FUNCIONANDO!");

    $("importar").addEventListener("click", () => {
        $("arquivo").click();
    });

    $("arquivo").addEventListener("change", async event => {

        const arquivo = event.target.files[0];

        if (!arquivo) return;

        try {
            const buffer = await arquivo.arrayBuffer();
            lerExcel(buffer, arquivo.name);
        } catch (erro) {
            exibirErro(erro);
        }

        event.target.value = "";
    });

    $("atualizar").addEventListener(
        "click",
        carregarAutomaticamente
    );

    ["regiao", "produto", "estado"].forEach(id => {
        $(id).addEventListener("change", atualizar);
    });

    $("pesquisa").addEventListener("input", atualizar);

    $("limpar").addEventListener("click", () => {

        ["regiao", "produto", "estado", "pesquisa"]
            .forEach(id => $(id).value = "");

        atualizar();
    });

    $("exportar").addEventListener("click", exportarCSV);

    // Carrega a planilha automaticamente.
    carregarAutomaticamente();
});

// ======================================
// 3. MENSAGENS
// ======================================

function aviso(mensagem, erro = false) {

    $("aviso").textContent = mensagem;
    $("aviso").classList.toggle("erro", erro);
}

function exibirErro(erro) {

    console.error("Erro no Excel:", erro);

    aviso("Erro: " + erro.message, true);

    $("status").textContent = "Falha no carregamento";
}

// ======================================
// 4. CONEXÃO AUTOMÁTICA COM EXCEL
// ======================================

async function carregarAutomaticamente() {

    aviso("Conectando à planilha Excel...");

    $("status").textContent = "Carregando Excel...";

    if (typeof XLSX === "undefined") {

        aviso(
            "A biblioteca XLSX não carregou. Verifique sua internet.",
            true
        );

        $("status").textContent = "Biblioteca indisponível";

        return;
    }

    try {

        // Localiza o Excel na mesma pasta do site.
         const endereco = new URL(
        "./vendas.xlsx",
        document.baseURI
        );

        // Evita carregar uma versão antiga em cache.
        endereco.searchParams.set(
            "atualizacao",
            Date.now()
        );

        console.log(
            "Procurando planilha em:",
            endereco.href
        );

        const resposta = await fetch(endereco.href, {
            cache: "no-store"
        });

        console.log(
            "Resposta do servidor:",
            resposta.status
        );

        if (!resposta.ok) {

            throw new Error(
                "Não foi possível encontrar vendas.xlsx. " +
                "Código HTTP: " + resposta.status
            );
        }

        const buffer = await resposta.arrayBuffer();

        lerExcel(buffer, ARQUIVO);

    } catch (erro) {

        console.error(
            "Falha no carregamento automático:",
            erro
        );

        $("status").textContent = "Excel não conectado";

        aviso(
            "Não foi possível carregar o Excel automaticamente. " +
            "Abra o index.html pelo Live Server. " +
            "Se necessário, clique em Importar Excel. " +
            "Detalhes: " + erro.message,
            true
        );
    }
}

// ======================================
// 5. CONVERSÃO DE VALORES
// ======================================

function numero(valor) {

    if (typeof valor === "number") {
        return Number.isFinite(valor) ? valor : 0;
    }

    let texto = String(valor ?? "")
        .replace(/R\$/gi, "")
        .replace(/\s/g, "")
        .trim();

    if (texto.includes(",")) {

        texto = texto
            .replace(/\./g, "")
            .replace(",", ".");
    }

    const resultado = Number(texto);

    return Number.isFinite(resultado) ? resultado : 0;
}

function texto(valor) {
    return String(valor ?? "").trim();
}

function chave(valor) {

    return texto(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .toLowerCase();
}

function dataExcel(valor) {

    if (valor instanceof Date) {

        return valor.toLocaleDateString("pt-BR", {
            timeZone: "UTC"
        });
    }

    if (typeof valor === "number") {

        const data = XLSX.SSF.parse_date_code(valor);

        if (data) {

            return [
                String(data.d).padStart(2, "0"),
                String(data.m).padStart(2, "0"),
                data.y
            ].join("/");
        }
    }

    return texto(valor);
}

// ======================================
// 6. LEITURA DAS CÉLULAS
// ======================================

function valorCelula(aba, endereco) {

    const celula = aba[endereco];

    if (!celula) return "";

    return celula.v ?? "";
}

function tabelaAuxiliar(
    aba,
    colunaNome,
    colunaValor,
    primeira,
    ultima
) {

    const resultado = new Map();

    for (let linha = primeira; linha <= ultima; linha++) {

        const nome = valorCelula(
            aba,
            colunaNome + linha
        );

        const valor = valorCelula(
            aba,
            colunaValor + linha
        );

        if (texto(nome)) {

            resultado.set(
                chave(nome),
                numero(valor)
            );
        }
    }

    return resultado;
}

// ======================================
// 7. PROCESSAMENTO DA PLANILHA
// ======================================

function lerExcel(buffer, nomeArquivo) {

    if (typeof XLSX === "undefined") {

        throw new Error(
            "Biblioteca XLSX indisponível."
        );
    }

    const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: false
    });

    console.log(
        "Abas encontradas:",
        workbook.SheetNames
    );

    const aba = workbook.SheetNames
        .map(nome => workbook.Sheets[nome])
        .find(planilha =>
            chave(valorCelula(planilha, "A1")) === "pedido"
        );

    if (!aba) {

        throw new Error(
            "Não encontrei uma aba com Pedido na célula A1."
        );
    }

    // Tabela de preços: L2:M8.
    const precos = tabelaAuxiliar(
        aba,
        "L",
        "M",
        2,
        8
    );

    // Tabela de fretes: O2:P6.
    const fretes = tabelaAuxiliar(
        aba,
        "O",
        "P",
        2,
        6
    );

    console.log("Preços:", [...precos]);
    console.log("Fretes:", [...fretes]);

    const novosPedidos = [];
    const problemas = [];

    const intervalo = XLSX.utils.decode_range(
        aba["!ref"]
    );

    for (
        let linha = 2;
        linha <= intervalo.e.r + 1;
        linha++
    ) {

        const codigo = valorCelula(
            aba,
            "A" + linha
        );

        if (!/^\d+$/.test(texto(codigo))) {
            continue;
        }

        const regiao = texto(
            valorCelula(aba, "B" + linha)
        );

        const estado = texto(
            valorCelula(aba, "C" + linha)
        );

        const produto = texto(
            valorCelula(aba, "D" + linha)
        );

        const data = dataExcel(
            valorCelula(aba, "E" + linha)
        );

        const quantidade = numero(
            valorCelula(aba, "F" + linha)
        );

        // Consulta o preço na tabela auxiliar.
        let valorUnitario = precos.get(
            chave(produto)
        );

        // Alternativa: valor salvo na coluna G.
        if (valorUnitario === undefined) {

            const salvo = valorCelula(
                aba,
                "G" + linha
            );

            if (numero(salvo) > 0) {

                valorUnitario = numero(salvo);

            } else {

                problemas.push(
                    "Preço não encontrado: " + produto
                );

                continue;
            }
        }

        // Consulta o frete na tabela auxiliar.
        let frete = fretes.get(
            chave(regiao)
        );

        // Alternativa: valor salvo na coluna H.
        if (frete === undefined) {

            const salvo = valorCelula(
                aba,
                "H" + linha
            );

            if (numero(salvo) > 0) {

                frete = numero(salvo);

            } else {

                problemas.push(
                    "Frete não encontrado: " + regiao
                );

                continue;
            }
        }

        // Calcula o total do pedido.
        const total =
            quantidade * valorUnitario + frete;

        novosPedidos.push({
            codigo: texto(codigo),
            regiao,
            estado,
            produto,
            data,
            quantidade,
            valorUnitario,
            frete,
            total
        });
    }

    if (!novosPedidos.length) {

        throw new Error(
            "Nenhum pedido válido foi encontrado."
        );
    }

    pedidos = novosPedidos;

    preencherFiltros();
    atualizar();

    const faturamento = pedidos.reduce(
        (soma, pedido) => soma + pedido.total,
        0
    );

    $("status").textContent =
        nomeArquivo + " • " +
        pedidos.length + " pedidos";

    aviso(
        "Excel conectado com sucesso! " +
        pedidos.length + " pedidos carregados. " +
        "Faturamento: " + dinheiro(faturamento) +
        (problemas.length
            ? ". Registros ignorados: " + problemas.length
            : "")
    );

    console.table(pedidos);

    if (problemas.length) {
        console.warn(problemas);
    }
}

// ======================================
// 8. FILTROS
// ======================================

function preencherFiltros() {

    const configuracoes = [
        ["regiao", "regiao"],
        ["produto", "produto"],
        ["estado", "estado"]
    ];

    configuracoes.forEach(([id, campo]) => {

        const select = $(id);
        const anterior = select.value;

        select.length = 1;

        const valores = [...new Set(
            pedidos.map(p => p[campo])
        )].sort();

        valores.forEach(valor => {

            const opcao = document.createElement("option");

            opcao.value = valor;
            opcao.textContent = valor;

            select.appendChild(opcao);
        });

        if (valores.includes(anterior)) {
            select.value = anterior;
        }
    });
}

function atualizar() {

    const regiao = $("regiao").value;
    const produto = $("produto").value;
    const estado = $("estado").value;

    const pesquisa = chave(
        $("pesquisa").value
    );

    filtrados = pedidos.filter(p => {

        const conteudo = chave(
            `${p.codigo} ${p.regiao} ${p.estado} ` +
            `${p.produto} ${p.data}`
        );

        return (
            (!regiao || p.regiao === regiao) &&
            (!produto || p.produto === produto) &&
            (!estado || p.estado === estado) &&
            conteudo.includes(pesquisa)
        );
    });

    atualizarIndicadores();
    atualizarTabela();
    atualizarGraficos();
}

// ======================================
// 9. INDICADORES
// ======================================

function soma(campo) {

    return filtrados.reduce(
        (total, pedido) => total + pedido[campo],
        0
    );
}

function atualizarIndicadores() {

    const faturamento = soma("total");
    const quantidade = soma("quantidade");
    const total = filtrados.length;

    $("faturamento").textContent =
        dinheiro(faturamento);

    $("totalPedidos").textContent = total;

    $("quantidade").textContent = quantidade;

    $("ticket").textContent = dinheiro(
        total ? faturamento / total : 0
    );
}

// ======================================
// 10. TABELA DE PEDIDOS
// ======================================

function atualizarTabela() {

    const tbody = $("tabela");

    tbody.replaceChildren();

    for (const pedido of filtrados) {

        const tr = document.createElement("tr");

        const valores = [
            pedido.codigo,
            pedido.regiao,
            pedido.estado,
            pedido.produto,
            pedido.data,
            pedido.quantidade,
            dinheiro(pedido.valorUnitario),
            dinheiro(pedido.frete),
            dinheiro(pedido.total)
        ];

        valores.forEach(valor => {

            const td = document.createElement("td");

            td.textContent = valor;

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    }

    if (!filtrados.length) {

        const tr = document.createElement("tr");
        const td = document.createElement("td");

        td.colSpan = 9;
        td.textContent = "Nenhum pedido encontrado.";

        tr.appendChild(td);
        tbody.appendChild(tr);
    }

    $("registros").textContent =
        filtrados.length + " registros";

    $("totalTabela").textContent =
        dinheiro(soma("total"));
}

// ======================================
// 11. AGRUPAMENTO DOS DADOS
// ======================================

function agrupar(campo, valor) {

    const resultado = {};

    filtrados.forEach(pedido => {

        const nome = pedido[campo];

        resultado[nome] =
            (resultado[nome] || 0) +
            pedido[valor];
    });

    return resultado;
}

// ======================================
// 12. CRIAÇÃO DOS GRÁFICOS
// ======================================

function criarGrafico(
    id,
    tipo,
    labels,
    valores,
    cor
) {

    return new Chart($(id), {

        type: tipo,

        data: {
            labels,

            datasets: [{
                label: "Vendas",
                data: valores,
                backgroundColor: cor,
                borderColor: "#176b4d",
                borderWidth: tipo === "line" ? 3 : 1,
                borderRadius: tipo === "bar" ? 6 : 0,
                tension: 0.35,
                fill: tipo === "line"
            }]
        },

        options: {
            responsive: true,
            maintainAspectRatio: false,

            plugins: {
                legend: {
                    display: tipo === "doughnut",
                    position: "bottom"
                }
            },

            scales: tipo === "doughnut"
                ? {}
                : {
                    y: {
                        beginAtZero: true
                    }
                }
        }
    });
}

// ======================================
// 13. ATUALIZAÇÃO DOS GRÁFICOS
// ======================================

function atualizarGraficos() {

    if (typeof Chart === "undefined") {

        console.warn(
            "A biblioteca Chart.js não carregou."
        );

        return;
    }

    graficos.forEach(grafico => {
        grafico.destroy();
    });

    graficos = [];

    // Faturamento por região.
    const regioes = agrupar(
        "regiao",
        "total"
    );

    graficos.push(criarGrafico(
        "graficoRegiao",
        "doughnut",
        Object.keys(regioes),
        Object.values(regioes),
        [
            "#176b4d",
            "#4b8bd8",
            "#a277d4",
            "#f2ae54",
            "#46b8a0"
        ]
    ));

    // Quantidade vendida por produto.
    const produtos = agrupar(
        "produto",
        "quantidade"
    );

    graficos.push(criarGrafico(
        "graficoProduto",
        "bar",
        Object.keys(produtos),
        Object.values(produtos),
        "#23845e"
    ));

    // Evolução por data.
    const datas = agrupar(
        "data",
        "total"
    );

    const ordenadas = Object.keys(datas).sort((a, b) => {

        const converter = data => {

            const [dia, mes, ano] = data.split("/");

            return `${ano}-${mes}-${dia}`;
        };

        return converter(a).localeCompare(converter(b));
    });

    graficos.push(criarGrafico(
        "graficoData",
        "line",
        ordenadas,
        ordenadas.map(data => datas[data]),
        "rgba(23,107,77,0.12)"
    ));
}

// ======================================
// 14. EXPORTAR CSV
// ======================================

function exportarCSV() {

    if (!filtrados.length) {

        aviso(
            "Não existem pedidos para exportar.",
            true
        );

        return;
    }

    const linhas = [
        [
            "Pedido",
            "Região",
            "Estado",
            "Produtos",
            "Data",
            "Quantidade",
            "Valor Unitário",
            "Frete",
            "Total"
        ],

        ...filtrados.map(p => [
            p.codigo,
            p.regiao,
            p.estado,
            p.produto,
            p.data,
            p.quantidade,
            p.valorUnitario,
            p.frete,
            p.total
        ])
    ];

    const aba = XLSX.utils.aoa_to_sheet(linhas);

    const csv = XLSX.utils.sheet_to_csv(aba, {
        FS: ";"
    });

    const blob = new Blob(
        ["\uFEFF" + csv],
        {
            type: "text/csv;charset=utf-8"
        }
    );

    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");

    link.href = url;
    link.download = "relatorio-vendas.csv";

    document.body.appendChild(link);

    link.click();
    link.remove();

    setTimeout(() => {
        URL.revokeObjectURL(url);
    }, 1000);
}
