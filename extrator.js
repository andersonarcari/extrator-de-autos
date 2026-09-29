/*
 * Extrator de Autos — BPM MAmb — GAIA Fiscalização (Ecosistemas/SEMAD-MG)
 * Versão para navegador (roda dentro da aba do Ecosistemas já logada).
 *
 * Mesma lógica do robô em Python (fase1_busca.py / fase2_extracao.py / report.py):
 *  - Fase 1: Atos do período por município; descarta "Cancelado"; filtra Unidade
 *    Responsável com "MAmb"; se vier vazia na listagem, faz 2ª verificação no
 *    detalhe do Ato; se continuar vazia → aba "Revisão manual".
 *  - Fase 2: por Ato aceito, descarta Cancelado (2ª camada), lista Autos,
 *    decide pela unidade do 1º Auto, extrai todos os Autos; município vem do ATO.
 *  - Excel com abas: Relatório | Revisão manual | Municípios com problema.
 *
 * Nenhuma senha, token ou dado sai do navegador do usuário.
 */
(function () {
  "use strict";

  var VERSAO = "0.3.1";
  var BRASAO = "data:image/png;base64,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";
  var FILTRO_UNIDADE = "MAmb";
  var STATUS_CANCELADO = "cancel";
  var CONCORRENCIA_FASE2 = 4;
  var URL_SHEETJS = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";

  if (location.hostname !== "ecosistemas.meioambiente.mg.gov.br") {
    alert("O Extrator de Autos só funciona dentro da aba do Ecosistemas.\n\nVá para a aba do GAIA Fiscalização (já logada) e clique no favorito com ela aberta.");
    return;
  }
  // Painel já aberto nesta aba: se for a mesma versão, só mostra de novo;
  // se for uma versão antiga, remove e abre a nova.
  if (window.__gaiaRelatorio) {
    if (window.__gaiaRelatorio.versao === VERSAO) { window.__gaiaRelatorio.mostrar(); return; }
    var antigo = document.getElementById("gaia-relatorio-painel");
    if (antigo) antigo.remove();
    window.__gaiaRelatorio = null;
  }

  // ------------------------------------------------------------------ SRAIs
  var SRAIS_TXT = {
    "SRAI Divinópolis": "Abaeté;Arcos;Araújos;Bambuí;Biquinhas;Bom Despacho;Camacho;Cedro do Abaeté;Cláudio;Conceição do Pará;Córrego Danta;Córrego Fundo;Divinópolis;Dores do Indaiá;Estrela do Indaiá;Formiga;Igaratinga;Iguatama;Itatiaiuçu;Itaúna;Japaraíba;Lagoa da Prata;Leandro Ferreira;Luz;Maravilhas;Martinho Campos;Medeiros;Moema;Morada Nova de Minas;Nova Serrana;Paineiras;Pains;Papagaios;Pará de Minas;Pedra do Indaiá;Pequi;Pimenta;Pitangui;Pompéu;Quartel Geral;Santo Antônio do Monte;São Gonçalo do Pará;São José da Varginha;São Sebastião do Oeste;Serra da Saudade;Tapiraí",
    "SRAI Poços de Caldas": "Alfenas;Alpinópolis;Alterosa;Andradas;Arceburgo;Areado;Bandeira do Sul;Bom Jesus da Penha;Botelhos;Cabo Verde;Caldas;Campestre;Campo do Meio;Campos Gerais;Capetinga;Capitólio;Carmo do Rio Claro;Carvalhópolis;Cássia;Claraval;Conceição da Aparecida;Delfinópolis;Divisa Nova;Doresópolis;Fama;Fortaleza de Minas;Guaranésia;Guaxupé;Ibiraci;Ibitiúra de Minas;Ipuiúna;Itamogi;Itaú de Minas;Jacuí;Juruaia;Machado;Monte Belo;Monte Santo de Minas;Muzambinho;Nova Resende;Paraguaçu;Passos;Piumhi;Poço Fundo;Poços de Caldas;Pratápolis;Santa Rita de Caldas;São João Batista do Glória;São José da Barra;São Pedro da União;São Roque de Minas;São Sebastião do Paraíso;São Tomás de Aquino;Serrania;Vargem Bonita",
    "SRAI Montes Claros": "Água Boa;Alvorada de Minas;Angelândia;Aricanduva;Augusto de Lima;Berizal;Berilo;Bocaiúva;Bonito de Minas;Botumirim;Brasília de Minas;Buenópolis;Buritizeiro;Campo Azul;Capitão Enéas;Capelinha;Carbonita;Catuti;Chapada do Norte;Claro dos Poções;Cônego Marinho;Conceição do Mato Dentro;Congonhas do Norte;Coração de Jesus;Corinto;Couto de Magalhães de Minas;Cristália;Curral de Dentro;Curvelo;Datas;Diamantina;Dom Joaquim;Engenheiro Navarro;Espinosa;Felício dos Santos;Felixlândia;Francisco Badaró;Francisco Dumont;Francisco Sá;Fruta de Leite;Gameleiras;Glaucilândia;Gouveia;Grão Mogol;Guaraciama;Ibiaí;Ibiracatu;Icaraí de Minas;Indaiabira;Inimutaba;Itacambira;Itacarambi;Itamarandiba;Jaíba;Janaúba;Januária;Japonvar;Jenipapo de Minas;Jequitaí;Joaquim Felício;José Gonçalves de Minas;Josenópolis;Juramento;Juvenília;Lagoa dos Patos;Lassance;Leme do Prado;Lontra;Luislândia;Mamonas;Manga;Matias Cardoso;Mato Verde;Minas Novas;Mirabela;Miravânia;Montalvânia;Monte Azul;Montes Claros;Monjolos;Montezuma;Morro da Garça;Morro do Pilar;Ninheira;Nova Porteirinha;Novorizonte;Olhos-dÁgua;Padre Carvalho;Pai Pedro;Patis;Pedras de Maria da Cruz;Pintópolis;Pirapora;Ponto Chique;Porteirinha;Presidente Juscelino;Presidente Kubitschek;Riacho dos Machados;Rio Pardo de Minas;Rubelita;Salinas;Santa Cruz de Salinas;Santa Fé de Minas;Santo Antônio do Itambé;Santo Antônio do Retiro;Santo Antônio do Rio Abaixo;Santo Hipólito;São Francisco;São Gonçalo do Rio Preto;São João da Lagoa;São João da Ponte;São João das Missões;São João do Pacuí;São João do Paraíso;São Romão;São Sebastião do Rio Preto;Senador Modestino Gonçalves;Serra Azul de Minas;Serranópolis de Minas;Serro;Taiobeiras;Três Marias;Turmalina;Ubaí;Vargem Grande do Rio Pardo;Várzea da Palma;Varzelândia;Veredinha;Verdelândia",
    "SRAI Governador Valadares": "Aimorés;Alpercata;Alvarenga;Cantagalo;Capitão Andrade;Central de Minas;Coluna;Conselheiro Pena;Coroaci;Cuparaque;Divino das Laranjeiras;Divinolândia de Minas;Dores de Guanhães;Engenheiro Caldas;Fernandes Tourinho;Frei Inocêncio;Frei Lagonegro;Galiléia;Goiabeira;Gonzaga;Governador Valadares;Guanhães;Itabirinha;Itanhomi;Itueta;José Raydan;Mantena;Marilac;Materlândia;Mathias Lobato;Mendes Pimentel;Nacip Raydan;Nova Belém;Paulistas;Peçanha;Periquito;Resplendor;Rio Vermelho;Sabinópolis;Santa Efigênia de Minas;Santa Maria do Suaçuí;Santa Rita do Itueto;São Félix de Minas;São Geraldo da Piedade;São Geraldo do Baixio;São João do Manteninha;São João Evangelista;São José da Safira;São José do Jacuri;São Pedro do Suaçuí;São Sebastião do Maranhão;Sardoá;Senhora do Porto;Sobrália;Tarumirim;Tumiritinga;Virginópolis;Virgolândia",
    "SRAI Teófilo Otoni": "Águas Formosas;Águas Vermelhas;Almenara;Araçuaí;Ataléia;Bandeira;Bertópolis;Campanário;Caraí;Carlos Chagas;Catuji;Cachoeira de Pajeú;Comercinho;Coronel Murta;Crisólita;Divisa Alegre;Divisópolis;Felisburgo;Frei Gaspar;Fronteira dos Vales;Itaipé;Itaobim;Itinga;Itambacuri;Jacinto;Jampruca;Jequitinhonha;Joaíma;Jordânia;Ladainha;Machacalis;Malacacheta;Mata Verde;Medina;Monte Formoso;Nanuque;Nova Módica;Novo Cruzeiro;Novo Oriente de Minas;Ouro Verde de Minas;Padre Paraíso;Palmópolis;Pavão;Pedra Azul;Pescador;Ponto dos Volantes;Poté;Rio do Prado;Rubim;Salto da Divisa;Santa Helena de Minas;Santa Maria do Salto;Santo Antônio do Jacinto;São José do Divino;Serra dos Aimorés;Setubinha;Teófilo Otoni;Umburatiba;Virgem da Lapa",
    "SRAI Juiz de Fora": "Além Paraíba;Antônio Prado de Minas;Araponga;Argirita;Astolfo Dutra;Barão de Monte Alto;Belmiro Braga;Bicas;Brás Pires;Cajuri;Canaã;Carangola;Cataguases;Chácara;Chiador;Coimbra;Coronel Pacheco;Descoberto;Divinésia;Divino;Dona Eusébia;Dores do Turvo;Ervália;Estrela Dalva;Eugenópolis;Faria Lemos;Fervedouro;Goianá;Guarani;Guarará;Guidoval;Guiricema;Itamarati de Minas;Juiz de Fora;Laranjal;Leopoldina;Lima Duarte;Mar de Espanha;Maripá de Minas;Matias Barbosa;Mercês;Miradouro;Miraí;Muriaé;Olaria;Orizânia;Palma;Patrocínio do Muriaé;Paula Cândido;Pedra do Anta;Pedra Dourada;Pedro Teixeira;Piau;Piraúba;Pirapetinga;Recreio;Rio Novo;Rio Pomba;Rio Preto;Rochedo de Minas;Rodeiro;Rosário da Limeira;Santa Bárbara do Monte Verde;Santa Rita de Jacutinga;Santana de Cataguases;Santana do Deserto;Santo Antônio do Aventureiro;São Francisco do Glória;São Geraldo;São João Nepomuceno;São Miguel do Anta;São Sebastião da Vargem Alegre;Senador Cortes;Senador Firmino;Silveirânia;Simão Pereira;Tabuleiro;Teixeiras;Tocantins;Tombos;Ubá;Vieiras;Viçosa;Visconde do Rio Branco;Volta Grande",
    "SRAI Ipatinga": "Abre Campo;Acaiaca;Alto Caparaó;Alto Jequitibá;Alvinópolis;Amparo do Serra;Antônio Dias;Barão de Cocais;Barra Longa;Bela Vista de Minas;Belo Oriente;Bom Jesus do Amparo;Bom Jesus do Galho;Braúnas;Bugre;Caiana;Caparaó;Caputira;Caratinga;Carmésia;Catas Altas;Chalé;Conceição de Ipanema;Coronel Fabriciano;Córrego Novo;Dionísio;Dom Cavati;Dom Silvério;Durandé;Entre Folhas;Espera Feliz;Ferros;Guaraciaba;Iapu;Imbé de Minas;Inhapim;Ipaba;Ipanema;Ipatinga;Itabira;Itambé do Mato Dentro;Jaguaraçu;Jequeri;Joanésia;João Monlevade;Lajinha;Luisburgo;Manhuaçu;Manhumirim;Marliéria;Martins Soares;Matipó;Mesquita;Mutum;Naque;Nova Era;Oratórios;Passabém;Pedra Bonita;Piedade de Caratinga;Piedade de Ponte Nova;Pingo-dÁgua;Pocrane;Ponte Nova;Raul Soares;Reduto;Rio Casca;Rio Doce;Rio Piracicaba;Santa Bárbara;Santa Bárbara do Leste;Santa Cruz do Escalvado;Santa Margarida;Santa Maria de Itabira;Santa Rita de Minas;Santana do Manhuaçu;Santana do Paraíso;Santo Antônio do Grama;São Domingos das Dores;São Domingos do Prata;São Gonçalo do Rio Abaixo;São João do Manhuaçu;São João do Oriente;São José do Goiabal;São José do Mantimento;São Pedro dos Ferros;São Sebastião do Anta;Sericita;Simonésia;Taparuba;Timóteo;Ubaporanga;Urucânia;Vargem Alegre;Vermelho Novo",
    "SRAI Belo Horizonte": "Araçaí;Baldim;Belo Horizonte;Betim;Bonfim;Brumadinho;Cachoeira da Prata;Caetanópolis;Caeté;Capim Branco;Confins;Contagem;Cordisburgo;Crucilândia;Diogo de Vasconcelos;Esmeraldas;Florestal;Fortuna de Minas;Funilândia;Ibirité;Igarapé;Inhaúma;Itabirito;Itaguara;Jaboticatubas;Jequitibá;Juatuba;Lagoa Santa;Mariana;Mário Campos;Mateus Leme;Matozinhos;Nova Lima;Nova União;Ouro Preto;Paraopeba;Pedro Leopoldo;Piedade dos Gerais;Prudente de Morais;Raposos;Ribeirão das Neves;Rio Acima;Rio Manso;Sabará;Santa Luzia;Santana de Pirapama;Santana do Riacho;São Joaquim de Bicas;São José da Lapa;Sarzedo;Sete Lagoas;Taquaraçu de Minas;Vespasiano",
    "SRAI Lavras": "Aguanil;Boa Esperança;Bom Sucesso;Cambuquira;Campanha;Campo Belo;Cana Verde;Candeias;Carmo da Cachoeira;Carmo da Mata;Carmópolis de Minas;Carrancas;Coqueiral;Conceição do Rio Verde;Cristais;Elói Mendes;Guapé;Ibituruna;Ijaci;Ilicínea;Ingaí;Itumirim;Itutinga;Jesuânia;Lambari;Lavras;Luminárias;Monsenhor Paulo;Nepomuceno;Olímpio Noronha;Oliveira;Passa Tempo;Perdões;Piracema;Ribeirão Vermelho;Santana da Vargem;Santana do Jacaré;Santo Antônio do Amparo;São Bento Abade;São Francisco de Paula;São Thomé das Letras;Três Corações;Três Pontas;Varginha",
    "SRAI Pouso Alegre": "Aiuruoca;Albertina;Alagoa;Baependi;Bocaina de Minas;Bom Repouso;Borda da Mata;Brasópolis;Bueno Brandão;Cachoeira de Minas;Camanducaia;Cambuí;Careaçu;Carmo de Minas;Carvalhos;Caxambu;Conceição das Pedras;Conceição dos Ouros;Congonhal;Consolação;Cordislândia;Córrego do Bom Jesus;Cristina;Cruzília;Delfim Moreira;Dom Viçoso;Espírito Santo do Dourado;Estiva;Extrema;Gonçalves;Heliodora;Inconfidentes;Itajubá;Itamonte;Itanhandu;Itapeva;Jacutinga;Liberdade;Maria da Fé;Marmelópolis;Minduri;Monte Sião;Munhoz;Natércia;Ouro Fino;Paraisópolis;Passa Quatro;Passa-Vinte;Pedralva;Piranguçu;Piranguinho;Pouso Alegre;Pouso Alto;Santa Rita do Sapucaí;Sapucaí-Mirim;São Gonçalo do Sapucaí;São João da Mata;São José do Alegre;São Lourenço;São Sebastião da Bela Vista;São Sebastião do Rio Verde;Senador Amaral;Senador José Bento;Seritinga;Serranos;Silvianópolis;Soledade de Minas;Tocos do Moji;Toledo;Turvolândia;Virgínia;Wenceslau Braz"
  };
  var SRAIS = {};
  Object.keys(SRAIS_TXT).forEach(function (k) { SRAIS[k] = SRAIS_TXT[k].split(";"); });

  // ------------------------------------------------------------ utilitários
  function normalizar(t) {
    return (t || "").trim().replace(/\s+/g, " ").normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }
  function contem(texto, palavra) {
    return !!texto && !!palavra && texto.toLowerCase().indexOf(palavra.toLowerCase()) >= 0;
  }
  function tituloNome(nome) {
    return (nome || "").trim().replace(/\s+/g, " ").toLowerCase().replace(/(^|[\s\-'])(\S)/g, function (m, a, b) { return a + b.toUpperCase(); });
  }
  function soDigitos(v) { return String(v || "").replace(/\D/g, ""); }
  function apenasData(v) {
    if (!v) return "";
    var s = String(v).trim();
    var m = s.match(/(\d{2}\/\d{2}\/\d{4})/); if (m) return m[1];
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[3] + "/" + m[2] + "/" + m[1];
    return s;
  }
  function formatarUnidade(u) {
    if (!u) return "";
    if (typeof u === "string") return u.trim();
    var sigla = (u.sigla || "").trim(), desc = (u.descricao || "").trim();
    return sigla ? sigla + " - " + desc : desc;
  }
  function isoParaBr(iso) { var p = iso.split("-"); return p[2] + "/" + p[1] + "/" + p[0]; }
  var MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  function slugData(br) { var p = br.split("/"); return parseInt(p[0], 10) + MESES[parseInt(p[1], 10) - 1]; }
  function slugSrai(n) {
    return normalizar(n.replace("SRAI", "")).split(" ").filter(Boolean).map(function (w) { return w[0].toUpperCase() + w.slice(1); }).join("") || "SRAI";
  }

  // -------------------------------------------------------------------- API
  var cancelado = false;
  function tokenAtual() { return localStorage.getItem("Authorization"); }
  function minutosRestantesToken() {
    try {
      var p = JSON.parse(atob(tokenAtual().split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
      return Math.round((p.exp * 1000 - Date.now()) / 60000);
    } catch (e) { return null; }
  }
  function esperar(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  // Falhas passageiras (rede instável, HTTP 429/5xx) são tentadas de novo até 3 vezes.
  async function fetchComRetentativa(caminho) {
    var ultimoErro;
    for (var tentativa = 1; tentativa <= 4; tentativa++) {
      if (cancelado) throw new Error("Cancelado pelo usuário");
      try {
        var resp = await fetch(caminho, { headers: { Authorization: "Bearer " + tokenAtual() } });
        if ((resp.status === 429 || resp.status >= 500) && tentativa < 4) { ultimoErro = new Error("HTTP " + resp.status); await esperar(1000 * tentativa); continue; }
        return resp;
      } catch (e) {
        ultimoErro = e;
        if (tentativa < 4) await esperar(1000 * tentativa);
      }
    }
    throw new Error("Falha de rede após 4 tentativas (" + (ultimoErro && ultimoErro.message) + ")");
  }
  async function api(caminho) {
    if (cancelado) throw new Error("Cancelado pelo usuário");
    var resp = await fetchComRetentativa(caminho);
    if (resp.status === 401 || resp.status === 403) {
      throw new Error("Sessão expirada ou sem permissão (HTTP " + resp.status + "). Faça login de novo no Ecosistemas.");
    }
    if (!resp.ok) {
      var corpo = ""; try { corpo = (await resp.text()).slice(0, 200); } catch (e) {}
      throw new Error("HTTP " + resp.status + " em " + caminho.split("?")[0] + (corpo ? " — " + corpo : ""));
    }
    var txt = await resp.text();
    return txt ? JSON.parse(txt) : null;
  }

  var mapaMun = null; // {porNome, porNomeNorm, porId}
  async function carregarMunicipios() {
    if (mapaMun) return mapaMun;
    var lista = await api("/fisc/municipios/MG");
    mapaMun = { porNome: {}, porNomeNorm: {}, porId: {} };
    lista.forEach(function (m) {
      mapaMun.porNome[m.descricao.trim().toLowerCase()] = m.id;
      mapaMun.porNomeNorm[normalizar(m.descricao)] = m.id;
      mapaMun.porId[m.id] = m.descricao;
    });
    return mapaMun;
  }
  function resolverIdMunicipio(nome) {
    var k = nome.trim().toLowerCase();
    if (k in mapaMun.porNome) return { id: mapaMun.porNome[k], fallback: false };
    var n = normalizar(nome);
    if (n in mapaMun.porNomeNorm) return { id: mapaMun.porNomeNorm[n], fallback: true };
    return null;
  }

  async function buscarAtos(idMun, dataIni, dataFim) {
    var todos = [], pagina = 0, tam = 2000;
    while (true) {
      var q = "/fisc/atofiscalizacao/buscarPorFiltro/?numeroProcesso=&statusProcesso=&idOpVinculada=" +
        "&idCodMunicipio=" + idMun + "&idUnidadeResponsavel=" +
        "&dataPeriodoInicial=" + encodeURIComponent(dataIni) +
        "&dataPeriodoFinal=" + encodeURIComponent(dataFim) +
        "&servidorResponsavel=&membroEquipe=&cpfCnpjEnvolvido=&page=" + pagina + "&size=" + tam;
      var r = await api(q) || {};
      var c = r.content || [];
      todos = todos.concat(c);
      var total = r.totalElements != null ? r.totalElements : todos.length;
      if (todos.length >= total || !c.length) break;
      pagina++;
    }
    return todos;
  }

  var cacheDetalheAto = {};
  function detalheAto(id) {
    if (!cacheDetalheAto[id]) cacheDetalheAto[id] = api("/fisc/atofiscalizacao/" + id);
    return cacheDetalheAto[id];
  }

  // ----------------------------------------------------------------- Fase 1
  async function fase1(municipios, dataIni, dataFim, ui) {
    var res = { atosMamb: [], semUnidade: [], cancelados: 0, munErro: [], totalAtos: 0 };
    var ids = {};
    ui.log("Validando " + municipios.length + " município(s)...");
    await carregarMunicipios();
    var validos = [];
    municipios.forEach(function (m) {
      var r = resolverIdMunicipio(m);
      if (!r) { res.munErro.push({ municipio: m, erro: "Nome não encontrado na lista oficial do sistema (mesmo ignorando acentos)." }); ui.log("  [ERRO] Município não reconhecido: " + m, "erro"); }
      else { if (r.fallback) ui.log("  [AVISO] '" + m + "' reconhecido só ignorando acentos.", "aviso"); validos.push({ nome: m, id: r.id }); }
    });

    for (var i = 0; i < validos.length; i++) {
      var mun = validos[i];
      ui.progresso("Fase 1: " + mun.nome, i / validos.length * 0.4);
      try {
        var regs = await buscarAtos(mun.id, dataIni, dataFim);
        res.totalAtos += regs.length;
        var nMamb = 0;
        for (var j = 0; j < regs.length; j++) {
          var reg = regs[j];
          if (contem((reg.statusFisc || "").trim(), STATUS_CANCELADO)) { res.cancelados++; continue; }
          var unidade = formatarUnidade(reg.unidadeResponsavel);
          var numero = reg.numProcesso || String(reg.idAtoFisc);
          if (contem(unidade, FILTRO_UNIDADE)) {
            if (!ids[reg.idAtoFisc]) { ids[reg.idAtoFisc] = 1; res.atosMamb.push({ id: reg.idAtoFisc, numero: numero }); nMamb++; }
            continue;
          }
          if (unidade) continue;
          // 2ª verificação: unidade vazia na listagem → detalhe do Ato
          var unidadeReal = "";
          try { unidadeReal = formatarUnidade((await detalheAto(reg.idAtoFisc) || {}).unidadeResponsavel); }
          catch (e) { if (cancelado) throw e; ui.log("  [AVISO] Falha na 2ª verificação do Ato " + numero + ": " + e.message, "aviso"); }
          if (contem(unidadeReal, FILTRO_UNIDADE)) {
            if (!ids[reg.idAtoFisc]) { ids[reg.idAtoFisc] = 1; res.atosMamb.push({ id: reg.idAtoFisc, numero: numero }); nMamb++; }
            ui.log("  Ato " + numero + ": unidade vazia na listagem, 2ª verificação achou '" + unidadeReal + "'.");
          } else {
            res.semUnidade.push({ numero: numero, municipio: mun.nome, servidor: reg.nomeRazaoServResponsavel || "", unidade2: unidadeReal });
          }
        }
        ui.log(mun.nome + ": " + regs.length + " Ato(s), " + nMamb + " MAmb.");
      } catch (e) {
        if (cancelado) throw e;
        res.munErro.push({ municipio: mun.nome, erro: e.message });
        ui.log("  [ERRO] " + mun.nome + ": " + e.message, "erro");
      }
    }
    return res;
  }

  // ----------------------------------------------------------------- Fase 2
  async function extrairAuto(item) {
    var idAuto = item.idAtoInfracao;
    var d = await api("/fisc/autodeinfracao/" + idAuto) || {};
    var resp = d.responsavel || {};
    return {
      numeroAuto: (item.numInfracao && item.anoInfracao) ? item.numInfracao + "/" + item.anoInfracao : String(idAuto),
      autuado: item.nomeFiscalizado || item.nomeAutuado || "",
      data: apenasData(d.dataHoraConfirmacao),
      servidor: tituloNome(resp.nomeRazaoSocial),
      masp: soDigitos(resp.matriculaMasp),
      unidade: formatarUnidade(resp.unidadeDoServidor)
    };
  }

  // ------------------------------------------------------- Cientificação
  // Rótulos iguais aos da tela do GAIA.
  var TIPOS_CIENTIFICACAO = {
    IMEDIATA: "Imediata (presencial)", ELETRONICA: "Eletrônica (e-mail)",
    TERMO_OFICIO: "Postal (Correios)", DIARIO_OFICIAL: "Diário Oficial"
  };
  var STATUS_CIENTIFICACAO = {
    EM_ELABORACAO: "Em elaboração", NAO_ENVIADO: "Não enviado", NAO_ASSINADO: "Não assinado",
    EM_ANDAMENTO: "Em andamento", CIENTIFICADO: "Cientificado", N_CONCRETIZADA: "Não concretizada"
  };
  function rotulo(mapa, cod) { return cod ? (mapa[cod] || String(cod)) : ""; }

  // Escolhe a cientificação de um auto: a vigente, se houver; senão a mais recente.
  function cientificacaoDoAuto(lista, numeroAuto) {
    var doAuto = (lista || []).filter(function (c) {
      return (c.documentosVinculados || []).some(function (v) { return v.tipo === "INFRACAO" && v.numeroDocumento === numeroAuto; });
    });
    if (!doAuto.length) return null;
    var vigente = doAuto.filter(function (c) { return c.cientificacaoVigente === true; });
    var base = vigente.length ? vigente : doAuto;
    base.sort(function (a, b) { return String(b.dataHoraCriacao || "").localeCompare(String(a.dataHoraCriacao || "")); });
    return { escolhida: base[0], total: doAuto.length };
  }

  async function processarAto(ato, ui) {
    var det = await detalheAto(ato.id) || {};
    if (contem((det.statusProcessoAtoFisc || "").trim(), STATUS_CANCELADO)) { ui.log("  Ato " + ato.numero + ": Cancelado. Descartado."); return []; }
    var itens = await api("/fisc/autodeinfracao/lavratura/" + ato.id + "/listarinfracoes") || [];
    if (!itens.length) return [];
    var primeiro = await extrairAuto(itens[0]);
    if (!contem(primeiro.unidade, FILTRO_UNIDADE)) {
      ui.log("  Ato " + ato.numero + ": unidade do 1º Auto ('" + primeiro.unidade + "') sem MAmb. Descartado.");
      return [];
    }
    var autos = [primeiro];
    for (var i = 1; i < itens.length; i++) autos.push(await extrairAuto(itens[i]));
    var municipio = mapaMun.porId[det.idMunicipio] || "";
    if (!municipio) ui.log("  [AVISO] Ato " + ato.numero + ": município (Local da fiscalização) não resolvido.", "aviso");
    var cients = null;
    try { cients = await api("/fisc/documentoscientificacao/atofiscalizacao/" + ato.id) || []; }
    catch (e) { if (cancelado) throw e; ui.log("  [AVISO] Ato " + ato.numero + ": não foi possível ler a cientificação (" + e.message + ").", "aviso"); }
    ui.log("  Ato " + ato.numero + ": " + autos.length + " Auto(s) aceito(s).");
    return autos.map(function (a) {
      var tipo = "", situacao = "";
      if (cients === null) { tipo = "Não verificado"; situacao = "Não verificado"; }
      else {
        var c = cientificacaoDoAuto(cients, a.numeroAuto);
        if (!c) { tipo = "Sem cientificação"; situacao = "Sem cientificação"; ui.log("  [AUDITORIA] Auto " + a.numeroAuto + " (Ato " + ato.numero + "): sem cientificação registrada.", "aviso"); }
        else {
          tipo = rotulo(TIPOS_CIENTIFICACAO, c.escolhida.tipoCientificacao);
          situacao = rotulo(STATUS_CIENTIFICACAO, c.escolhida.status);
          if (c.total > 1) ui.log("  Auto " + a.numeroAuto + ": " + c.total + " cientificações; usada a " + (c.escolhida.cientificacaoVigente ? "vigente" : "mais recente") + ".");
        }
      }
      return { ato: ato.numero, auto: a.numeroAuto, data: a.data, autuado: a.autuado, municipio: municipio, militar: a.servidor, matricula: a.masp, tipoCient: tipo, situacaoCient: situacao };
    });
  }

  async function fase2(atos, ui) {
    var linhas = new Array(atos.length), prox = 0, feitos = 0, erros = [];
    async function trabalhador() {
      while (prox < atos.length) {
        var k = prox++;
        try { linhas[k] = await processarAto(atos[k], ui); }
        catch (e) { if (cancelado) throw e; linhas[k] = []; erros.push(atos[k].numero); ui.log("  [ERRO] Ato " + atos[k].numero + ": " + e.message, "erro"); }
        feitos++;
        ui.progresso("Fase 2: " + feitos + "/" + atos.length + " Atos", 0.4 + feitos / atos.length * 0.55);
      }
    }
    var ts = []; for (var t = 0; t < CONCORRENCIA_FASE2; t++) ts.push(trabalhador());
    await Promise.all(ts);
    return { linhas: [].concat.apply([], linhas), erros: erros };
  }

  // ------------------------------------------------------------- Relatório
  function carregarSheetJS() {
    if (window.XLSX) return Promise.resolve();
    return new Promise(function (ok, falha) {
      var s = document.createElement("script");
      s.src = URL_SHEETJS; s.onload = ok;
      s.onerror = function () { falha(new Error("Não foi possível carregar a biblioteca de Excel (sem internet ou bloqueada).")); };
      document.head.appendChild(s);
    });
  }
  function ajustarLarguras(ws, linhas) {
    var larg = [];
    linhas.forEach(function (l) { l.forEach(function (v, i) { larg[i] = Math.min(60, Math.max(larg[i] || 0, String(v == null ? "" : v).length + 2)); }); });
    ws["!cols"] = larg.map(function (w) { return { wch: w }; });
  }
  function gerarExcel(nomeSrai, dataIni, dataFim, linhas, f1, errosF2) {
    var wb = XLSX.utils.book_new();
    var cab = ["ATO DE FISCALIZAÇÃO", "AUTO DE INFRAÇÃO", "DATA", "AUTUADO", "MUNICÍPIO", "MILITAR", "MATRÍCULA", "TIPO DE CIENTIFICAÇÃO", "SITUAÇÃO DA CIENTIFICAÇÃO"];
    var dados = [cab].concat(linhas.map(function (l) { return [l.ato, l.auto, l.data, l.autuado, l.municipio, l.militar, l.matricula, l.tipoCient, l.situacaoCient]; }));
    var ws = XLSX.utils.aoa_to_sheet(dados); ajustarLarguras(ws, dados);
    XLSX.utils.book_append_sheet(wb, ws, "Relatório");

    if (f1.semUnidade.length) {
      var d2 = [["ATO DE FISCALIZAÇÃO", "MUNICÍPIO PESQUISADO", "SERVIDOR RESPONSÁVEL (lista)", "UNIDADE NA 2ª VERIFICAÇÃO"]]
        .concat(f1.semUnidade.map(function (r) { return [r.numero, r.municipio, r.servidor, r.unidade2 || "(também vazia)"]; }));
      var ws2 = XLSX.utils.aoa_to_sheet(d2); ajustarLarguras(ws2, d2);
      XLSX.utils.book_append_sheet(wb, ws2, "Revisão manual");
    }
    var probs = f1.munErro.map(function (m) { return [m.municipio, m.erro, "Conferir a grafia/lista da SRAI e rodar de novo para este município."]; })
      .concat(errosF2.map(function (n) { return ["Ato " + n, "Falha ao processar o Ato na Fase 2", "Conferir manualmente no GAIA."]; }));
    if (probs.length) {
      var d3 = [["MUNICÍPIO / ATO", "MOTIVO", "AÇÃO RECOMENDADA"]].concat(probs);
      var ws3 = XLSX.utils.aoa_to_sheet(d3); ajustarLarguras(ws3, d3);
      XLSX.utils.book_append_sheet(wb, ws3, "Municípios com problema");
    }
    var nome = "relatorio_" + slugSrai(nomeSrai) + "_" + slugData(dataIni) + "_" + slugData(dataFim) + ".xlsx";
    XLSX.writeFile(wb, nome);
    return nome;
  }

  // -------------------------------------------------------------------- UI
  var cor = { azul: "#2e4a36", fundo: "#ffffff", borda: "#c9d3de", texto: "#1d2733", suave: "#5b6b7c", erro: "#b3261e", aviso: "#8a5a00", ok: "#1e6b34" };
  var raiz = document.createElement("div");
  raiz.id = "gaia-relatorio-painel";
  raiz.style.cssText = "position:fixed;top:16px;right:16px;z-index:2147483647;width:420px;max-width:calc(100vw - 32px);" +
    "background:" + cor.fundo + ";color:" + cor.texto + ";border:1px solid " + cor.borda + ";border-radius:10px;" +
    "box-shadow:0 8px 28px rgba(0,0,0,.25);font:14px/1.4 Segoe UI,Arial,sans-serif;";
  var hoje = new Date(), ini = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  function isoLocal(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  var campo = "width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid " + cor.borda + ";border-radius:6px;font:inherit;color:inherit;background:#fff;";
  raiz.innerHTML =
    '<div style="background:' + cor.azul + ';color:#fff;padding:8px 12px;border-radius:10px 10px 0 0;display:flex;align-items:center;justify-content:space-between">' +
      '<div style="display:flex;align-items:center;gap:10px"><img src="' + BRASAO + '" alt="" style="height:40px;width:auto;display:block">' +
      '<div><strong style="font-size:15px">Extrator de Autos</strong><div style="font-size:11px;opacity:.75">BPM MAmb · v' + VERSAO + '</div></div></div>' +
      '<button data-x="fechar" title="Fechar" style="background:none;border:0;color:#fff;font-size:20px;cursor:pointer;line-height:1">×</button></div>' +
    '<div style="padding:14px">' +
      '<label style="display:block;margin-bottom:10px">SRAI<select data-x="srai" style="' + campo + 'margin-top:4px"></select></label>' +
      '<div style="display:flex;gap:10px;margin-bottom:12px">' +
        '<label style="flex:1">Data inicial<input data-x="ini" type="date" style="' + campo + 'margin-top:4px"></label>' +
        '<label style="flex:1">Data final<input data-x="fim" type="date" style="' + campo + 'margin-top:4px"></label></div>' +
      '<div style="display:flex;gap:8px">' +
        '<button data-x="gerar" style="flex:1;padding:8px;background:' + cor.azul + ';color:#fff;border:0;border-radius:6px;font:inherit;font-weight:600;cursor:pointer">Gerar relatório</button>' +
        '<button data-x="cancelar" style="display:none;padding:8px 12px;background:#fff;color:' + cor.erro + ';border:1px solid ' + cor.erro + ';border-radius:6px;font:inherit;cursor:pointer">Cancelar</button></div>' +
      '<div data-x="barraBox" style="display:none;margin-top:12px"><div data-x="etapa" style="font-size:12px;color:' + cor.suave + ';margin-bottom:4px"></div>' +
        '<div style="height:8px;background:#e6ebf0;border-radius:4px;overflow:hidden"><div data-x="barra" style="height:100%;width:0;background:' + cor.azul + ';transition:width .2s"></div></div></div>' +
      '<div data-x="log" style="margin-top:12px;max-height:220px;overflow:auto;font:12px/1.45 Consolas,monospace;background:#f5f7f9;border-radius:6px;padding:8px;display:none;white-space:pre-wrap"></div>' +
      '<div data-x="sessao" style="margin-top:10px;font-size:12px;color:' + cor.suave + '"></div>' +
      '<div style="margin-top:10px;padding-top:8px;border-top:1px solid ' + cor.borda + ';font-size:11px;color:' + cor.suave + '">Desenvolvedor: 1º Sgt Anderson Arcari Pereira - 5ª Cia PM MAmb</div>' +
    '</div>';
  document.body.appendChild(raiz);
  function el(x) { return raiz.querySelector('[data-x="' + x + '"]'); }
  Object.keys(SRAIS).forEach(function (n) { var o = document.createElement("option"); o.value = n; o.textContent = n; el("srai").appendChild(o); });
  el("ini").value = isoLocal(ini); el("fim").value = isoLocal(hoje);

  var ui = {
    log: function (msg, tipo) {
      var box = el("log"); box.style.display = "block";
      var linha = document.createElement("div"); linha.textContent = msg;
      if (tipo) linha.style.color = cor[tipo] || cor.texto;
      box.appendChild(linha); box.scrollTop = box.scrollHeight;
    },
    progresso: function (txt, frac) { el("barraBox").style.display = "block"; el("etapa").textContent = txt; el("barra").style.width = Math.round(frac * 100) + "%"; }
  };
  function atualizarSessao() {
    var m = minutosRestantesToken();
    el("sessao").textContent = m == null ? "Não encontrei a sessão — faça login no Ecosistemas." :
      m <= 0 ? "Sessão expirada — faça login de novo." : "Sessão válida por mais " + (m >= 60 ? Math.floor(m / 60) + "h" + String(m % 60).padStart(2, "0") : m + " min") + ".";
  }
  atualizarSessao(); var timerSessao = setInterval(atualizarSessao, 60000);

  var rodando = false;
  el("fechar").onclick = function () { if (rodando) cancelado = true; raiz.style.display = "none"; };
  el("cancelar").onclick = function () { cancelado = true; ui.log("Cancelando...", "aviso"); };
  el("gerar").onclick = async function () {
    var srai = el("srai").value, iIso = el("ini").value, fIso = el("fim").value;
    if (!iIso || !fIso) { ui.log("Preencha as duas datas.", "erro"); return; }
    if (iIso > fIso) { ui.log("A data inicial é depois da data final.", "erro"); return; }
    var m = minutosRestantesToken();
    if (m == null || m <= 0) { ui.log("Sessão não encontrada ou expirada. Faça login no Ecosistemas.", "erro"); return; }
    var dIni = isoParaBr(iIso), dFim = isoParaBr(fIso);
    rodando = true; cancelado = false; cacheDetalheAto = {};
    el("gerar").disabled = true; el("gerar").style.opacity = ".5"; el("cancelar").style.display = "block";
    el("log").innerHTML = "";
    var t0 = Date.now();
    try {
      ui.log("Extrator de Autos v" + VERSAO + " · " + srai + " — " + dIni + " a " + dFim);
      var libPronta = carregarSheetJS();
      var f1 = await fase1(SRAIS[srai], dIni, dFim, ui);
      ui.log("Fase 1: " + f1.totalAtos + " Atos lidos, " + f1.atosMamb.length + " MAmb, " + f1.cancelados + " cancelado(s), " + f1.semUnidade.length + " para revisão manual.", "ok");
      var f2 = f1.atosMamb.length ? await fase2(f1.atosMamb, ui) : { linhas: [], erros: [] };
      ui.progresso("Gerando Excel...", 0.97);
      await libPronta;
      var arq = gerarExcel(srai, dIni, dFim, f2.linhas, f1, f2.erros);
      ui.progresso("Concluído em " + Math.round((Date.now() - t0) / 1000) + " s", 1);
      ui.log("Concluído: " + f2.linhas.length + " Auto(s) no relatório. Arquivo baixado: " + arq, "ok");
      if (f1.munErro.length || f2.erros.length) ui.log("Atenção: veja a aba 'Municípios com problema' no Excel.", "aviso");
    } catch (e) {
      ui.log(cancelado ? "Execução cancelada." : "[ERRO] " + e.message, cancelado ? "aviso" : "erro");
      ui.progresso(cancelado ? "Cancelado" : "Interrompido por erro", 0);
    } finally {
      rodando = false; el("gerar").disabled = false; el("gerar").style.opacity = "1"; el("cancelar").style.display = "none";
    }
  };

  window.__gaiaRelatorio = { versao: VERSAO, mostrar: function () { raiz.style.display = "block"; atualizarSessao(); } };
})();
