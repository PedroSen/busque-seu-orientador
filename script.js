// URL do arquivo Excel hospedado no GitHub com os dados dos professores e suas áreas de interesse.
const WORKBOOK_URL =
  'https://raw.githubusercontent.com/PedroSen/busque-seu-orientador/main/Tags%20TCCs%20IC(1).xlsx';

const form = document.querySelector('#search-form');
const searchInput = document.querySelector('#especialidade-search');
const sugestoes = document.querySelector('#especialidade-options');
const results = document.querySelector('#results');

let professores = [];

function normalizar(text) {
  return String(text ?? '').trim().toLocaleLowerCase('pt-BR');
}

function primeiroValor(rows, columnIndex, headers) {
  const ignorar = new Set(headers.map(normalizar));
  const value = rows
    .map((row) => String(row[columnIndex] ?? '').trim())
    .find((item) => item && !ignorar.has(normalizar(item)));

  return value ?? '';
}

function criarProfessores(workbook) {
  return workbook.SheetNames.map((nome) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[nome], {
      header: 1,
      defval: '',
    });

    const especialidades = rows
      .map(([label, peso]) => {
        const score = Number(peso);
        if (!label || !Number.isFinite(score) || score <= 1) return null;
        return { nome: String(label).trim(), peso: score };
      })
      .filter(Boolean);

    // As colunas C e D contêm, respectivamente, o Lattes e o e-mail do professor.
    const lattes = primeiroValor(rows, 2, ['lattes', 'link lattes', 'currículo lattes']);
    const email = primeiroValor(rows, 3, ['e-mail', 'email', 'e mail']);

    return { nome, especialidades, lattes, email };
  }).filter((professor) => professor.especialidades.length > 0);
}

function obterBusca() {
  return searchInput.value
    .split(',')
    .map((termo) => normalizar(termo))
    .filter(Boolean);
}

function encontrarMatches(termos) {
  return professores
    .map((professor) => {
      const especialidades = professor.especialidades.filter((item) =>
        termos.some((termo) => normalizar(item.nome).includes(termo)),
      );
      const score = especialidades.reduce((total, item) => total + item.peso, 0);
      return { ...professor, matchEspecialidades: especialidades, score };
    })
    .filter((professor) => termos.length === 0 || professor.score > 0)
    .sort((a, b) => b.score - a.score);
}

function adicionarLink(container, rotulo, valor, tipo) {
  const bloco = document.createElement('div');
  const titulo = document.createElement('h4');
  titulo.textContent = rotulo;
  bloco.append(titulo);

  if (valor) {
    const link = document.createElement('a');
    if (tipo === 'email') {
      link.href = `mailto:${valor.replace(/^mailto:/i, '')}`;
      link.textContent = valor.replace(/^mailto:/i, '');
    } else {
      const url = /^https?:\/\//i.test(valor) ? valor : `https://${valor}`;
      try {
        const parsedUrl = new URL(url);
        if (parsedUrl.protocol === 'https:' || parsedUrl.protocol === 'http:') {
          link.href = parsedUrl.href;
          link.textContent = 'Acessar Lattes';
          link.target = '_blank';
          link.rel = 'noopener noreferrer';
        }
      } catch {
        // Não exibe link inválido.
      }
    }
    if (link.href) bloco.append(link);
  } else {
    const indisponivel = document.createElement('span');
    indisponivel.textContent = 'Não informado';
    bloco.append(indisponivel);
  }

  container.append(bloco);
}

function createCard(professor, termos) {
  const card = document.createElement('article');
  card.className = 'professor-card';

  const heading = document.createElement('div');
  heading.className = 'card-heading';
  const nomeContainer = document.createElement('div');
  const nome = document.createElement('h3');
  nome.textContent = professor.nome;
  const cargo = document.createElement('p');
  cargo.className = 'card-label';
  cargo.textContent = 'PROFESSOR(A) TITULAR';
  nomeContainer.append(nome, cargo);
  heading.append(nomeContainer);
  card.append(heading);

  const especialidades = termos.length
    ? professor.matchEspecialidades
    : professor.especialidades.slice(0, 4);
  const tags = document.createElement('div');
  tags.className = 'especialidades';
  especialidades.forEach((item) => {
    const tag = document.createElement('span');
    tag.className = 'especialidade-tag';
    tag.textContent = item.nome;
    tags.append(tag);
  });
  card.append(tags);

  const links = document.createElement('div');
  links.className = 'links';
  adicionarLink(links, 'E-mail', professor.email, 'email');
  adicionarLink(links, 'Lattes', professor.lattes, 'lattes');
  card.append(links);

  return card;
}

function exibirResultados() {
  const termos = obterBusca();
  const matches = encontrarMatches(termos);
  results.replaceChildren();
  matches.forEach((professor) => results.append(createCard(professor, termos)));
}

function addSugestoes() {
  const especialidades = new Set(
    professores.flatMap((professor) =>
      professor.especialidades.map((especialidade) => especialidade.nome),
    ),
  );

  especialidades.forEach((nome) => {
    const option = document.createElement('option');
    option.value = nome;
    sugestoes.append(option);
  });
}

form.addEventListener('submit', (event) => {
  event.preventDefault();
  exibirResultados();
});

async function start() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Falha ao carregar a planilha (${response.status}).`);

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    professores = criarProfessores(workbook);
    addSugestoes();
    exibirResultados();
  } catch (error) {
    const status = document.querySelector('#status');
    status.textContent = 'Não foi possível carregar os dados dos professores. Tente novamente mais tarde.';
    console.error(error);
  }
}

start();
