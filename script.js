//Define a URL do arquivo Excel hospedado no GitHub que contém os dados dos professores e suas áreas de interesse
const WORKBOOK_URL =
  'https://raw.githubusercontent.com/PedroSen/busque-seu-orientador/main/Tags%20TCCs%20IC(1).xlsx';

//Variáveis dos elementos HTML que vão ser manipulados
const form = document.querySelector('#search-form');
const searchInput = document.querySelector('#especialidade-search');
const sugestoes = document.querySelector('#especialidade-options');
const results = document.querySelector('#results');

let professores = [];

//Transforma os inputs do usuário em strings
function normalizar(text) {
  return String(text)
}

function criarProfessores(workbook) {
  return workbook.SheetNames.map((nome) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[nome], {
      header: 1,
      defval: '',
    });

    //Cria um objeto especialidades que será um dos atributos do objeto professor
    const especialidades = rows
      .map(([nome, peso], index) => {
        const label = String(nome);
        const score = Number(peso);

        if (score <= 1) {
          return null;
        }

        return { nome: String(label), peso: score };
      })
      .filter(Boolean);

    const lattes = String(rows[0][3]);
    const email = String(rows[0][2]);

    //Retorna um objeto "professor" com o nome da aba/professor e suas especialidades
    return { nome, especialidades, lattes, email };
  }).filter((professor) => professor.especialidades.length > 0);
}

// Obtém os termos digitados pelo usuário
function obterBusca() {
  return searchInput.value
    .split(',')
    .filter(Boolean);
}

// Busca os professores que melhor combinam com os termos informados.
function encontrarMatches(termos) {
  return professores
  //a função abaixo itera por todos os professores e verifica se o termo buscado está entre suas especialidades, calcula os scores com base no termo e retorna os professores com scores maiores que 0
    .map((professor) => {
      const especialidades = professor.especialidades.filter((especialidade) =>
        termos.some((termo) => especialidade.nome.includes(termo)),
      );
      const score = especialidades.reduce((total, item) => total + item.peso, 0);
      return { ...professor, matchEspecialidades: especialidades, score };
    })
    .filter((professor) => termos.length === 0 || professor.score > 0)
    .sort((a, b) => b.score - a.score);
}

// Cria um card HTML para exibir um professor na tela
function createCard(professor, termos) {
  const card = document.createElement('div');
  // Define quais especialidades serão exibidas: as que deram match se houver busca, ou as quatro primeiras em caso de listagem geral.
  const especialidades = termos.length
    ? professor.matchEspecialidades
    : professor.especialidades.slice(0, 4);

  const tags = especialidades
    .map((item) => `<span class="especialidade-tag">${item.nome}</span>`)
    .join('');

  card.className = 'professor-card';
  card.innerHTML = `
    <div class="card-heading">
      <div>
        <h3>${professor.nome}</h3>
        <p class="card-label" style="margin: 10px 0 0 2px">PROFESSOR(A) TITULAR</p>
      </div>
    </div>
    <div class="especialidades">${tags}</div>
    <div class="links">
        <div class="div-email">
            <h4>E-mail</h4>
            <a href = mailto:${professor.email}>${professor.email}</a>
        </div>
        <div class="div-lattes">
            <h4>Lattes</h4>
            <a href=${professor.lattes}>${professor.lattes}</a>
        </div>
    </div>
  `;
  return card;
}

// exibe na tela os resultados filtrados de acordo com a busca.
function exibirResultados() {
  const termos = obterBusca();
  const matches = encontrarMatches(termos);

  // Limpa o container de resultados antes de preencher novamente.
  results.replaceChildren();

  // Para cada professor encontrado, cria um card e o adiciona na página.
  matches.forEach((professor) => {
    results.append(createCard(professor, termos));
  });
}

// Preenche o campo de sugestões com todas as especialidades disponíveis.
function addSugestoes() {
  // Cria um conjunto para remover especialidades duplicadas.
  const especialidades = new Set(
    // Para cada professor, pega todas as especialidades e transforma em um único array.
    professores.flatMap((professor) =>
      professor.especialidades.map((especialidade) => especialidade.nome),
    ),
  );

  // Para cada especialidade única, cria um option dentro do select de sugestões.
  especialidades.forEach((nome) => {
    const option = document.createElement('option');
    option.value = nome;
    sugestoes.append(option);
  });
}

//Realiza a busca quando o usuário insere um termo
form.addEventListener('submit', (event) => {
  event.preventDefault();
  exibirResultados();
});

async function start() {
    const response = await fetch(WORKBOOK_URL);
    // Converte a resposta em um array de bytes para poder ler o workbook.
    const data = await response.arrayBuffer();
    // Lê o arquivo Excel em memória usando a biblioteca XLSX.
    const workbook = XLSX.read(data, { type: 'array' });
    // Armazena os professores carregados no array global.
    professores = criarProfessores(workbook);

    addSugestoes();
    exibirResultados();
}

start();
