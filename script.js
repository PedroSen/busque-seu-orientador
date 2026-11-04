//Define a URL do arquivo Excel hospedado no GitHub que contém os dados dos professores e suas áreas de interesse
const WORKBOOK_URL =
  'https://raw.githubusercontent.com/PedroSen/busque-seu-orientador/main/Tags%20TCCs%20IC(1).xlsx';

//Variáveis dos elementos HTML que vão ser manipulados
const form = document.querySelector('#search-form');
const searchInput = document.querySelector('#especialidade-search');
const sugestoes = document.querySelector('#especialidade-options');
const results = document.querySelector('#results');
const resultsCount = document.querySelector('#results-count');

let professors = [];

//Transforma os inputs do usuário em strings
function normalizar(text) {
  return String(text)
}

//Transforma as abas do excel em objetos
function readProfessors(workbook) {
  // Percorre cada aba do arquivo Excel e cria um objeto por professor.
  return workbook.SheetNames.map((nome) => {
    // Converte a planilha atual para uma matriz de linhas, usando a primeira linha como cabeçalho.
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[nome], {
      // Define que cada linha será um array, sem cabeçalho fixo.
      header: 1,
      // Define valor padrão para células vazias.
      defval: '',
    });

    // Filtra as linhas e monta um array de especialidades válidas.
    const especialidades = rows
      // Para cada linha, extrai nome e peso da especialidade.
      .map(([nome, peso], index) => {
        const label = String(nome);
        const score = Number(peso);

        if (score <= 1) {
          return null;
        }

        // Retorna um objeto "especialidade" com o nome e o peso da especialidade.
        return { nome: String(label), peso: score };
      })
      // Remove os itens nulos gerados por especialidades inválidas.
      .filter(Boolean);

    // Retorna um objeto "professor" com o nome da aba/professor e suas especialidades
    return { nome, especialidades };
  }).filter((professor) => professor.especialidades.length > 0);
}

// Obtém os termos digitados pelo usuário na busca separados por vírgulas.
function getSearchTerms() {
  // Lê o texto digitado no campo de busca.
  return searchInput.value
    // Divide a string em termos usando a vírgula como separador.
    .split(',')
    // Normaliza cada termo para padronizar a comparação.
    .map(normalizar)
    // Remove termos vazios para evitar buscas sem valor.
    .filter(Boolean);
}

// Busca os professores que melhor combinam com os termos informados.
function encontrarMatches(terms) {
  // Mapeia todos os professores para calcular sua pontuação.
  return professors
    .map((professor) => {
      // Filtra apenas as especialidades que contêm algum termo buscado.
      const especialidades = professor.especialidades.filter((especialidade) =>
        // Verifica se pelo menos um termo está contido no nome normalizado da especialidade.
        terms.some((term) => especialidade.nome.includes(term)),
      );
      // Soma os pesos das especialidades compatíveis para calcular a relevância do professor.
      const score = especialidades.reduce((total, item) => total + item.peso, 0);

      // Retorna o professor com suas especialidades encontradas e a pontuação total.
      return { ...professor, matchEspecialidades: especialidades, score };
    })
    // Se houver termos de busca, mantém apenas professores com pontuação maior que zero; se não houver, mostra todos.
    .filter((professor) => terms.length === 0 || professor.score > 0)
    // Ordena os professores pela pontuação mais alta primeiro.
    .sort((a, b) => b.score - a.score);
}

// Cria um card HTML para exibir um professor na tela.
function createCard(professor, terms) {
  // Cria um novo elemento div para representar o professor.
  const card = document.createElement('div');
  // Define quais especialidades serão exibidas: as que deram match se houver busca, ou as quatro primeiras em caso de listagem geral.
  const especialidades = terms.length
    ? professor.matchEspecialidades
    : professor.especialidades.slice(0, 4);
  // Gera os elementos HTML de tags para cada especialidade.
  const tags = especialidades
    .map((item) => `<span class="especialidade-tag">${item.nome}</span>`)
    .join('');
  // Se houver busca, exibe a pontuação do professor; caso contrário, não mostra nada.
  const score = terms.length
    ? `<span class="match-score">Peso ${professor.score}</span>`
    : '';

  // Define a classe CSS do card para estilizar a interface.
  card.className = 'professor-card';
  // Define o conteúdo interno do card com o nome do professor, sua classificação e as especialidades.
  card.innerHTML = `
    <div class="card-heading">
      <div>
        <h3>${professor.nome}</h3>
        <p class="card-label" style="margin: 10px 0 0 2px">PROFESSOR(A) TITULAR</p>
      </div>
      ${score}
    </div>
    <div class="especialidades">${tags}</div>
    <div class="links">
        <div>
            <h4>E-mail</h4>
            <a></a>
        </div>
        <div>
            <h4>Lattes</h4>
            <a></a>
        </div>
    </div>
  `;

  // Retorna o elemento pronto para ser inserido na página.
  return card;
}

// Exibe na tela os resultados filtrados de acordo com a busca.
function exibirResultados() {
  // Obtém os termos da busca.
  const terms = getSearchTerms();
  // Busca os professores que combinam com esses termos.
  const matches = encontrarMatches(terms);

  // Limpa o container de resultados antes de preencher novamente.
  results.replaceChildren();

  // Para cada professor encontrado, cria um card e o adiciona na página.
  matches.forEach((professor) => {
    results.append(createCard(professor, terms));
  });
}

// Preenche o campo de sugestões com todas as especialidades disponíveis.
function addSugestoes() {
  // Cria um conjunto para remover especialidades duplicadas.
  const especialidades = new Set(
    // Para cada professor, pega todas as especialidades e transforma em um único array.
    professors.flatMap((professor) =>
      professor.especialidades.map((especialidade) => especialidade.nome),
    ),
  );

  // Para cada especialidade única, cria um option dentro do select de sugestões.
  especialidades.forEach((nome) => {
    // Cria um elemento <option> para a lista de sugestões.
    const option = document.createElement('option');
    // Define o valor exibido e enviado pela opção.
    option.value = nome;
    // Adiciona a opção ao contêiner de sugestões.
    sugestoes.append(option);
  });
}

// Escuta o evento de envio do formulário para executar a busca.
form.addEventListener('submit', (event) => {
  // Impede o comportamento padrão do formulário de recarregar a página.
  event.preventDefault();
  // Chama a função que mostra os resultados relevantes.
  exibirResultados();
});

async function start() {
    const response = await fetch(WORKBOOK_URL);
    // Converte a resposta em um array de bytes para poder ler o workbook.
    const data = await response.arrayBuffer();
    // Lê o arquivo Excel em memória usando a biblioteca XLSX.
    const workbook = XLSX.read(data, { type: 'array' });
    // Armazena os professores carregados no array global.
    professors = readProfessors(workbook);

    addSugestoes();
    exibirResultados();
}

//Executa a função principal ao acessar a página
start();
