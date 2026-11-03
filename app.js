const WORKBOOK_URL =
  //'https://raw.githubusercontent.com/PedroSen/busque-seu-orientador/main/Tags%20TCCs%20IC(1).xlsx';
const form = document.querySelector('#search-form');
const searchInput = document.querySelector('#specialty-search');
const suggestions = document.querySelector('#specialty-options');
const results = document.querySelector('#results');
const resultsCount = document.querySelector('#results-count');
const status = document.querySelector('#status');

let professors = [];

function normalize(text) {
  return String(text)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function readProfessors(workbook) {
  return workbook.SheetNames.map((name) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      defval: '',
    });

    const specialties = rows
      .map(([name, weight], index) => {
        const label = String(name).trim();
        const score = Number(String(weight).replace(',', '.'));
        const isHeader = index === 0 && /especialidade/i.test(label);

        if (!label || isHeader || !Number.isFinite(score) || score <= 1) {
          return null;
        }

        return { name: label, normalizedName: normalize(label), weight: score };
      })
      .filter(Boolean);

    return { name, specialties };
  }).filter((professor) => professor.specialties.length > 0);
}

function getSearchTerms() {
  return searchInput.value
    .split(',')
    .map(normalize)
    .filter(Boolean);
}

function findMatches(terms) {
  return professors
    .map((professor) => {
      const specialties = professor.specialties.filter((specialty) =>
        terms.some((term) => specialty.normalizedName.includes(term)),
      );
      const score = specialties.reduce((total, item) => total + item.weight, 0);

      return { ...professor, matchedSpecialties: specialties, score };
    })
    .filter((professor) => terms.length === 0 || professor.score > 0)
    .sort((a, b) => b.score - a.score || a.name.localeCompare(b.name, 'pt-BR'));
}

function showStatus(message, isError = false) {
  status.hidden = !message;
  status.textContent = message;
  status.classList.toggle('error', isError);
}

function createCard(professor, terms) {
  const card = document.createElement('article');
  const specialties = terms.length
    ? professor.matchedSpecialties
    : professor.specialties.slice(0, 4);
  const tags = specialties
    .map((item) => `<span class="specialty-tag">${item.name}</span>`)
    .join('');
  const score = terms.length
    ? `<span class="match-score">Peso ${professor.score}</span>`
    : '';

  card.className = 'professor-card';
  card.innerHTML = `
    <div class="card-heading">
      <div>
        <h3>${professor.name}</h3>
        <p class="card-label" style="margin: 10px 0 0 2px">PROFESSOR(A) TITULAR</p>
      </div>
      ${score}
    </div>
    <div class="specialties">${tags}</div>
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

  return card;
}

function renderResults() {
  const terms = getSearchTerms();
  const matches = findMatches(terms);

  results.replaceChildren();
  resultsCount.textContent = terms.length
    ? `${matches.length} ${matches.length === 1 ? 'resultado' : 'resultados'}`
    : `${matches.length} ${matches.length === 1 ? 'professor' : 'professores'}`;

  if (matches.length === 0) {
    showStatus('Nenhum professor encontrado. Tente buscar por outra especialidade.');
    return;
  }

  showStatus('');
  matches.forEach((professor) => {
    results.append(createCard(professor, terms));
  });
}

function addSuggestions() {
  const specialties = new Set(
    professors.flatMap((professor) =>
      professor.specialties.map((specialty) => specialty.name),
    ),
  );

  specialties.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    suggestions.append(option);
  });
}

form.addEventListener('submit', (event) => {
  event.preventDefault();
  renderResults();
});

async function start() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error('Não foi possível carregar a planilha.');

    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    professors = readProfessors(workbook);

    addSuggestions();
    renderResults();
  } catch (error) {
    console.error(error);
    showStatus('Não foi possível carregar os professores. Tente novamente mais tarde.', true);
  }
}

start();
