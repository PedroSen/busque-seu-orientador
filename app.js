const WORKBOOK_URL = 'https://raw.githubusercontent.com/PedroSen/busque-seu-orientador/main/Tags%20TCCs%20IC(1).xlsx';
const searchInput = document.querySelector('#specialty-search');
const searchForm = document.querySelector('#search-form');
const results = document.querySelector('#results');
const statusBox = document.querySelector('#status');
const resultsCount = document.querySelector('#results-count');
const options = document.querySelector('#specialty-options');
const popularTopics = document.querySelector('#popular-topics');
let professors = [];
let allSpecialties = [];

function normalize(value) {
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('pt-BR').trim();
}

function parseWeight(value) {
  if (typeof value === 'number') return value;
  const parsed = Number(String(value ?? '').trim().replace(',', '.'));
  return Number.isFinite(parsed) ? parsed : 0;
}

function readWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  return workbook.SheetNames.map((name) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' });
    const specialties = rows.flatMap((row, index) => {
      const label = String(row[0] ?? '').trim();
      const weight = parseWeight(row[1]);
      if (!label || weight <= 1) return [];
      // Permite que a planilha tenha uma linha de cabeçalho, sem tratar "Especialidade" como tag.
      if (index === 0 && /especialidade|specialty/i.test(label)) return [];
      return [{ name: label, normalized: normalize(label), weight }];
    });
    return { name: name.trim(), specialties };
  }).filter((professor) => professor.name && professor.specialties.length);
}

function setStatus(message, kind = '', visible = true) {
  statusBox.className = `status ${kind}`.trim();
  statusBox.innerHTML = kind === 'loading' ? '<span class="loader" aria-hidden="true"></span>' : '';
  statusBox.append(document.createTextNode(message));
  statusBox.hidden = !visible;
}

function queryParts(query) {
  return query.split(',').map(normalize).filter(Boolean);
}

function matchesQuery(specialty, parts) {
  return parts.some((part) => {
    if (specialty.normalized.includes(part)) return true;
    const words = part.split(/\s+/).filter(Boolean);
    return words.length > 1 && words.every((word) => specialty.normalized.includes(word));
  });
}

function initials(name) {
  const words = name.split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words[0][0] + words[words.length - 1][0] : name.slice(0, 2)).toUpperCase();
}

function escapeHTML(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
}

function renderCards(query = '') {
  const parts = queryParts(query);
  const ranked = professors.map((professor) => {
    const matched = parts.length ? professor.specialties.filter((specialty) => matchesQuery(specialty, parts)) : [];
    const score = matched.reduce((sum, specialty) => sum + specialty.weight, 0);
    return { ...professor, matched, score };
  }).filter((professor) => !parts.length || professor.score > 0)
    .sort((a, b) => b.score - a.score || a.name.localeCompare(b.name, 'pt-BR'));

  results.replaceChildren();
  resultsCount.textContent = parts.length
    ? `${ranked.length} ${ranked.length === 1 ? 'professor encontrado' : 'professores encontrados'}`
    : `${professors.length} ${professors.length === 1 ? 'professor cadastrado' : 'professores cadastrados'}`;

  if (!ranked.length) {
    setStatus(parts.length ? 'Não encontramos professores para esse tema. Tente outro termo ou uma área mais ampla.' : 'Nenhum professor com especialidades válidas foi encontrado na planilha.', 'empty');
    return;
  }
  setStatus('', '', false);

  ranked.forEach((professor, index) => {
    const visibleTags = parts.length ? professor.matched : professor.specialties.slice(0, 3);
    const extraCount = parts.length ? 0 : Math.max(0, professor.specialties.length - visibleTags.length);
    const card = document.createElement('article');
    card.className = 'professor-card';
    card.style.animationDelay = `${Math.min(index * 45, 270)}ms`;
    card.innerHTML = `
      <div class="card-top">
        <div class="avatar" aria-hidden="true">${escapeHTML(initials(professor.name))}</div>
        <h3 class="professor-name">${escapeHTML(professor.name)}</h3>
        ${parts.length ? `<span class="match-score" title="Soma dos pesos das especialidades correspondentes">Relevância ${professor.score}</span>` : ''}
      </div>
      <p class="card-label">${parts.length ? 'ESPECIALIDADES CORRESPONDENTES' : 'ESPECIALIDADES EM DESTAQUE'}</p>
      <div class="tags">${visibleTags.map((tag) => `<span class="tag" title="Peso ${tag.weight}">${escapeHTML(tag.name)}</span>`).join('')}${extraCount ? `<span class="tag tag-more">+${extraCount} outras</span>` : ''}</div>`;
    results.append(card);
  });
}

function setSearch(query) {
  searchInput.value = query;
  renderCards(query);
  document.querySelector('#results-title').textContent = query.trim() ? 'Resultados da sua busca' : 'Professores orientadores';
  document.querySelector('#results-section')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function populateSuggestions() {
  const totals = new Map();
  professors.forEach((professor) => professor.specialties.forEach((specialty) => {
    const current = totals.get(specialty.normalized);
    if (!current) totals.set(specialty.normalized, { name: specialty.name, count: 1 });
    else current.count += 1;
  }));
  allSpecialties = [...totals.values()].sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, 'pt-BR'));
  options.replaceChildren(...allSpecialties.map((item) => {
    const option = document.createElement('option');
    option.value = item.name;
    return option;
  }));
  const popular = allSpecialties.slice(0, 5);
  if (popular.length) {
    popularTopics.hidden = false;
    popularTopics.innerHTML = `<span class="topic-label">Áreas populares:</span>`;
    popular.forEach((item) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'topic-chip';
      button.textContent = item.name;
      button.addEventListener('click', () => setSearch(item.name));
      popularTopics.append(button);
    });
  }
}

searchForm.addEventListener('submit', (event) => {
  event.preventDefault();
  renderCards(searchInput.value);
  document.querySelector('.results-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
});
searchInput.addEventListener('input', () => {
  if (!searchInput.value.trim()) renderCards('');
});

async function init() {
  try {
    if (!window.XLSX) throw new Error('A biblioteca de leitura da planilha não foi carregada. Verifique sua conexão e atualize a página.');
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Não foi possível carregar a planilha (HTTP ${response.status}).`);
    professors = readWorkbook(await response.arrayBuffer());
    if (!professors.length) throw new Error('A planilha foi carregada, mas não contém professores com especialidades de peso maior que 1.');
    populateSuggestions();
    renderCards();
  } catch (error) {
    console.error(error);
    setStatus(`${error.message} Confira se a planilha está disponível e tente novamente.`, 'error');
    resultsCount.textContent = '';
  }
}

init();
