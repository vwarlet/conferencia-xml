import './style.css';
import * as conferencia from './tools/conferencia.js';

// Para criar uma ferramenta nova: crie src/tools/nome.js com render(container) e adicione uma linha aqui.
const tools = [
  { id: 'conferencia', nome: 'Conferência de XMLs fiscais', desc: 'NF-e e NFC-e em planilha, com conferência do lote.', render: conferencia.render },
];

const app = document.getElementById('app');

// Segue o tema do sistema até a pessoa escolher uma preferência manual.
const root = document.documentElement;
const temaSistema = window.matchMedia('(prefers-color-scheme: dark)');
let temaManual = false;
try {
  const salvo = localStorage.getItem('tema');
  temaManual = salvo === 'light' || salvo === 'dark';
  root.dataset.theme = salvo || (temaSistema.matches ? 'dark' : 'light');
} catch { root.dataset.theme = temaSistema.matches ? 'dark' : 'light'; }
const escuro = () => root.dataset.theme === 'dark';
const iconeTema = (darkMode) => darkMode
  ? '<svg viewBox="0 0 24 24" aria-hidden="true"><circle class="sol-contorno" cx="12" cy="12" r="4"></circle><path class="sol-contorno" d="M12 2v2.5M12 19.5V22M4.93 4.93l1.77 1.77M17.3 17.3l1.77 1.77M2 12h2.5M19.5 12H22M4.93 19.07l1.77-1.77M17.3 6.7l1.77-1.77"></path></svg>'
  : '<svg viewBox="0 0 24 24" aria-hidden="true"><path class="lua-fill" d="M21 12.8A9 9 0 0 1 11.2 3a9 9 0 1 0 9.8 9.8Z"></path></svg>';
const topo = document.createElement('header');
topo.className = 'topo';
topo.innerHTML = '<button class="inicio" type="button"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-6v-7h-4v7H4a1 1 0 0 1-1-1V10Z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Início</span></button><button class="tema" type="button" aria-label="Alternar tema"></button>';
const btnInicio = topo.querySelector('.inicio');
const btn = topo.querySelector('.tema');
const rotulo = () => {
  const darkMode = escuro();
  btn.innerHTML = iconeTema(darkMode);
  btn.setAttribute('aria-label', darkMode ? 'Ativar tema claro' : 'Ativar tema escuro');
  btn.title = darkMode ? 'Tema claro' : 'Tema escuro';
};
btn.onclick = () => {
  temaManual = true;
  root.dataset.theme = escuro() ? 'light' : 'dark';
  try { localStorage.setItem('tema', root.dataset.theme); } catch { /* sem armazenamento */ }
  rotulo();
};
rotulo();
temaSistema.addEventListener('change', (e) => {
  if (!temaManual) {
    root.dataset.theme = e.matches ? 'dark' : 'light';
    rotulo();
  }
});
document.body.insertBefore(topo, app);

function route() {
  app.innerHTML = '';
  const main = document.createElement('main');
  const tool = tools.length === 1 ? tools[0] : tools.find((t) => t.id === location.hash.slice(1));
  if (tool) {
    if (tools.length > 1) app.insertAdjacentHTML('beforeend', '<a class="voltar" href="#">Todas as ferramentas</a>');
    app.append(main);
    tool.render(main);
  } else {
    main.innerHTML = `<h1>Ferramentas para XML fiscal</h1><ul class="lista">${tools.map((t) => `<li><a href="#${t.id}"><strong>${t.nome}</strong><span>${t.desc}</span></a></li>`).join('')}</ul>`;
    app.append(main);
  }
}
window.addEventListener('hashchange', route);
btnInicio.onclick = route;
route();
