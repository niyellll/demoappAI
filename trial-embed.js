import { mountTrial } from './trial-data.js';

const host = document.getElementById('trial-data');

async function renderTrial() {
  const response = await fetch('/trial.html');
  if (!response.ok) throw new Error('Halaman trial belum tersedia.');
  const source = new DOMParser().parseFromString(await response.text(), 'text/html');
  const workspace = source.querySelector('main.workspace');
  const css = source.querySelector('link[rel="stylesheet"]');
  const parser = source.querySelector('script[integrity]');
  if (!workspace || !css || !parser) throw new Error('Komponen trial belum lengkap.');

  workspace.classList.add('workspace--embedded');
  const oldHomeLink = workspace.querySelector('.masthead a');
  oldHomeLink.remove();
  const title = workspace.querySelector('.masthead h1');
  const heading = source.createElement('h2');
  heading.className = 'workspace-title';
  heading.textContent = title.textContent;
  title.replaceWith(heading);
  const content = source.createElement('div');
  content.className = workspace.className;
  content.append(...workspace.childNodes);

  const container = document.createElement('div');
  const root = container.attachShadow({ mode: 'open' });
  root.append(css, content);
  mountTrial(root);
  for (const link of root.querySelectorAll('.jump-links a[href^="#"]')) {
    link.addEventListener('click', event => {
      event.preventDefault();
      const panel = root.getElementById(link.hash.slice(1));
      panel?.scrollIntoView({ block: 'start' });
    });
  }
  host.replaceChildren(container);
  const library = document.createElement('script');
  library.src = parser.src;
  library.integrity = parser.integrity;
  library.crossOrigin = parser.crossOrigin;
  document.head.append(library);
}

renderTrial().catch(() => {
  const message = document.createElement('p');
  message.textContent = 'Bagian trial belum bisa dimuat di halaman ini.';
  host.prepend(message);
});
