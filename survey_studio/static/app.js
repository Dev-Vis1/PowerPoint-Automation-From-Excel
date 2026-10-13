const fileInput = document.querySelector('#workbook');
const dropzone = document.querySelector('#dropzone');
const fileSelection = document.querySelector('#file-selection');
const fileName = document.querySelector('#file-name');
const fileSize = document.querySelector('#file-size');
const sheetSelect = document.querySelector('#sheet');
const sheetHint = document.querySelector('#sheet-hint');
const generateButton = document.querySelector('#generate');
const formError = document.querySelector('#form-error');
const result = document.querySelector('#result');
let selectedFile = null;
let inspection = 0;
let uploadedTemplate = null;
let chosenTemplate = '';
let previewSlides = [];
let currentSlide = 0;
const templateInput = document.querySelector('#template-file');
const templateList = document.querySelector('#template-list');
const templateDialog = document.querySelector('#template-dialog');
const narrowScreen = window.matchMedia('(max-width: 720px)');

function placeTemplate() {
  const anchor = document.querySelector(narrowScreen.matches ? '#template-mobile-anchor' : '#template-desktop-anchor');
  anchor.after(document.querySelector('.template-panel'));
}

narrowScreen.addEventListener('change', placeTemplate);
window.addEventListener('resize', placeTemplate);
placeTemplate();

function renderTemplateList() {
  for (const button of templateList.querySelectorAll('button')) {
    button.setAttribute('aria-pressed', String(button.dataset.template === chosenTemplate && !uploadedTemplate));
  }
}

function renderPreview() {
  const slide = previewSlides[currentSlide];
  const canvas = document.querySelector('#template-slide');
  canvas.replaceChildren();
  canvas.style.backgroundColor = slide.background;
  for (const item of slide.items) {
    const element = document.createElement(item.kind === 'image' ? 'img' : 'div');
    element.className = item.kind === 'text' ? 'slide-item slide-text'
      : item.kind === 'image' ? 'slide-item slide-image' : 'slide-item slide-block';
    element.style.left = `${item.x}%`;
    element.style.top = `${item.y}%`;
    element.style.width = `${item.w}%`;
    element.style.height = `${item.h}%`;
    if (item.kind === 'text') {
      element.textContent = item.text;
      element.style.fontFamily = item.font || '';
      element.style.color = item.color || '';
    } else if (item.kind === 'image') {
      element.src = item.data;
      element.alt = 'Template logo';
    }
    else element.style.backgroundColor = item.color;
    canvas.append(element);
  }
  document.querySelector('#slide-position').textContent = `Slide ${currentSlide + 1} of ${previewSlides.length}`;
  document.querySelector('#previous-slide').disabled = currentSlide === 0;
  document.querySelector('#next-slide').disabled = currentSlide === previewSlides.length - 1;
}

async function showPreview(name, source) {
  try {
    const response = await fetch(source.url, source.options);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || 'Cannot preview template.');
    previewSlides = payload.slides;
    currentSlide = 0;
    document.querySelector('#preview-title').textContent = name;
    if (!previewSlides.length) {
      previewSlides = [{ background: '#ffffff', items: [] }];
    }
    renderPreview();
    templateDialog.showModal();
  } catch (reason) {
    error(reason.message || 'Cannot preview template.');
  }
}

async function loadTemplates() {
  try {
    const response = await fetch('/api/templates');
    if (!response.ok) throw new Error('Templates are unavailable.');
    const templates = await response.json();
    templateList.replaceChildren();
    for (const template of templates) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'template-choice';
      button.dataset.template = template.id;
      button.innerHTML = '<span class="choice-icon"></span><span class="choice-name"></span><span class="choice-action"></span>';
      button.querySelector('.choice-icon').textContent = '▣';
      button.querySelector('.choice-name').textContent = template.name;
      button.querySelector('.choice-action').textContent = 'Preview';
      button.addEventListener('click', () => {
        chosenTemplate = template.id;
        uploadedTemplate = null;
        templateInput.value = '';
        document.querySelector('#uploaded-template').hidden = true;
        renderTemplateList();
        showPreview(template.name, { url: `/api/templates/${encodeURIComponent(template.id)}/preview` });
      });
      templateList.append(button);
    }
    renderTemplateList();
  } catch (reason) {
    document.querySelector('#template-hint').textContent = reason.message;
  }
}

templateInput.addEventListener('change', () => {
  const file = templateInput.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pptx') || file.size > 15 * 1024 * 1024) {
    templateInput.value = '';
    return error('Select a .pptx template no larger than 15 MB.');
  }
  error('');
  chosenTemplate = '';
  uploadedTemplate = file;
  document.querySelector('#uploaded-template-name').textContent = file.name;
  document.querySelector('#uploaded-template').hidden = false;
  renderTemplateList();
  const data = new FormData();
  data.append('file', file);
  showPreview(file.name, { url: '/api/templates/preview', options: { method: 'POST', body: data } });
});
document.querySelector('#preview-upload').addEventListener('click', () => {
  if (!uploadedTemplate) return;
  const data = new FormData();
  data.append('file', uploadedTemplate);
  showPreview(uploadedTemplate.name, { url: '/api/templates/preview', options: { method: 'POST', body: data } });
});
document.querySelector('#clear-template').addEventListener('click', () => {
  uploadedTemplate = null;
  templateInput.value = '';
  document.querySelector('#uploaded-template').hidden = true;
  renderTemplateList();
});
document.querySelector('#close-preview').addEventListener('click', () => templateDialog.close());
document.querySelector('#previous-slide').addEventListener('click', () => { currentSlide--; renderPreview(); });
document.querySelector('#next-slide').addEventListener('click', () => { currentSlide++; renderPreview(); });
templateDialog.addEventListener('click', event => { if (event.target === templateDialog) templateDialog.close(); });

function icons() {
  if (window.lucide) window.lucide.createIcons();
}

function error(message) {
  formError.textContent = message;
  formError.hidden = !message;
}

function reset() {
  selectedFile = null;
  fileInput.value = '';
  fileSelection.hidden = true;
  sheetSelect.disabled = true;
  sheetSelect.replaceChildren(new Option('Select a workbook first', ''));
  sheetHint.textContent = 'Choose the sheet containing the responses you want to present.';
  generateButton.disabled = true;
  result.hidden = true;
  error('');
}

async function choose(file) {
  inspection += 1;
  const current = inspection;
  reset();
  if (!file) return;
  if (!/\.(xlsx|csv)$/i.test(file.name)) return error('Please select an .xlsx or .csv survey file.');
  if (file.size > 15 * 1024 * 1024) return error('Workbook exceeds the 15 MB limit.');
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = `${(file.size / 1024 / 1024).toFixed(2)} MB`;
  fileSelection.hidden = false;
  sheetSelect.replaceChildren(new Option('Reading worksheets...', ''));
  sheetHint.textContent = 'Inspecting your workbook';
  const data = new FormData();
  data.append('file', file);
  try {
    const response = await fetch('/api/inspect', { method: 'POST', body: data });
    const payload = await response.json();
    if (current !== inspection) return;
    if (!response.ok) throw new Error(payload.detail || 'Unable to read the workbook.');
    sheetSelect.replaceChildren(...payload.sheets.map(name => new Option(name, name)));
    sheetSelect.value = payload.suggested;
    sheetSelect.disabled = false;
    sheetHint.textContent = `${payload.sheets.length} data ${payload.sheets.length === 1 ? 'sheet' : 'sheets'} available`;
    generateButton.disabled = false;
  } catch (reason) {
    if (current === inspection) {
      sheetHint.textContent = 'Unable to inspect workbook';
      error(reason.message || 'Unable to inspect workbook.');
    }
  }
}

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', event => {
  if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileInput.click(); }
});
fileInput.addEventListener('change', () => choose(fileInput.files[0]));
document.querySelector('#remove-file').addEventListener('click', () => { inspection++; reset(); });
for (const eventName of ['dragenter', 'dragover']) {
  dropzone.addEventListener(eventName, event => { event.preventDefault(); dropzone.classList.add('dragging'); });
}
for (const eventName of ['dragleave', 'drop']) {
  dropzone.addEventListener(eventName, event => { event.preventDefault(); dropzone.classList.remove('dragging'); });
}
dropzone.addEventListener('drop', event => choose(event.dataTransfer.files[0]));
document.querySelector('#create-form').addEventListener('submit', async event => {
  event.preventDefault();
  if (!selectedFile || !sheetSelect.value) return error('Choose a workbook and worksheet.');
  error('');
  result.hidden = true;
  generateButton.disabled = true;
  const label = generateButton.querySelector('span');
  label.textContent = 'Generating presentation...';
  const data = new FormData();
  data.append('file', selectedFile);
  data.append('sheet', sheetSelect.value);
  data.append('ai', document.querySelector('#use-ai').checked ? 'true' : 'false');
  if (uploadedTemplate) data.append('template_file', uploadedTemplate);
  else if (chosenTemplate) data.append('template_id', chosenTemplate);
  try {
    const response = await fetch('/api/generate', { method: 'POST', body: data });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.detail || 'Generation failed.');
    document.querySelector('#result-title').textContent = payload.filename;
    document.querySelector('#result-details').textContent = `${payload.respondents} responses  ·  ${payload.slides} editable slides  ·  ${payload.charts} native charts  ·  Available for 7 days`;
    document.querySelector('#download').href = payload.download;
    result.hidden = false;
    result.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } catch (reason) {
    error(reason.message || 'Generation failed. Please try again.');
  } finally {
    generateButton.disabled = false;
    label.textContent = 'Generate presentation';
  }
});
icons();
loadTemplates();