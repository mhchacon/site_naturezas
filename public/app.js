const form = document.querySelector('#validator-form');
const fileInput = document.querySelector('#file-input');
const dropZone = document.querySelector('#drop-zone');
const fileLabel = document.querySelector('#file-label');
const message = document.querySelector('#form-message');
const menuButton = document.querySelector('.menu-button');
const nav = document.querySelector('.nav');

menuButton?.addEventListener('click', () => {
  const isOpen = nav.classList.toggle('open');
  menuButton.setAttribute('aria-expanded', String(isOpen));
});

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' || event.key === ' ') fileInput.click();
});
fileInput.addEventListener('change', () => updateFileLabel(fileInput.files[0]));
['dragenter', 'dragover'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  dropZone.classList.add('dragging');
}));
['dragleave', 'drop'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  dropZone.classList.remove('dragging');
}));
dropZone.addEventListener('drop', (event) => {
  const [file] = event.dataTransfer.files;
  if (file) {
    fileInput.files = event.dataTransfer.files;
    updateFileLabel(file);
  }
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file) return setMessage('Escolha uma planilha antes de continuar.', true);

  const formData = new FormData(form);
  const submitButton = form.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.firstChild.textContent = 'Conferindo...';
  setMessage('');

  try {
    const response = await fetch('/api/validate', { method: 'POST', body: formData });
    if (!response.ok) {
      const result = await response.json();
      throw new Error(result.error || 'Não foi possível validar o arquivo.');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = getDownloadName(file.name);
    link.click();
    URL.revokeObjectURL(url);
    const checked = response.headers.get('X-Checked-Rows');
    const invalid = response.headers.get('X-Invalid-Rows');
    setMessage(`${checked} linhas conferidas. ${invalid} ${invalid === '1' ? 'linha marcada' : 'linhas marcadas'} em amarelo.`);
  } catch (error) {
    setMessage(error.message, true);
  } finally {
    submitButton.disabled = false;
    submitButton.firstChild.textContent = 'Conferir planilha ';
  }
});

function updateFileLabel(file) {
  if (!file) return;
  fileLabel.textContent = file.name;
  setMessage('Arquivo selecionado. Agora informe as duas colunas.');
}
function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle('error', isError);
}
function getDownloadName(name) {
  return `${name.replace(/\.(xlsx|xls)$/i, '')}-validada.xlsx`;
}
