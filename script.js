(() => {
  const fileInput = document.getElementById('fileInput');
  const dropzone = document.getElementById('dropzone');
  const fileListBody = document.getElementById('fileListBody');
  const emptyRow = document.getElementById('emptyRow');
  const fileCount = document.getElementById('fileCount');
  const clearBtn = document.getElementById('clearBtn');
  const convertBtn = document.getElementById('convertBtn');
  const downloadZipBtn = document.getElementById('downloadZipBtn');
  const overallProgressWrap = document.getElementById('overallProgressWrap');
  const overallProgressFill = document.getElementById('overallProgressFill');
  const overallProgressLabel = document.getElementById('overallProgressLabel');
  const renderSandbox = document.getElementById('renderSandbox');

  // queue item shape: { id, file, name, size, status, pdfBlob, rowEl, errorMsg }
  let queue = [];
  let idSeq = 0;
  let isConverting = false;

  const STATUS_LABELS = {
    waiting: 'Aguardando',
    converting: 'Convertendo…',
    done: 'Concluído',
    error: 'Erro',
  };

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function pdfNameFor(originalName) {
    return originalName.replace(/\.docx$/i, '') + '.pdf';
  }

  function addFiles(fileList) {
    const docxFiles = Array.from(fileList).filter(f =>
      f.name.toLowerCase().endsWith('.docx')
    );
    docxFiles.forEach(file => {
      const item = {
        id: ++idSeq,
        file,
        name: file.name,
        size: file.size,
        status: 'waiting',
        pdfBlob: null,
        errorMsg: '',
      };
      queue.push(item);
      renderRow(item);
    });
    updateToolbar();
  }

  function renderRow(item) {
    emptyRow.hidden = true;
    const tr = document.createElement('tr');
    tr.dataset.id = item.id;
    tr.innerHTML = `
      <td class="file-name">${escapeHtml(item.name)}</td>
      <td class="file-size">${formatSize(item.size)}</td>
      <td><span class="status-badge status-${item.status}">${STATUS_LABELS[item.status]}</span></td>
      <td>
        <div class="row-actions">
          <button class="icon-btn download" data-action="download" disabled>Baixar PDF</button>
          <button class="icon-btn" data-action="remove">Remover</button>
        </div>
      </td>
    `;
    fileListBody.appendChild(tr);
    item.rowEl = tr;

    tr.querySelector('[data-action="remove"]').addEventListener('click', () => removeItem(item.id));
    tr.querySelector('[data-action="download"]').addEventListener('click', () => downloadSingle(item));
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function updateRowStatus(item) {
    const badge = item.rowEl.querySelector('.status-badge');
    badge.className = `status-badge status-${item.status}`;
    badge.textContent = item.status === 'error' && item.errorMsg
      ? `Erro: ${item.errorMsg}`
      : STATUS_LABELS[item.status];
    const downloadBtn = item.rowEl.querySelector('[data-action="download"]');
    downloadBtn.disabled = item.status !== 'done';
  }

  function removeItem(id) {
    queue = queue.filter(i => i.id !== id);
    const row = fileListBody.querySelector(`tr[data-id="${id}"]`);
    if (row) row.remove();
    if (queue.length === 0) emptyRow.hidden = false;
    updateToolbar();
  }

  function updateToolbar() {
    fileCount.textContent = `${queue.length} arquivo${queue.length === 1 ? '' : 's'}`;
    clearBtn.disabled = queue.length === 0 || isConverting;
    convertBtn.disabled = queue.length === 0 || isConverting;
    const anyDone = queue.some(i => i.status === 'done');
    downloadZipBtn.disabled = !anyDone || isConverting;
  }

  function updateOverallProgress(done, total) {
    if (total === 0) {
      overallProgressWrap.hidden = true;
      return;
    }
    overallProgressWrap.hidden = false;
    const pct = Math.round((done / total) * 100);
    overallProgressFill.style.width = `${pct}%`;
    overallProgressLabel.textContent = `${done} / ${total}`;
  }

  async function convertFileToPdf(item) {
    const arrayBuffer = await item.file.arrayBuffer();

    renderSandbox.innerHTML = '';
    const styleContainer = document.createElement('div');
    const bodyContainer = document.createElement('div');
    renderSandbox.appendChild(styleContainer);
    renderSandbox.appendChild(bodyContainer);

    await window.docx.renderAsync(arrayBuffer, bodyContainer, styleContainer, {
      inWrapper: true,
      ignoreLastRenderedPageBreak: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      experimental: true,
    });

    const pageSections = bodyContainer.querySelectorAll('section.docx');
    if (pageSections.length === 0) {
      renderSandbox.innerHTML = '';
      throw new Error('não foi possível interpretar o documento');
    }

    if (document.fonts && document.fonts.ready) {
      await document.fonts.ready;
    }

    const { jsPDF } = window.jspdf;
    let doc = null;

    for (const section of pageSections) {
      const widthPx = Math.ceil(section.getBoundingClientRect().width);
      const heightPx = Math.ceil(section.getBoundingClientRect().height);
      const widthPt = widthPx * 0.75;
      const heightPt = heightPx * 0.75;

      const canvas = await html2canvas(section, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        windowWidth: widthPx,
        windowHeight: heightPx,
      });
      const imgData = canvas.toDataURL('image/jpeg', 0.92);

      if (!doc) {
        doc = new jsPDF({ unit: 'pt', format: [widthPt, heightPt] });
      } else {
        doc.addPage([widthPt, heightPt]);
      }
      doc.addImage(imgData, 'JPEG', 0, 0, widthPt, heightPt);
    }

    renderSandbox.innerHTML = '';
    return doc.output('blob');
  }

  async function convertAll() {
    if (isConverting) return;
    isConverting = true;
    updateToolbar();

    const pending = queue.filter(i => i.status === 'waiting' || i.status === 'error');
    let done = 0;
    updateOverallProgress(done, pending.length);

    for (const item of pending) {
      item.status = 'converting';
      item.errorMsg = '';
      updateRowStatus(item);
      try {
        item.pdfBlob = await convertFileToPdf(item);
        item.status = 'done';
      } catch (err) {
        item.status = 'error';
        item.errorMsg = err.message || 'falha na conversão';
        console.error(`Erro convertendo ${item.name}:`, err);
      }
      updateRowStatus(item);
      done++;
      updateOverallProgress(done, pending.length);
    }

    isConverting = false;
    updateToolbar();
  }

  function downloadSingle(item) {
    if (!item.pdfBlob) return;
    saveAs(item.pdfBlob, pdfNameFor(item.name));
  }

  async function downloadZip() {
    const doneItems = queue.filter(i => i.status === 'done' && i.pdfBlob);
    if (doneItems.length === 0) return;

    const zip = new JSZip();
    const usedNames = new Set();
    doneItems.forEach(item => {
      let name = pdfNameFor(item.name);
      let counter = 1;
      while (usedNames.has(name)) {
        name = pdfNameFor(item.name).replace(/\.pdf$/, `-${counter}.pdf`);
        counter++;
      }
      usedNames.add(name);
      zip.file(name, item.pdfBlob);
    });

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    saveAs(zipBlob, 'documentos-convertidos.zip');
  }

  // events
  fileInput.addEventListener('change', e => {
    addFiles(e.target.files);
    fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach(evt => {
    dropzone.addEventListener(evt, e => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, e => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', e => {
    if (e.dataTransfer?.files?.length) addFiles(e.dataTransfer.files);
  });

  clearBtn.addEventListener('click', () => {
    queue = [];
    fileListBody.innerHTML = '';
    fileListBody.appendChild(emptyRow);
    emptyRow.hidden = false;
    updateOverallProgress(0, 0);
    updateToolbar();
  });

  convertBtn.addEventListener('click', convertAll);
  downloadZipBtn.addEventListener('click', downloadZip);

  updateToolbar();
})();
