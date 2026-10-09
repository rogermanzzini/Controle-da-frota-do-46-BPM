/**
 * main.js - Scripts auxiliares para o Sistema do 46º BPM (Bagé/RS)
 */

function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function handleFileSelect(input, titleId, descId, dropzoneId) {
    if (input.files && input.files.length > 0) {
        const file = input.files[0];
        const titleEl = document.getElementById(titleId);
        const descEl = document.getElementById(descId);
        const dropzoneEl = document.getElementById(dropzoneId);

        if (titleEl) {
            titleEl.innerHTML = `<i class="fa-solid fa-circle-check text-success me-1"></i> ${file.name}`;
        }
        if (descEl) {
            descEl.innerHTML = `<span class="badge bg-success-subtle text-success">${formatBytes(file.size)}</span> &bull; Pronto para cruzamento`;
        }
        if (dropzoneEl) {
            dropzoneEl.classList.add('has-file');
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // Inicializar tooltips do Bootstrap
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    // Animação de fade-out suave em mensagens de alerta após 6 segundos
    setTimeout(() => {
        const alerts = document.querySelectorAll('.alert-dismissible');
        alerts.forEach(al => {
            const bsAlert = bootstrap.Alert.getInstance(al);
            if (bsAlert) {
                bsAlert.close();
            }
        });
    }, 6000);

    // Suporte para Drag and Drop nos dropzones
    const dropzones = document.querySelectorAll('.upload-dropzone');
    dropzones.forEach(dropzone => {
        const input = dropzone.querySelector('input[type="file"]');
        if (!input) return;

        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.remove('dragover');
            }, false);
        });

        dropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                input.files = files;
                input.dispatchEvent(new Event('change'));
            }
        }, false);
    });

    // Feedback de carregamento nos formulários
    const dashForm = document.getElementById('dashUploadForm');
    if (dashForm) {
        dashForm.addEventListener('submit', () => {
            const btn = document.getElementById('btnDashSubmit');
            const spinner = document.getElementById('dashSpinner');
            if (btn) btn.disabled = true;
            if (spinner) spinner.style.display = 'block';
        });
    }

    const uploadForm = document.getElementById('uploadForm');
    if (uploadForm) {
        uploadForm.addEventListener('submit', () => {
            const btn = document.getElementById('btnSubmit');
            const spinner = document.getElementById('processandoSpinner');
            if (btn) btn.disabled = true;
            if (spinner) spinner.style.display = 'block';
        });
    }
});

function alternarModoPlacasDash(modo) {
    const boxPlanilha = document.getElementById('boxPlanilhaAlvoDash');
    const boxTexto = document.getElementById('boxTextoAlvoDash');
    const btnPlanilha = document.getElementById('btnModoPlanilhaDash');
    const btnTexto = document.getElementById('btnModoTextoDash');

    if (modo === 'planilha') {
        boxPlanilha.style.display = 'block';
        boxTexto.style.display = 'none';
        btnPlanilha.classList.add('active');
        btnTexto.classList.remove('active');
    } else {
        boxPlanilha.style.display = 'none';
        boxTexto.style.display = 'block';
        btnTexto.classList.add('active');
        btnPlanilha.classList.remove('active');
        const txt = document.getElementById('dash_texto_placas');
        if (txt) txt.focus();
    }
}

function alternarModoPlacasUpload(modo) {
    const boxPlanilha = document.getElementById('boxPlanilhaAlvoUpload');
    const boxTexto = document.getElementById('boxTextoAlvoUpload');
    const btnPlanilha = document.getElementById('btnModoPlanilhaUpload');
    const btnTexto = document.getElementById('btnModoTextoUpload');

    if (modo === 'planilha') {
        boxPlanilha.style.display = 'block';
        boxTexto.style.display = 'none';
        btnPlanilha.classList.add('active');
        btnTexto.classList.remove('active');
    } else {
        boxPlanilha.style.display = 'none';
        boxTexto.style.display = 'block';
        btnTexto.classList.add('active');
        btnPlanilha.classList.remove('active');
        const txt = document.getElementById('upload_texto_placas');
        if (txt) txt.focus();
    }
}
