// =============================================
// KONFIGURASI BACKEND URL
// Kosongkan "" jika berjalan satu server / lokal.
// Isi dengan URL Hugging Face Spaces saat frontend di-deploy di Vercel:
// Contoh: const API_BASE_URL = "https://username-space.hf.space";
// =============================================
const API_BASE_URL = "";

let currentMode = 'pdf-to-word';
let selectedFiles = [];
let imageFiles = [];
let draggedIndex = null;
let modalCallback = null;
let compressFile = null;
let compressImageFile = null;
let mergeMode = 'images';
let mergeImageFiles = [];
let mergePdfFiles = [];
let mergeDraggedIndex = null;
let watermarkFile = null;
let splitFile = null;
let splitMode = 'range';
let toastTimeout;

// Scan variables (CamScanner-Style 4-Corner Quadrilateral)
let scanFile = null;
let scanCorners = [
    { x: 0, y: 0 },
    { x: 0, y: 0 },
    { x: 0, y: 0 },
    { x: 0, y: 0 }
];
let scanImageWidth = 0;
let scanImageHeight = 0;
let activeCornerIndex = null;
let isDraggingQuad = false;
let dragStartX = 0;
let dragStartY = 0;
let initialCorners = [];

// PDF.js worker
if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ==========================================
// KATEGORI & NAVIGASI 2 TINGKAT
// ==========================================
const CATEGORY_TOOLS = {
    convert: ['pdf-to-word', 'word-to-pdf', 'pdf-to-markdown', 'word-to-markdown', 'image-to-pdf', 'pdf-to-image', 'ppt-to-pdf', 'pdf-to-ppt', 'excel-to-pdf', 'pdf-to-excel'],
    organize: ['merge-pdf', 'split-pdf', 'compress-pdf', 'compress-image', 'compress-image-target'],
    clean: ['remove-watermark', 'scan-document']
};
const TOOL_CATEGORY = {};
Object.entries(CATEGORY_TOOLS).forEach(([cat, tools]) => {
    tools.forEach(tool => { TOOL_CATEGORY[tool] = cat; });
});

// Mode yang memakai area upload generik #singleFileUpload (satu file masuk, satu file keluar)
const SINGLE_FILE_MODES = ['pdf-to-word', 'word-to-pdf', 'pdf-to-markdown', 'word-to-markdown', 'ppt-to-pdf', 'pdf-to-ppt', 'excel-to-pdf', 'pdf-to-excel'];
const SINGLE_FILE_ACCEPT = {
    'pdf-to-word': '.pdf',
    'word-to-pdf': '.docx,.doc',
    'pdf-to-markdown': '.pdf',
    'word-to-markdown': '.docx,.doc',
    'ppt-to-pdf': '.pptx,.ppt',
    'pdf-to-ppt': '.pdf',
    'excel-to-pdf': '.xlsx,.xls',
    'pdf-to-excel': '.pdf'
};

// State untuk fitur baru
let pdfToImageFile = null;
let compressImageTargetFile = null;

function switchCategory(cat) {
    switchTab(CATEGORY_TOOLS[cat][0]);
}

// Elemen area upload yang relevan untuk tiap mode - dipakai untuk auto-scroll di mobile
function getActiveUploadAreaId(mode) {
    const map = {
        'pdf-to-word': 'singleFileUpload',
        'word-to-pdf': 'singleFileUpload',
        'pdf-to-markdown': 'singleFileUpload',
        'word-to-markdown': 'singleFileUpload',
        'ppt-to-pdf': 'singleFileUpload',
        'pdf-to-ppt': 'singleFileUpload',
        'excel-to-pdf': 'singleFileUpload',
        'pdf-to-excel': 'singleFileUpload',
        'image-to-pdf': 'multipleFileUpload',
        'pdf-to-image': 'pdfToImageUpload',
        'compress-pdf': 'compressUpload',
        'split-pdf': 'splitUpload',
        'remove-watermark': 'watermarkUpload',
        'scan-document': 'scanUpload',
        'compress-image': 'compressImageUpload',
        'compress-image-target': 'compressImageTargetUpload'
    };
    if (mode === 'merge-pdf') {
        return mergeMode === 'images' ? 'mergeImagesUpload' : 'mergePdfsUpload';
    }
    return map[mode];
}

// Auto-scroll ke area drag & drop - hanya di layar sempit (mobile), supaya
// desktop (yang biasanya seluruh box sudah kelihatan) tidak ikut lompat
function scrollToActiveUploadArea(mode) {
    if (window.innerWidth > 768) return;
    const targetId = getActiveUploadAreaId(mode);
    const el = targetId && document.getElementById(targetId);
    if (el) {
        requestAnimationFrame(() => el.scrollIntoView({ behavior: 'smooth', block: 'start' }));
    }
}

// ==========================================
// MODAL SYSTEM
// ==========================================
function showModal(title, message, confirmText, type, callback) {
    const modal = document.getElementById('confirmModal');
    const modalBox = modal.querySelector('.modal-box');
    const modalTitle = document.getElementById('modalTitle');
    const modalMessage = document.getElementById('modalMessage');
    const confirmBtn = document.getElementById('modalConfirmBtn');

    modalTitle.textContent = title;
    modalMessage.innerHTML = message;
    confirmBtn.textContent = confirmText;
    confirmBtn.className = 'modal-btn modal-btn-confirm';

    if (type === 'danger') confirmBtn.classList.add('danger');
    else if (type === 'warning') confirmBtn.classList.add('warning');

    modalBox.dataset.type = type || 'info';

    modalCallback = callback;
    modal.classList.add('show');
    setTimeout(() => document.querySelector('.modal-btn-cancel')?.focus(), 50);
}

function closeModal() { document.getElementById('confirmModal').classList.remove('show'); modalCallback = null; }
function confirmModalAction() { if (typeof modalCallback === 'function') modalCallback(); closeModal(); }

// ==========================================
// TOAST SYSTEM
// ==========================================
function showToast(message, type = 'info', duration = 3000) {
    const toast = document.getElementById('toast');
    const toastLabel = document.getElementById('toastLabel');
    const toastMessage = document.getElementById('toastMessage');

    const labels = { success: 'Berhasil', error: 'Gagal', info: 'Info' };
    const className = type === 'success' ? 'success' : type === 'error' ? 'error' : 'info';

    toast.className = 'toast-notification';
    toast.style.display = 'none';
    toast.classList.add(className);

    toastLabel.textContent = labels[className];
    toastMessage.textContent = message;

    void toast.offsetWidth;
    toast.style.display = 'flex';
    clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() => { toast.style.display = 'none'; }, duration);
}

function showStatus(message, type) {
    if (type === 'success' || type === 'error' || type === 'info') showToast(message, type);
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = `status ${type}`;
    status.style.display = 'block';
    if (type === 'success') setTimeout(() => { status.style.display = 'none'; }, 4000);
}

// ==========================================
// TAB SWITCHING (kategori + sub-tool)
// ==========================================
function switchTab(mode) {
    currentMode = mode;
    const cat = TOOL_CATEGORY[mode];

    document.querySelectorAll('.category-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.category === cat);
    });
    document.querySelectorAll('.tool-group').forEach(group => {
        group.style.display = group.dataset.category === cat ? 'flex' : 'none';
    });

    const tabTexts = {
        'pdf-to-word': 'PDF → Word',
        'word-to-pdf': 'Word → PDF',
        'pdf-to-markdown': 'PDF → Markdown (AI)',
        'word-to-markdown': 'Word → Markdown (AI)',
        'image-to-pdf': 'Image → PDF',
        'pdf-to-image': 'PDF → Gambar',
        'ppt-to-pdf': 'PowerPoint → PDF',
        'pdf-to-ppt': 'PDF → PowerPoint',
        'excel-to-pdf': 'Excel → PDF',
        'pdf-to-excel': 'PDF → Excel',
        'compress-pdf': 'Compress PDF',
        'compress-image': 'Compress Gambar',
        'compress-image-target': 'Standar Administrasi',
        'merge-pdf': 'Merge PDF',
        'split-pdf': 'Split PDF',
        'remove-watermark': 'Remove Watermark',
        'scan-document': 'Scan Dokumen'
    };
    document.querySelectorAll('.tool-tab').forEach(tab => {
        tab.classList.toggle('active', tab.textContent.trim() === tabTexts[mode]);
    });

    // Reset state
    selectedFiles = []; imageFiles = []; compressFile = null; watermarkFile = null; scanFile = null; splitFile = null;
    compressImageFile = null;
    pdfToImageFile = null;
    compressImageTargetFile = null;
    document.getElementById('singleFileInput').value = '';
    document.getElementById('multipleFileInput').value = '';
    document.getElementById('compressFileInput').value = '';
    document.getElementById('mergeImagesInput').value = '';
    document.getElementById('mergePdfsInput').value = '';
    document.getElementById('watermarkFileInput').value = '';
    document.getElementById('scanFileInput').value = '';
    document.getElementById('splitFileInput').value = '';
    document.getElementById('compressImageFileInput').value = '';
    document.getElementById('pdfToImageFileInput').value = '';
    document.getElementById('compressImageTargetFileInput').value = '';
    document.getElementById('singleFileInfo').textContent = '';
    document.getElementById('compressFileInfo').textContent = '';
    document.getElementById('watermarkFileInfo').textContent = '';
    document.getElementById('splitFileInfo').textContent = '';
    document.getElementById('compressImageFileInfo').textContent = '';
    document.getElementById('pdfToImageFileInfo').textContent = '';
    document.getElementById('compressImageTargetFileInfo').textContent = '';
    document.getElementById('convertBtn').disabled = true;
    document.getElementById('convertBtn').style.display = 'flex';
    document.getElementById('status').className = 'status';
    document.getElementById('status').style.display = 'none';
    document.getElementById('btnText').textContent = ['pdf-to-markdown', 'word-to-markdown'].includes(mode) ? 'Convert ke Markdown' : 'Convert Sekarang';
    document.getElementById('compressOptions').style.display = 'none';
    document.getElementById('compressResult').style.display = 'none';
    document.getElementById('compressImageOptions').style.display = 'none';
    document.getElementById('compressImageResult').style.display = 'none';
    document.getElementById('pdfConvertOptions').style.display = 'none';
    document.getElementById('markdownInfoBanner').style.display = ['pdf-to-markdown', 'word-to-markdown'].includes(mode) ? 'flex' : 'none';
    document.getElementById('markdownResult').style.display = 'none';
    document.getElementById('watermarkOptions').style.display = 'none';
    document.getElementById('scanPreviewContainer').style.display = 'none';
    document.getElementById('splitOptions').style.display = 'none';
    document.getElementById('pdfToImageOptions').style.display = 'none';
    document.getElementById('compressImageTargetOptions').style.display = 'none';
    document.getElementById('compressImageTargetResult').style.display = 'none';
    window.compressedImageBlob = null;
    window.compressedTargetBlob = null;
    window.currentMarkdownText = '';
    window.currentMarkdownBlob = null;

    // Show/hide containers
    document.getElementById('singleFileUpload').style.display = SINGLE_FILE_MODES.includes(mode) ? 'block' : 'none';
    document.getElementById('imageToPdfContainer').style.display = mode === 'image-to-pdf' ? 'block' : 'none';
    document.getElementById('pdfToImageContainer').style.display = mode === 'pdf-to-image' ? 'block' : 'none';
    document.getElementById('compressPdfContainer').style.display = mode === 'compress-pdf' ? 'block' : 'none';
    document.getElementById('mergePdfContainer').style.display = mode === 'merge-pdf' ? 'block' : 'none';
    document.getElementById('splitPdfContainer').style.display = mode === 'split-pdf' ? 'block' : 'none';
    document.getElementById('removeWatermarkContainer').style.display = mode === 'remove-watermark' ? 'block' : 'none';
    document.getElementById('scanContainer').style.display = mode === 'scan-document' ? 'block' : 'none';
    document.getElementById('compressImageContainer').style.display = mode === 'compress-image' ? 'block' : 'none';
    document.getElementById('compressImageTargetContainer').style.display = mode === 'compress-image-target' ? 'block' : 'none';
    document.getElementById('imagePreviewContainer').style.display = 'none';

    if (SINGLE_FILE_MODES.includes(mode)) {
        document.getElementById('singleFileInput').accept = SINGLE_FILE_ACCEPT[mode];
    }
    renderImageGrid();
    scrollToActiveUploadArea(mode);
}

// ==========================================
// FILE INPUT HANDLERS
// ==========================================
document.getElementById('singleFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleSingleFile(e.target.files[0]);
});

document.getElementById('multipleFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) { addImages(e.target.files); e.target.value = ''; }
});

document.getElementById('compressFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleCompressFile(e.target.files[0]);
});

document.getElementById('mergeImagesInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) { addMergeImages(e.target.files); e.target.value = ''; }
});

document.getElementById('mergePdfsInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) { addMergePdfs(e.target.files); e.target.value = ''; }
});

document.getElementById('watermarkFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleWatermarkFile(e.target.files[0]);
});

document.getElementById('scanFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleScanFile(e.target.files[0]);
});

document.getElementById('splitFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleSplitFile(e.target.files[0]);
});

document.getElementById('compressImageFileInput').addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleCompressImageFile(e.target.files[0]);
});

// ==========================================
// UPLOAD AREA CLICK
// ==========================================
function setupUploadAreaClick(uploadAreaId, inputId) {
    const uploadArea = document.getElementById(uploadAreaId);
    if (!uploadArea) return;
    const uploadContent = uploadArea.querySelector('.upload-content');
    if (!uploadContent) return;

    uploadContent.addEventListener('click', (e) => {
        const tagName = e.target.tagName.toLowerCase();
        if (['select', 'option', 'button', 'label', 'input'].includes(tagName)) return;
        document.getElementById(inputId).click();
    });
}

setupUploadAreaClick('singleFileUpload', 'singleFileInput');
setupUploadAreaClick('multipleFileUpload', 'multipleFileInput');
setupUploadAreaClick('compressUpload', 'compressFileInput');
setupUploadAreaClick('mergeImagesUpload', 'mergeImagesInput');
setupUploadAreaClick('mergePdfsUpload', 'mergePdfsInput');
setupUploadAreaClick('watermarkUpload', 'watermarkFileInput');
setupUploadAreaClick('scanUpload', 'scanFileInput');
setupUploadAreaClick('splitUpload', 'splitFileInput');
setupUploadAreaClick('compressImageUpload', 'compressImageFileInput');
setupUploadAreaClick('pdfToImageUpload', 'pdfToImageFileInput');
setupUploadAreaClick('compressImageTargetUpload', 'compressImageTargetFileInput');

// ==========================================
// DRAG & DROP
// ==========================================
function setupDragDrop(element, callback, multiple = false) {
    if (!element) return;
    element.addEventListener('dragover', (e) => { e.preventDefault(); element.classList.add('dragover'); });
    element.addEventListener('dragleave', () => element.classList.remove('dragover'));
    element.addEventListener('drop', (e) => {
        e.preventDefault();
        element.classList.remove('dragover');
        callback(multiple ? e.dataTransfer.files : e.dataTransfer.files[0]);
    });
}

setupDragDrop(document.getElementById('singleFileUpload'), (f) => handleSingleFile(f));
setupDragDrop(document.getElementById('multipleFileUpload'), (f) => addImages(f instanceof FileList ? f : [f]), true);
setupDragDrop(document.getElementById('compressUpload'), (f) => handleCompressFile(f));
setupDragDrop(document.getElementById('mergeImagesUpload'), (f) => addMergeImages(f instanceof FileList ? f : [f]), true);
setupDragDrop(document.getElementById('mergePdfsUpload'), (f) => addMergePdfs(f instanceof FileList ? f : [f]), true);
setupDragDrop(document.getElementById('watermarkUpload'), (f) => handleWatermarkFile(f));
setupDragDrop(document.getElementById('scanUpload'), (f) => handleScanFile(f));
setupDragDrop(document.getElementById('splitUpload'), (f) => handleSplitFile(f));
setupDragDrop(document.getElementById('compressImageUpload'), (f) => handleCompressImageFile(f));
setupDragDrop(document.getElementById('pdfToImageUpload'), (f) => handlePdfToImageFile(f));
setupDragDrop(document.getElementById('compressImageTargetUpload'), (f) => handleCompressImageTargetFile(f));

// ==========================================
// SINGLE FILE (PDF<->Word, PPT->PDF, PDF->PPT, Excel->PDF, PDF->Excel, Markdown)
// ==========================================
function handleSingleFile(file) {
    if (!file) return;
    selectedFiles = [file];
    document.getElementById('singleFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = ['pdf-to-markdown', 'word-to-markdown'].includes(currentMode) ? 'Convert ke Markdown' : 'Convert Sekarang';
    document.getElementById('pdfConvertOptions').style.display = currentMode === 'pdf-to-word' ? 'block' : 'none';
    document.getElementById('markdownInfoBanner').style.display = ['pdf-to-markdown', 'word-to-markdown'].includes(currentMode) ? 'flex' : 'none';
}

// ==========================================
// COMPRESS PDF
// ==========================================
function handleCompressFile(file) {
    if (!file) return;
    if (!file.name.endsWith('.pdf')) { showToast('File harus PDF!', 'error'); return; }
    if (file.size > 50 * 1024 * 1024) { showToast('File terlalu besar! Maksimal 50MB.', 'error'); return; }

    compressFile = file;
    document.getElementById('compressFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('compressOptions').style.display = 'block';
    document.getElementById('compressResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Kompres PDF';
}

document.querySelectorAll('.compress-level-card').forEach(card => {
    card.addEventListener('click', function() {
        document.querySelectorAll('.compress-level-card').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        this.querySelector('input[type="radio"]').checked = true;
    });
});

document.querySelectorAll('.compress-image-level-card').forEach(card => {
    card.addEventListener('click', function() {
        document.querySelectorAll('.compress-image-level-card').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        this.querySelector('input[type="radio"]').checked = true;
    });
});

function resetCompress() {
    compressFile = null;
    document.getElementById('compressFileInput').value = '';
    document.getElementById('compressFileInfo').textContent = '';
    document.getElementById('compressOptions').style.display = 'none';
    document.getElementById('compressResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = true;
    document.getElementById('btnText').textContent = 'Convert Sekarang';
    window.compressedBlob = null;
}

// ==========================================
// COMPRESS IMAGE
// ==========================================
function handleCompressImageFile(file) {
    if (!file) return;
    const validExts = ['.jpg', '.jpeg', '.png', '.webp'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) { showToast('File harus JPG, PNG, atau WEBP!', 'error'); return; }
    if (file.size > 30 * 1024 * 1024) { showToast('File terlalu besar! Maksimal 30MB.', 'error'); return; }

    compressImageFile = file;
    document.getElementById('compressImageFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('compressImageOptions').style.display = 'block';
    document.getElementById('compressImageResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Kompres Gambar';
}

function resetCompressImage() {
    compressImageFile = null;
    document.getElementById('compressImageFileInput').value = '';
    document.getElementById('compressImageFileInfo').textContent = '';
    document.getElementById('compressImageOptions').style.display = 'none';
    document.getElementById('compressImageResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = true;
    document.getElementById('btnText').textContent = 'Convert Sekarang';
    window.compressedImageBlob = null;
}

// ==========================================
// PDF TO IMAGE
// ==========================================
function handlePdfToImageFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.pdf')) { showToast('File harus PDF!', 'error'); return; }
    if (file.size > 50 * 1024 * 1024) { showToast('File terlalu besar! Maksimal 50MB.', 'error'); return; }

    pdfToImageFile = file;
    document.getElementById('pdfToImageFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('pdfToImageOptions').style.display = 'block';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Konversi ke Gambar';
}

document.getElementById('pdfToImagePages')?.addEventListener('change', function() {
    document.getElementById('pdfToImageCustomRange').style.display = this.value === 'custom' ? 'block' : 'none';
});

document.getElementById('pdfToImageFileInput')?.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handlePdfToImageFile(e.target.files[0]);
});

// ==========================================
// COMPRESS IMAGE — STANDAR ADMINISTRASI
// ==========================================
function handleCompressImageTargetFile(file) {
    if (!file) return;
    const validExts = ['.jpg', '.jpeg', '.png', '.webp'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) { showToast('File harus JPG, PNG, atau WEBP!', 'error'); return; }
    if (file.size > 30 * 1024 * 1024) { showToast('File terlalu besar! Maksimal 30MB.', 'error'); return; }

    compressImageTargetFile = file;
    document.getElementById('compressImageTargetFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('compressImageTargetOptions').style.display = 'block';
    document.getElementById('compressImageTargetResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Kompres ke Target';
}

document.getElementById('compressImageTargetFileInput')?.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleCompressImageTargetFile(e.target.files[0]);
});

document.querySelectorAll('.compress-target-level-card').forEach(card => {
    card.addEventListener('click', function() {
        document.querySelectorAll('.compress-target-level-card').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        this.querySelector('input[type="radio"]').checked = true;
        document.getElementById('customTargetInput').style.display = this.dataset.target === 'custom' ? 'block' : 'none';
    });
});

function resetCompressImageTarget() {
    compressImageTargetFile = null;
    document.getElementById('compressImageTargetFileInput').value = '';
    document.getElementById('compressImageTargetFileInfo').textContent = '';
    document.getElementById('compressImageTargetOptions').style.display = 'none';
    document.getElementById('compressImageTargetResult').style.display = 'none';
    document.getElementById('convertBtn').disabled = true;
    document.getElementById('btnText').textContent = 'Convert Sekarang';
    window.compressedTargetBlob = null;
}

// ==========================================
// SPLIT PDF
// ==========================================
function handleSplitFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.pdf')) { showToast('File harus PDF!', 'error'); return; }
    if (file.size > 50 * 1024 * 1024) { showToast('File terlalu besar! Maksimal 50MB.', 'error'); return; }

    splitFile = file;
    document.getElementById('splitFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('splitOptions').style.display = 'block';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Split PDF';
}

function switchSplitMode(mode) {
    splitMode = mode;
    document.querySelectorAll('#splitOptions .merge-mode-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('#splitOptions .merge-mode-tab').forEach(t => {
        if ((mode === 'range' && t.textContent.trim() === 'Berdasarkan Range') ||
            (mode === 'every' && t.textContent.trim() === 'Tiap N Halaman')) {
            t.classList.add('active');
        }
    });
    document.getElementById('splitRangeMode').style.display = mode === 'range' ? 'block' : 'none';
    document.getElementById('splitEveryMode').style.display = mode === 'every' ? 'block' : 'none';
}

// ==========================================
// REMOVE WATERMARK
// ==========================================
function handleWatermarkFile(file) {
    if (!file) return;

    const validExts = ['.pdf', '.jpg', '.jpeg', '.png'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();

    if (!validExts.includes(ext)) {
        showToast('File harus PDF, JPG, PNG, atau JPEG!', 'error');
        return;
    }

    if (file.size > 50 * 1024 * 1024) {
        showToast('File terlalu besar! Maksimal 50MB.', 'error');
        return;
    }

    watermarkFile = file;
    document.getElementById('watermarkFileInfo').textContent =
        `${file.name} (${formatFileSize(file.size)})`;
    document.getElementById('watermarkOptions').style.display = 'block';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Hapus Watermark';
}

// ==========================================
// SCAN DOCUMENT
// ==========================================
function handleScanFile(file) {
    if (!file) return;

    const validExts = ['.jpg', '.jpeg', '.png'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();

    if (!validExts.includes(ext)) {
        showToast('File harus JPG, PNG, atau JPEG!', 'error');
        return;
    }

    scanFile = file;
    document.getElementById('scanPreviewContainer').style.display = 'block';
    document.getElementById('convertBtn').disabled = false;
    document.getElementById('btnText').textContent = 'Scan & Crop';

    const reader = new FileReader();
    reader.onload = (e) => {
        const img = document.getElementById('scanPreviewImage');
        img.src = e.target.result;
        img.onload = () => {
            scanImageWidth = img.naturalWidth;
            scanImageHeight = img.naturalHeight;
            // Tunggu 2 frame agar browser selesai layout image di dalam flex container
            requestAnimationFrame(() => {
                requestAnimationFrame(() => {
                    resetCrop();
                    autoDetectCrop();
                });
            });
        };
    };
    reader.readAsDataURL(file);
}

function resetCrop() {
    const img = document.getElementById('scanPreviewImage');
    if (!img) return;
    const w = img.offsetWidth;
    const h = img.offsetHeight;
    if (w <= 0 || h <= 0) return;

    const padX = w * 0.06;
    const padY = h * 0.06;

    scanCorners = [
        { x: padX, y: padY },                 // 0: Top-Left
        { x: w - padX, y: padY },             // 1: Top-Right
        { x: w - padX, y: h - padY },         // 2: Bottom-Right
        { x: padX, y: h - padY }              // 3: Bottom-Left
    ];

    updateCropOverlay();
}

async function autoDetectCrop() {
    if (!scanFile) return;

    const formData = new FormData();
    formData.append('file', scanFile);

    try {
        const response = await fetch(`${API_BASE_URL}/detect-boundary`, {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            const data = await response.json();
            const img = document.getElementById('scanPreviewImage');
            const renderW = img.offsetWidth;
            const renderH = img.offsetHeight;

            const scaleX = renderW / data.image_width;
            const scaleY = renderH / data.image_height;

            if (data.corners && Array.isArray(data.corners) && data.corners.length === 4) {
                // Ekstrak 4 titik sudut asli (mendukung bentuk miring / jajar genjang / trapesium)
                scanCorners = data.corners.map(pt => ({
                    x: Math.max(0, Math.min(renderW, pt[0] * scaleX)),
                    y: Math.max(0, Math.min(renderH, pt[1] * scaleY))
                }));
            } else {
                // Fallback dari bounding box persegi
                const x1 = data.x * scaleX;
                const y1 = data.y * scaleY;
                const x2 = (data.x + data.width) * scaleX;
                const y2 = (data.y + data.height) * scaleY;
                scanCorners = [
                    { x: x1, y: y1 },
                    { x: x2, y: y1 },
                    { x: x2, y: y2 },
                    { x: x1, y: y2 }
                ];
            }

            updateCropOverlay();
            showToast('Sudut dokumen terdeteksi otomatis', 'success', 2000);
        }
    } catch (err) {
        showToast('Auto detect gagal, sesuaikan sudut secara manual', 'error');
    }
}

function updateCropOverlay() {
    const polygon = document.getElementById('scanCropPolygon');
    if (!polygon || !scanCorners || scanCorners.length !== 4) return;

    // Perbarui titik-titik pada poligon SVG
    const pointsStr = scanCorners.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
    polygon.setAttribute('points', pointsStr);

    // Perbarui posisi 4 titik kontrol sudut
    const cornerIds = ['cornerTL', 'cornerTR', 'cornerBR', 'cornerBL'];
    cornerIds.forEach((id, idx) => {
        const handle = document.getElementById(id);
        if (handle && scanCorners[idx]) {
            handle.style.left = `${scanCorners[idx].x}px`;
            handle.style.top = `${scanCorners[idx].y}px`;
        }
    });
}

// Interaksi geser sudut dan poligon (Mouse & Touch)
function getPointFromEvent(e) {
    if (e.touches && e.touches.length > 0) {
        return { x: e.touches[0].clientX, y: e.touches[0].clientY };
    }
    return { x: e.clientX, y: e.clientY };
}

function handleCornerDragStart(idx, e) {
    e.stopPropagation();
    activeCornerIndex = idx;
    isDraggingQuad = false;
    const pt = getPointFromEvent(e);
    dragStartX = pt.x;
    dragStartY = pt.y;
    initialCorners = scanCorners.map(p => ({ ...p }));
    if (e.cancelable) e.preventDefault();
}

function handlePolygonDragStart(e) {
    activeCornerIndex = null;
    isDraggingQuad = true;
    const pt = getPointFromEvent(e);
    dragStartX = pt.x;
    dragStartY = pt.y;
    initialCorners = scanCorners.map(p => ({ ...p }));
    if (e.cancelable) e.preventDefault();
}

function handleCropMove(e) {
    if (activeCornerIndex === null && !isDraggingQuad) return;
    const pt = getPointFromEvent(e);
    const dx = pt.x - dragStartX;
    const dy = pt.y - dragStartY;

    const img = document.getElementById('scanPreviewImage');
    const maxW = img ? img.offsetWidth : 9999;
    const maxH = img ? img.offsetHeight : 9999;

    if (activeCornerIndex !== null) {
        // Geser satu sudut bebas (bisa membentuk jajar genjang / trapesium)
        const orig = initialCorners[activeCornerIndex];
        scanCorners[activeCornerIndex] = {
            x: Math.max(0, Math.min(maxW, orig.x + dx)),
            y: Math.max(0, Math.min(maxH, orig.y + dy))
        };
        updateCropOverlay();
    } else if (isDraggingQuad) {
        // Geser seluruh bidang dokumen bersamaan
        let validDx = dx;
        let validDy = dy;
        initialCorners.forEach(p => {
            if (p.x + validDx < 0) validDx = -p.x;
            if (p.x + validDx > maxW) validDx = maxW - p.x;
            if (p.y + validDy < 0) validDy = -p.y;
            if (p.y + validDy > maxH) validDy = maxH - p.y;
        });

        scanCorners = initialCorners.map(p => ({
            x: p.x + validDx,
            y: p.y + validDy
        }));
        updateCropOverlay();
    }

    if (e.cancelable) e.preventDefault();
}

function handleCropEnd() {
    activeCornerIndex = null;
    isDraggingQuad = false;
}

document.addEventListener('DOMContentLoaded', () => {
    // Daftarkan listener untuk 4 titik sudut
    ['cornerTL', 'cornerTR', 'cornerBR', 'cornerBL'].forEach((id, idx) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('mousedown', (e) => handleCornerDragStart(idx, e));
            el.addEventListener('touchstart', (e) => handleCornerDragStart(idx, e), { passive: false });
        }
    });

    // Daftarkan listener pada area dalam poligon (untuk geser serempak)
    const polygon = document.getElementById('scanCropPolygon');
    if (polygon) {
        polygon.addEventListener('mousedown', handlePolygonDragStart);
        polygon.addEventListener('touchstart', handlePolygonDragStart, { passive: false });
    }

    document.addEventListener('mousemove', handleCropMove);
    document.addEventListener('touchmove', handleCropMove, { passive: false });
    document.addEventListener('mouseup', handleCropEnd);
    document.addEventListener('touchend', handleCropEnd);
});

// ==========================================
// IMAGE TO PDF
// ==========================================
function addImages(files) {
    const validExts = ['.jpg', '.jpeg', '.png'];
    let added = 0, invalid = [];

    for (let file of files) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (validExts.includes(ext)) {
            imageFiles.push({ file, rotation: 0, id: Date.now() + Math.random(), previewUrl: URL.createObjectURL(file) });
            added++;
        } else invalid.push(file.name);
    }

    if (invalid.length) showToast(`File tidak valid: ${invalid.join(', ')}`, 'error', 4000);
    if (added > 0) { updatePreview(); showToast(`${added} gambar ditambahkan`, 'success', 2000); }
}

function removeImage(index) {
    showModal('Hapus Gambar', `Hapus <span class="highlight-filename">${imageFiles[index].file.name}</span>?`, 'Hapus', 'danger', () => {
        URL.revokeObjectURL(imageFiles[index].previewUrl);
        imageFiles.splice(index, 1);
        updatePreview();
    });
}

function clearAllImages() {
    if (!imageFiles.length) return;
    showModal('Hapus Semua', `Hapus <strong>${imageFiles.length} gambar</strong>?`, 'Hapus Semua', 'danger', () => {
        imageFiles.forEach(img => URL.revokeObjectURL(img.previewUrl));
        imageFiles = [];
        updatePreview();
    });
}

function rotateImage(index) { imageFiles[index].rotation = (imageFiles[index].rotation + 90) % 360; updatePreview(); }

// Alternatif reorder selain drag & drop - drag HTML5 tidak terpicu oleh
// sentuhan di layar mobile, jadi tombol ini jadi satu-satunya cara reorder
// yang bisa dipakai di HP (ditampilkan otomatis lewat CSS @media pointer:coarse)
function moveImage(index, direction) {
    const target = index + direction;
    if (target < 0 || target >= imageFiles.length) return;
    [imageFiles[index], imageFiles[target]] = [imageFiles[target], imageFiles[index]];
    updatePreview();
}

function updatePreview() {
    const container = document.getElementById('imagePreviewContainer');
    const count = document.getElementById('imageCount');
    if (imageFiles.length > 0) {
        container.style.display = 'block';
        count.textContent = imageFiles.length;
        document.getElementById('convertBtn').disabled = false;
        document.getElementById('btnText').textContent = `Convert ${imageFiles.length} Gambar`;
    } else {
        container.style.display = 'none';
        document.getElementById('convertBtn').disabled = true;
    }
    renderImageGrid();
}

function renderImageGrid() {
    const grid = document.getElementById('imageGrid');
    if (!grid) return;
    grid.innerHTML = '';

    const pageSize = document.getElementById('pageSize')?.value || 'A4';
    const fitMode = document.getElementById('fitMode')?.value || 'inside';

    imageFiles.forEach((imgData, index) => {
        const item = document.createElement('div');
        item.className = 'image-item';
        item.draggable = true;
        item.dataset.index = index;

        item.addEventListener('dragstart', (e) => { draggedIndex = index; item.classList.add('dragging'); });
        item.addEventListener('dragover', (e) => { e.preventDefault(); item.classList.add('drag-over'); });
        item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
        item.addEventListener('drop', (e) => {
            e.preventDefault();
            item.classList.remove('drag-over');
            const toIndex = parseInt(item.dataset.index);
            if (draggedIndex !== null && draggedIndex !== toIndex) {
                const moved = imageFiles.splice(draggedIndex, 1)[0];
                imageFiles.splice(toIndex, 0, moved);
                updatePreview();
            }
            draggedIndex = null;
        });
        item.addEventListener('dragend', () => { item.classList.remove('dragging'); draggedIndex = null; });

        const rotationStyle = imgData.rotation > 0 ? `transform: rotate(${imgData.rotation}deg);` : '';
        const sizeLabel = { A4: 'A4', Letter: 'Letter', Legal: 'Legal', original: 'Original' }[pageSize] || pageSize;
        const fitLabel = { inside: 'Inside', fill: 'Fill', stretch: 'Stretch' }[fitMode] || fitMode;

        item.innerHTML = `
            <div class="page-number">${index + 1}</div>
            <div class="image-preview-wrapper">
                <div class="page-frame" data-size="${pageSize}" data-fit="${fitMode}">
                    <img src="${imgData.previewUrl}" style="${rotationStyle}">
                </div>
                <div class="page-label">${sizeLabel} · ${fitLabel}</div>
            </div>
            <div class="image-actions">
                <span class="image-name">${imgData.file.name}</span>
                <div class="action-buttons">
                    <button class="action-btn move-btn" onclick="moveImage(${index}, -1)" ${index === 0 ? 'disabled' : ''}>Mundur</button>
                    <button class="action-btn move-btn" onclick="moveImage(${index}, 1)" ${index === imageFiles.length - 1 ? 'disabled' : ''}>Maju</button>
                    <button class="action-btn rotate-btn" onclick="rotateImage(${index})">Putar</button>
                    <button class="action-btn delete-btn" onclick="removeImage(${index})">Hapus</button>
                </div>
            </div>`;
        grid.appendChild(item);
    });
}

// ==========================================
// MERGE PDF
// ==========================================
function switchMergeMode(mode) {
    mergeMode = mode;
    document.querySelectorAll('#mergePdfContainer .merge-mode-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('#mergePdfContainer .merge-mode-tab').forEach(t => {
        if ((mode === 'images' && t.textContent.trim() === 'Gabungkan Gambar') || (mode === 'pdfs' && t.textContent.trim() === 'Gabungkan PDF')) t.classList.add('active');
    });
    document.getElementById('mergeImagesMode').style.display = mode === 'images' ? 'block' : 'none';
    document.getElementById('mergePdfsMode').style.display = mode === 'pdfs' ? 'block' : 'none';
    updateMergeButton();
    scrollToActiveUploadArea('merge-pdf');
}

function addMergeImages(files) {
    const validExts = ['.jpg', '.jpeg', '.png'];
    let added = 0;
    for (let file of files) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (validExts.includes(ext)) {
            mergeImageFiles.push({ file, rotation: 0, id: Date.now() + Math.random(), previewUrl: URL.createObjectURL(file) });
            added++;
        }
    }
    if (added > 0) { updateMergeImagesPreview(); showToast(`${added} gambar ditambahkan`, 'success', 2000); }
}

function addMergePdfs(files) {
    let added = 0;
    for (let file of files) {
        if (file.name.endsWith('.pdf')) {
            mergePdfFiles.push({ file, id: Date.now() + Math.random(), size: file.size });
            added++;
        }
    }
    if (added > 0) { updateMergePdfsPreview(); showToast(`${added} PDF ditambahkan`, 'success', 2000); }
}

function updateMergeImagesPreview() {
    const container = document.getElementById('mergeImagesPreview');
    const count = document.getElementById('mergeImageCount');
    const grid = document.getElementById('mergeImageGrid');
    const pageSize = document.getElementById('mergePageSize')?.value || 'A4';
    const fitMode = document.getElementById('mergeFitMode')?.value || 'inside';

    if (mergeImageFiles.length > 0) { container.style.display = 'block'; count.textContent = mergeImageFiles.length; }
    else container.style.display = 'none';

    grid.innerHTML = '';

    mergeImageFiles.forEach((imgData, index) => {
        const item = document.createElement('div');
        item.className = 'image-item';
        item.draggable = true;
        item.dataset.index = index;

        item.addEventListener('dragstart', (e) => { mergeDraggedIndex = index; item.classList.add('dragging'); });
        item.addEventListener('dragover', (e) => { e.preventDefault(); item.classList.add('drag-over'); });
        item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
        item.addEventListener('drop', (e) => {
            e.preventDefault();
            item.classList.remove('drag-over');
            const toIndex = parseInt(item.dataset.index);
            if (mergeDraggedIndex !== null && mergeDraggedIndex !== toIndex) {
                const moved = mergeImageFiles.splice(mergeDraggedIndex, 1)[0];
                mergeImageFiles.splice(toIndex, 0, moved);
                updateMergeImagesPreview();
            }
            mergeDraggedIndex = null;
        });

        const sizeLabel = { A4: 'A4', Letter: 'Letter', Legal: 'Legal', original: 'Original' }[pageSize] || pageSize;
        const fitLabel = { inside: 'Inside', fill: 'Fill', stretch: 'Stretch' }[fitMode] || fitMode;

        item.innerHTML = `
            <div class="page-number">${index + 1}</div>
            <div class="image-preview-wrapper">
                <div class="page-frame" data-size="${pageSize}" data-fit="${fitMode}">
                    <img src="${imgData.previewUrl}">
                </div>
                <div class="page-label">${sizeLabel} · ${fitLabel}</div>
            </div>
            <div class="image-actions">
                <span class="image-name">${imgData.file.name}</span>
                <div class="action-buttons">
                    <button class="action-btn move-btn" onclick="moveMergeImage(${index}, -1)" ${index === 0 ? 'disabled' : ''}>Mundur</button>
                    <button class="action-btn move-btn" onclick="moveMergeImage(${index}, 1)" ${index === mergeImageFiles.length - 1 ? 'disabled' : ''}>Maju</button>
                    <button class="action-btn delete-btn" onclick="removeMergeImage(${index})">Hapus</button>
                </div>
            </div>`;
        grid.appendChild(item);
    });
    updateMergeButton();
}

function updateMergePdfsPreview() {
    const container = document.getElementById('mergePdfsPreview');
    const count = document.getElementById('mergePdfCount');
    const grid = document.getElementById('mergePdfGrid');

    if (mergePdfFiles.length > 0) { container.style.display = 'block'; count.textContent = mergePdfFiles.length; }
    else container.style.display = 'none';

    grid.innerHTML = '';

    mergePdfFiles.forEach((pdfData, index) => {
        const item = document.createElement('div');
        item.className = 'merge-pdf-item';
        item.draggable = true;
        item.dataset.index = index;

        item.addEventListener('dragstart', (e) => { mergeDraggedIndex = index; item.classList.add('dragging'); });
        item.addEventListener('dragover', (e) => { e.preventDefault(); item.classList.add('drag-over'); });
        item.addEventListener('dragleave', () => item.classList.remove('drag-over'));
        item.addEventListener('drop', (e) => {
            e.preventDefault();
            item.classList.remove('drag-over');
            const toIndex = parseInt(item.dataset.index);
            if (mergeDraggedIndex !== null && mergeDraggedIndex !== toIndex) {
                const moved = mergePdfFiles.splice(mergeDraggedIndex, 1)[0];
                mergePdfFiles.splice(toIndex, 0, moved);
                updateMergePdfsPreview();
            }
            mergeDraggedIndex = null;
        });

        item.innerHTML = `
            <div class="merge-pdf-order">${index + 1}</div>
            <button class="merge-pdf-delete" onclick="removeMergePdf(${index})">Hapus</button>
            <div class="merge-pdf-preview" id="pdfPreview_${pdfData.id}">
                <div class="pdf-preview-loading"><div class="spinner-mini"></div></div>
            </div>
            <div class="merge-pdf-name">${pdfData.file.name}</div>
            <div class="merge-pdf-size">${formatFileSize(pdfData.size)}</div>
            <div class="merge-pdf-actions">
                <button class="action-btn move-btn" onclick="moveMergePdf(${index}, -1)" ${index === 0 ? 'disabled' : ''}>Mundur</button>
                <button class="action-btn move-btn" onclick="moveMergePdf(${index}, 1)" ${index === mergePdfFiles.length - 1 ? 'disabled' : ''}>Maju</button>
            </div>`;
        grid.appendChild(item);

        renderPdfPreview(pdfData.file, pdfData.id);
    });
    updateMergeButton();
}

async function renderPdfPreview(file, id) {
    const container = document.getElementById(`pdfPreview_${id}`);
    if (!container || !window.pdfjsLib) return;

    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const page = await pdf.getPage(1);
        const viewport = page.getViewport({ scale: 0.4 });

        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.className = 'pdf-preview-canvas';

        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        container.innerHTML = '';
        container.appendChild(canvas);
        pdf.destroy();
    } catch (e) {
        container.innerHTML = '<div class="pdf-preview-error"><span>Pratinjau tidak tersedia</span></div>';
    }
}

function removeMergeImage(index) { URL.revokeObjectURL(mergeImageFiles[index].previewUrl); mergeImageFiles.splice(index, 1); updateMergeImagesPreview(); }
function removeMergePdf(index) { mergePdfFiles.splice(index, 1); updateMergePdfsPreview(); }

function moveMergeImage(index, direction) {
    const target = index + direction;
    if (target < 0 || target >= mergeImageFiles.length) return;
    [mergeImageFiles[index], mergeImageFiles[target]] = [mergeImageFiles[target], mergeImageFiles[index]];
    updateMergeImagesPreview();
}

function moveMergePdf(index, direction) {
    const target = index + direction;
    if (target < 0 || target >= mergePdfFiles.length) return;
    [mergePdfFiles[index], mergePdfFiles[target]] = [mergePdfFiles[target], mergePdfFiles[index]];
    updateMergePdfsPreview();
}

function clearMergeImages() { mergeImageFiles.forEach(img => URL.revokeObjectURL(img.previewUrl)); mergeImageFiles = []; updateMergeImagesPreview(); }
function clearMergePdfs() { mergePdfFiles = []; updateMergePdfsPreview(); }

function updateMergeButton() {
    const total = mergeMode === 'images' ? mergeImageFiles.length : mergePdfFiles.length;
    document.getElementById('convertBtn').disabled = total === 0;
    document.getElementById('btnText').textContent = total > 0 ? (mergeMode === 'images' ? `Gabungkan ${total} Gambar → PDF` : `Gabungkan ${total} PDF`) : 'Convert Sekarang';
}

// ==========================================
// CONVERT BUTTON
// ==========================================
document.getElementById('convertBtn').addEventListener('click', async () => {
    const btnText = document.getElementById('btnText');
    const spinner = document.getElementById('spinner');

    if (currentMode === 'image-to-pdf' && !imageFiles.length) return showToast('Tidak ada gambar!', 'error');
    if (currentMode === 'pdf-to-image' && !pdfToImageFile) return showToast('Tidak ada PDF!', 'error');
    if (SINGLE_FILE_MODES.includes(currentMode) && !selectedFiles.length) return showToast('Tidak ada file!', 'error');
    if (currentMode === 'compress-pdf' && !compressFile) return showToast('Tidak ada PDF!', 'error');
    if (currentMode === 'compress-image' && !compressImageFile) return showToast('Tidak ada gambar!', 'error');
    if (currentMode === 'compress-image-target' && !compressImageTargetFile) return showToast('Tidak ada gambar!', 'error');
    if (currentMode === 'remove-watermark' && !watermarkFile) return showToast('Tidak ada file!', 'error');
    if (currentMode === 'scan-document' && !scanFile) return showToast('Tidak ada gambar!', 'error');
    if (currentMode === 'split-pdf') {
        if (!splitFile) return showToast('Tidak ada PDF!', 'error');
        if (splitMode === 'range' && !document.getElementById('splitRangeInput').value.trim()) {
            return showToast('Range halaman tidak boleh kosong! (Contoh: 1-3, 5)', 'error');
        }
        if (splitMode === 'every') {
            const val = parseInt(document.getElementById('splitEveryInput').value);
            if (!val || val < 1) {
                return showToast('Jumlah halaman per file minimal 1!', 'error');
            }
        }
    }
    if (currentMode === 'merge-pdf' && ((mergeMode === 'images' && !mergeImageFiles.length) || (mergeMode === 'pdfs' && !mergePdfFiles.length))) return showToast('Tidak ada file!', 'error');

    btnText.style.display = 'none';
    spinner.style.display = 'block';
    document.getElementById('convertBtn').disabled = true;
    document.getElementById('status').style.display = 'none';

    const formData = new FormData();

    if (currentMode === 'pdf-to-word') {
        formData.append('file', selectedFiles[0]);
        formData.append('method', document.getElementById('convertMethod')?.value || 'auto');
    } else if (SINGLE_FILE_MODES.includes(currentMode)) {
        formData.append('file', selectedFiles[0]);
    } else if (currentMode === 'image-to-pdf') {
        imageFiles.forEach(img => formData.append('files', img.file));
        formData.append('rotations', JSON.stringify(imageFiles.map(img => img.rotation)));
        formData.append('page_size', document.getElementById('pageSize').value);
        formData.append('fit', document.getElementById('fitMode').value);
    } else if (currentMode === 'compress-pdf') {
        formData.append('file', compressFile);
        formData.append('compression_level', document.querySelector('input[name="compression"]:checked')?.value || 'medium');
    } else if (currentMode === 'compress-image') {
        formData.append('file', compressImageFile);
        formData.append('compression_level', document.querySelector('input[name="image_compression"]:checked')?.value || 'medium');
    } else if (currentMode === 'compress-image-target') {
        formData.append('file', compressImageTargetFile);
        const targetCard = document.querySelector('input[name="target_size"]:checked');
        let targetKb = 1000;
        if (targetCard) {
            if (targetCard.value === 'custom') {
                targetKb = parseInt(document.getElementById('customTargetKb')?.value) || 1000;
            } else {
                targetKb = parseInt(targetCard.value) || 1000;
            }
        }
        formData.append('target_kb', targetKb);
    } else if (currentMode === 'pdf-to-image') {
        formData.append('file', pdfToImageFile);
        formData.append('format', document.getElementById('pdfToImageFormat')?.value || 'jpg');
        formData.append('dpi', document.getElementById('pdfToImageDpi')?.value || '150');
        const pageMode = document.getElementById('pdfToImagePages')?.value || 'all';
        if (pageMode === 'custom') {
            const customPages = document.getElementById('pdfToImageRangeInput')?.value.trim();
            formData.append('pages', customPages || 'all');
        } else {
            formData.append('pages', pageMode);
        }
    } else if (currentMode === 'merge-pdf') {
        if (mergeMode === 'images') {
            mergeImageFiles.forEach(img => formData.append('files', img.file));
            formData.append('rotations', JSON.stringify(mergeImageFiles.map(img => img.rotation)));
            formData.append('page_size', document.getElementById('mergePageSize').value);
            formData.append('fit', document.getElementById('mergeFitMode').value);
        } else {
            mergePdfFiles.forEach(pdf => formData.append('files', pdf.file));
        }
    } else if (currentMode === 'split-pdf') {
        formData.append('file', splitFile);
        formData.append('mode', splitMode);
        if (splitMode === 'range') {
            formData.append('ranges', document.getElementById('splitRangeInput').value);
        } else {
            formData.append('every_n', document.getElementById('splitEveryInput').value);
        }
    } else if (currentMode === 'remove-watermark') {
        formData.append('file', watermarkFile);
        formData.append('mode', document.getElementById('watermarkMode').value);
    } else if (currentMode === 'scan-document') {
        const img = document.getElementById('scanPreviewImage');
        const renderW = img.offsetWidth;
        const renderH = img.offsetHeight;
        const scaleX = scanImageWidth / renderW;
        const scaleY = scanImageHeight / renderH;

        // Skala 4 titik sudut ke resolusi asli gambar untuk perspective transform
        const originalCorners = scanCorners.map(p => [
            Math.round(p.x * scaleX),
            Math.round(p.y * scaleY)
        ]);

        formData.append('file', scanFile);
        formData.append('corners', JSON.stringify(originalCorners));

        // Bounding box fallback
        const xs = originalCorners.map(p => p[0]);
        const ys = originalCorners.map(p => p[1]);
        const minX = Math.max(0, Math.min(...xs));
        const minY = Math.max(0, Math.min(...ys));
        const maxX = Math.min(scanImageWidth, Math.max(...xs));
        const maxY = Math.min(scanImageHeight, Math.max(...ys));

        formData.append('crop_x', minX);
        formData.append('crop_y', minY);
        formData.append('crop_w', maxX - minX);
        formData.append('crop_h', maxY - minY);
    }

    const endpoints = {
        'pdf-to-word': '/convert/pdf-to-word',
        'word-to-pdf': '/convert/word-to-pdf',
        'pdf-to-markdown': '/convert/pdf-to-markdown',
        'word-to-markdown': '/convert/word-to-markdown',
        'image-to-pdf': '/convert/image-to-pdf',
        'pdf-to-image': '/convert/pdf-to-image',
        'ppt-to-pdf': '/convert/ppt-to-pdf',
        'pdf-to-ppt': '/convert/pdf-to-ppt',
        'excel-to-pdf': '/convert/excel-to-pdf',
        'pdf-to-excel': '/convert/pdf-to-excel',
        'compress-pdf': '/compress/pdf',
        'compress-image': '/compress/image',
        'compress-image-target': '/compress/image-to-target',
        'merge-pdf': mergeMode === 'images' ? '/merge/images' : '/merge/pdf',
        'split-pdf': '/split/pdf',
        'remove-watermark': '/remove-watermark',
        'scan-document': '/scan-document'
    };

    try {
        const response = await fetch(endpoints[currentMode], { method: 'POST', body: formData });
        if (!response.ok) { const err = await response.json(); throw new Error(err.detail || 'Gagal'); }

        const blob = await response.blob();

        if (currentMode === 'compress-pdf') {
            const os = parseInt(response.headers.get('X-Original-Size') || compressFile.size);
            const cs = parseInt(response.headers.get('X-Compressed-Size') || blob.size);
            const ratio = parseFloat(response.headers.get('X-Compression-Ratio') || ((1 - cs/os) * 100).toFixed(1));
            document.getElementById('originalSize').textContent = formatFileSize(os);
            document.getElementById('compressedSize').textContent = formatFileSize(cs);
            document.getElementById('savedSize').textContent = `${ratio}%`;
            document.getElementById('compressOptions').style.display = 'none';
            document.getElementById('compressResult').style.display = 'block';
            window.compressedBlob = blob;
        } else if (currentMode === 'compress-image') {
            const os = parseInt(response.headers.get('X-Original-Size') || compressImageFile.size);
            const cs = parseInt(response.headers.get('X-Compressed-Size') || blob.size);
            const ratio = parseFloat(response.headers.get('X-Compression-Ratio') || ((1 - cs/os) * 100).toFixed(1));
            document.getElementById('imageOriginalSize').textContent = formatFileSize(os);
            document.getElementById('imageCompressedSize').textContent = formatFileSize(cs);
            document.getElementById('imageSavedSize').textContent = `${ratio}%`;
            document.getElementById('compressImageOptions').style.display = 'none';
            document.getElementById('compressImageResult').style.display = 'block';
            window.compressedImageBlob = blob;
            showStatus('Kompresi gambar berhasil!', 'success');
        } else if (currentMode === 'compress-image-target') {
            const os = parseInt(response.headers.get('X-Original-Size') || compressImageTargetFile.size);
            const cs = parseInt(response.headers.get('X-Compressed-Size') || blob.size);
            const ratio = parseFloat(response.headers.get('X-Compression-Ratio') || ((1 - cs/os) * 100).toFixed(1));
            document.getElementById('targetOriginalSize').textContent = formatFileSize(os);
            document.getElementById('targetCompressedSize').textContent = formatFileSize(cs);
            document.getElementById('targetSavedSize').textContent = `${ratio}%`;
            document.getElementById('compressImageTargetOptions').style.display = 'none';
            document.getElementById('compressImageTargetResult').style.display = 'block';
            window.compressedTargetBlob = blob;
            showStatus('Kompresi Standar Administrasi berhasil!', 'success');
        } else if (currentMode === 'pdf-to-markdown' || currentMode === 'word-to-markdown') {
            // Simpan blob untuk download ulang
            window.currentMarkdownBlob = blob;

            // Baca teks markdown untuk statistik & preview
            const markdownText = await blob.text();
            window.currentMarkdownText = markdownText;

            // Hitung statistik
            const charCount = markdownText.length;
            const wordCount = markdownText.trim().split(/\s+/).filter(Boolean).length;
            const tokenEstimate = Math.ceil(charCount / 4); // ~4 karakter per token

            document.getElementById('mdCharCount').textContent = charCount.toLocaleString('id-ID');
            document.getElementById('mdWordCount').textContent = wordCount.toLocaleString('id-ID');
            document.getElementById('mdTokenEstimate').textContent = `~${tokenEstimate.toLocaleString('id-ID')}`;

            // Set preview (batasi 3000 char agar tidak lag)
            const previewText = markdownText.length > 3000
                ? markdownText.substring(0, 3000) + '\n\n... (terpotong, download file untuk konten lengkap)'
                : markdownText;
            document.getElementById('markdownPreviewCode').textContent = previewText;

            // Auto-download file .md
            const contentDisp = response.headers.get('content-disposition') || '';
            const match = contentDisp.match(/filename\*?=(?:UTF-8'')?["']?([^"';\n]+)["']?/i);
            const filename = match ? match[1].trim() : `${selectedFiles[0]?.name?.replace(/\.[^.]+$/, '') || 'dokumen'}.md`;

            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);

            // Tampilkan result box, sembunyikan convert btn & upload
            document.getElementById('singleFileUpload').style.display = 'none';
            document.getElementById('convertBtn').style.display = 'none';
            document.getElementById('status').style.display = 'none';
            document.getElementById('markdownResult').style.display = 'block';
            document.getElementById('markdownPreviewArea').style.display = 'none';
        } else {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const contentDisp = response.headers.get('content-disposition') || '';
            const match = contentDisp.match(/filename\*?=(?:UTF-8'')?["']?([^"';\n]+)["']?/i);
            a.download = match ? match[1].trim() : 'file';
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
            showStatus('Berhasil! File terdownload otomatis.', 'success');
        }
    } catch (err) {
        showStatus(`Error: ${err.message}`, 'error');
    } finally {
        btnText.style.display = 'inline';
        spinner.style.display = 'none';
        if (currentMode !== 'pdf-to-markdown' && currentMode !== 'word-to-markdown') {
            document.getElementById('convertBtn').disabled = false;
        }
    }
});

document.getElementById('downloadCompressedBtn')?.addEventListener('click', () => {
    if (window.compressedBlob) {
        const url = URL.createObjectURL(window.compressedBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `compressed_${compressFile?.name || 'file.pdf'}`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }
});

document.getElementById('downloadCompressedImageBtn')?.addEventListener('click', () => {
    if (window.compressedImageBlob) {
        const url = URL.createObjectURL(window.compressedImageBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `compressed_${compressImageFile?.name || 'image.jpg'}`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }
});

document.getElementById('downloadCompressedTargetBtn')?.addEventListener('click', () => {
    if (window.compressedTargetBlob) {
        const url = URL.createObjectURL(window.compressedTargetBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `admin_${compressImageTargetFile?.name?.replace(/\.[^.]+$/, '') || 'dokumen'}.jpg`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }
});

// ==========================================
// MARKDOWN RESULT ACTIONS
// ==========================================
async function copyMarkdownContent() {
    if (!window.currentMarkdownText) return;
    try {
        await navigator.clipboard.writeText(window.currentMarkdownText);
        const btn = document.getElementById('copyBtnText');
        btn.textContent = '✅ Tersalin!';
        setTimeout(() => { btn.textContent = 'Salin Teks Markdown'; }, 2500);
    } catch (e) {
        showToast('Gagal menyalin ke clipboard', 'error');
    }
}

function downloadCurrentMarkdown() {
    if (!window.currentMarkdownBlob) return;
    const filename = `${selectedFiles[0]?.name?.replace(/\.[^.]+$/, '') || 'dokumen'}.md`;
    const url = URL.createObjectURL(window.currentMarkdownBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast('File .md terdownload!', 'success');
}

function toggleMarkdownPreview() {
    const area = document.getElementById('markdownPreviewArea');
    const btn = document.getElementById('togglePreviewBtn');
    const isShowing = area.style.display !== 'none';
    area.style.display = isShowing ? 'none' : 'block';
    btn.textContent = isShowing ? '👁️ Lihat Pratinjau' : '🙈 Sembunyikan';
}

function resetMarkdown() {
    window.currentMarkdownText = '';
    window.currentMarkdownBlob = null;
    switchTab(currentMode);
}


function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024, sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// ==========================================
// INIT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });

    const modal = document.getElementById('confirmModal');
    modal?.addEventListener('click', (e) => { if (e.target.id === 'confirmModal') closeModal(); });

    document.getElementById('pageSize')?.addEventListener('change', renderImageGrid);
    document.getElementById('fitMode')?.addEventListener('change', renderImageGrid);
    document.getElementById('mergePageSize')?.addEventListener('change', updateMergeImagesPreview);
    document.getElementById('mergeFitMode')?.addEventListener('change', updateMergeImagesPreview);
});