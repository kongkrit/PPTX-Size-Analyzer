// --- Configuration ---
let config = {
    numberOfFiles: 40,
    minNumberOfFiles: 1,
    maxNumberOfFiles: 500,
    showImage: true,
    showSlides: false,
    showOther: false
};

// --- Utility ---
const byId = id  => document.getElementById(id);
const qs   = sel => document.querySelector(sel);
const qsa  = sel => Array.from(document.querySelectorAll(sel));

// debounce wrapper
function debounce(fn, delay = 100) {
    let t;
    return (...args) => {
        clearTimeout(t);
        t = setTimeout(() => fn.apply(null, args), delay);
    };
}

// --- DOM elements ---
const dom = {
    dropZone: byId("drop-zone"),
    fileInput: byId("file-input"),
    results: byId("results"),
    totalSize: byId("file-info"),
    catTable: qs("#category-table tbody"),
    fileTable: qs("#largest-files-table tbody"),
    errorMsg: byId("error-msg"),
    fileLimit: byId("fileLimit"),
    showImage: byId("show-images"),
    showSlides: byId("show-slides"),
    showOther: byId("show-others"),
    imageViewer: byId("image-viewer")
};

// --- State ---
let state = {
    zip: null,
    filename: "",
    totalUncompressed: 0,
    categories: {},
    files: [] // Sorted list of all files
};

// --- Init UI ---
dom.fileLimit.value = config.numberOfFiles;
dom.fileLimit.min = config.minNumberOfFiles;
dom.fileLimit.max = config.maxNumberOfFiles;
dom.showImage.checked = config.showImage;
dom.showSlides.checked = config.showSlides;
dom.showOther.checked = config.showOther;

// --- Logic ---

function formatBytes(bytes, decimals = 2) {
    if (!+bytes) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}

function getCategory(filename) {
    if (filename.match(/ppt\/media\/.+/)) {
        if (filename.match(/\.(png|jpg|jpeg|gif|tiff|bmp)$/i)) return 'Images';
        if (filename.match(/\.(mp4|avi|mov|wmv|m4v)$/i)) return 'Video';
        if (filename.match(/\.(mp3|wav|m4a)$/i)) return 'Audio';
        return 'Other Media';
    }
    if (filename.match(/ppt\/fonts\/.+/)) return 'Embedded Fonts';
    if (filename.match(/ppt\/slides\/.+/)) return 'Slides (XML)';
    if (filename.match(/ppt\/embeddings\/.+/)) return 'Excel/Object Embeds';
    return 'Structure/XML';
}

async function analyzePPTX(file) {
    dom.errorMsg.textContent = "";
    dom.results.classList.add('hidden');
    
    try {
        const zip = await JSZip.loadAsync(file);
        
        // Reset state
        state.zip = zip;
        state.files = [];
        state.categories = {
            'Images': 0, 'Video': 0, 'Audio': 0, 
            'Embedded Fonts': 0, 'Slides (XML)': 0, 
            'Excel/Object Embeds': 0, 'Structure/XML': 0, 
            'Other Media': 0
        };
        state.totalUncompressed = 0;
        state.filename = file.name;

        zip.forEach((relativePath, zipEntry) => {
            if (!zipEntry.dir) {
                const size = zipEntry._data.uncompressedSize; 
                state.totalUncompressed += size;
                
                const cat = getCategory(relativePath);
                state.categories[cat] += size;

                state.files.push({
                    name: relativePath,
                    size: size,
                    cat: cat
                });
            }
        });

        // Sort once (Listing order preserved as per requirement)
        state.files.sort((a, b) => b.size - a.size);

        renderFullView();

    } catch (e) {
        console.error(e);
        dom.errorMsg.textContent = "Error parsing file. Ensure it is a valid .pptx";
    }
}

async function showImagePreview(filename) {
    if (!state.zip) return;
    try {
        const file = state.zip.file(filename);
        if (file) {
            const blob = await file.async("blob");
            const url = URL.createObjectURL(blob);
            
            dom.imageViewer.innerHTML = "";
            const img = document.createElement("img");
            img.src = url;
            img.title = filename;
            dom.imageViewer.appendChild(img);
            
            // Add a close button or similar if needed, or just replace content
            const caption = document.createElement("div");
            caption.textContent = filename;
            caption.style.marginTop = "0.5rem";
            caption.style.fontSize = "0.9em";
            dom.imageViewer.appendChild(caption);
            
            dom.imageViewer.classList.remove("hidden");
        }
    } catch (e) {
        console.error("Error showing image:", e);
    }
}

function renderFullView() {
    dom.totalSize.textContent = `${state.filename} (${formatBytes(state.totalUncompressed)})`;
    dom.catTable.innerHTML = "";
    dom.imageViewer.innerHTML = "";
    dom.imageViewer.classList.add("hidden");
    
    // Render Categories
    const sortedCats = Object.entries(state.categories).sort((a, b) => b[1] - a[1]);
    
    sortedCats.forEach(([name, size]) => {
        if (size === 0) return;
        const pct = ((size / state.totalUncompressed) * 100).toFixed(1);
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <div>${name}</div>
                <div class="bar-container"><div class="bar-fill" style="width:${pct}%"></div></div>
            </td>
            <td>${formatBytes(size)}</td>
            <td>${pct}%</td>
        `;
        dom.catTable.appendChild(row);
    });

    renderFilesList();
    dom.results.classList.remove('hidden');
}

function renderFilesList() {
    dom.fileTable.innerHTML = "";
    
    // Filter based on checkboxes
    const filteredFiles = state.files.filter(f => {
        if (f.cat === 'Images' && !config.showImage) return false;
        if (f.cat === 'Slides (XML)' && !config.showSlides) return false;
        if (f.cat !== 'Images' && f.cat !== 'Slides (XML)' && !config.showOther) return false;
        return true;
    });

    // Slice based on config SoT
    const limit = Math.max(config.minNumberOfFiles, Math.min(config.numberOfFiles, config.maxNumberOfFiles));
    const listToRender = filteredFiles.slice(0, limit);

    listToRender.forEach(f => {
        const row = document.createElement('tr');
        
        let nameHtml = f.name;
        if (f.cat === 'Images') {
            nameHtml = `<span class="clickable-file" data-file="${f.name}">${f.name}</span>`;
        }

        row.innerHTML = `
            <td class="file-name">${nameHtml}</td>
            <td>${formatBytes(f.size)}</td>
            <td>${f.cat}</td>
        `;
        dom.fileTable.appendChild(row);
    });

    // Add click listeners
    const clickables = dom.fileTable.querySelectorAll('.clickable-file');
    clickables.forEach(el => {
        el.addEventListener('click', (e) => {
            showImagePreview(e.target.dataset.file);
        });
    });
}

// --- Event Listeners ---

// File Inputs
dom.dropZone.addEventListener('click', () => dom.fileInput.click());
dom.fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) analyzePPTX(e.target.files[0]);
});
dom.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dom.dropZone.style.borderColor = '#888';
});
dom.dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dom.dropZone.style.borderColor = '#444';
});
dom.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dom.dropZone.style.borderColor = '#444';
    if (e.dataTransfer.files[0]) {
        analyzePPTX(e.dataTransfer.files[0]);
    }
});

// Config Input Handler
dom.fileLimit.addEventListener('input', debounce((e) => {
    const val = parseInt(e.target.value, 10);
    if (!isNaN(val)) {
        config.numberOfFiles = val;
        // Re-render only the file list part if we have data
        if (state.files.length > 0) {
            renderFilesList();
        }
    }
}, 300));

// Checkbox Handlers
const refreshList = () => {
    if (state.files.length > 0) renderFilesList();
};
dom.showImage.addEventListener('change', (e) => { config.showImage = e.target.checked; refreshList(); });
dom.showSlides.addEventListener('change', (e) => { config.showSlides = e.target.checked; refreshList(); });
dom.showOther.addEventListener('change', (e) => { config.showOther = e.target.checked; refreshList(); });
