document.addEventListener('DOMContentLoaded', function() {
    const dropArea = document.getElementById('dropArea');
    const fileInput = document.getElementById('fileInput');
    const statusDiv = document.getElementById('status');
    const viewerContainer = document.getElementById('viewerContainer');
    const officeFrame = document.getElementById('officeFrame');
    const fullscreenBtn = document.getElementById('fullscreenBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    
    let currentFile = null;

    // Drag and drop events
    dropArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropArea.classList.add('highlight');
    });

    dropArea.addEventListener('dragleave', () => {
        dropArea.classList.remove('highlight');
    });

    dropArea.addEventListener('drop', (e) => {
        e.preventDefault();
        dropArea.classList.remove('highlight');
        
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // File input change
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) {
            handleFile(fileInput.files[0]);
        }
    });

    // Button events
    fullscreenBtn.addEventListener('click', toggleFullscreen);
    downloadBtn.addEventListener('click', downloadFile);

    function handleFile(file) {
        // Validate file type
        if (!file.name.match(/\.(ppt|pptx)$/i)) {
            showStatus('Please upload a PowerPoint file (.ppt or .pptx)', 'error');
            return;
        }

        currentFile = file;
        showStatus('File loaded: ' + file.name, 'success');
        
        // Display in Office Online Viewer
        const fileUrl = URL.createObjectURL(file);
        officeFrame.src = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(fileUrl)}`;
        
        viewerContainer.style.display = 'block';
    }

    function showStatus(message, type) {
        statusDiv.textContent = message;
        statusDiv.className = 'status ' + type;
        statusDiv.style.display = 'block';
    }

    function toggleFullscreen() {
        if (!document.fullscreenElement) {
            officeFrame.requestFullscreen().catch(err => {
                showStatus('Fullscreen failed: ' + err.message, 'error');
            });
        } else {
            document.exitFullscreen();
        }
    }

    function downloadFile() {
        if (!currentFile) return;
        
        const a = document.createElement('a');
        a.href = URL.createObjectURL(currentFile);
        a.download = currentFile.name;
        a.click();
    }
});