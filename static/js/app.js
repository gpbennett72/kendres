let currentSessionId = null;
let uploadedFiles = {};

// Tab switching is now defined below with admin auto-load

// Show/hide loading overlay
function showLoading(message = 'Processing...') {
    document.getElementById('loading-message').textContent = message;
    document.getElementById('loading-overlay').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loading-overlay').style.display = 'none';
}

// Show toast notification
function showToast(message, type = 'error') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = `toast ${type}`;
    toast.style.display = 'block';
    
    setTimeout(() => {
        toast.style.display = 'none';
    }, 5000);
}

// Show error message (backward compatibility)
function showError(message) {
    showToast(message, 'error');
}

// Reset interface function
function resetInterface() {
    // Clear form
    document.getElementById('upload-form').reset();
    document.getElementById('google-form').reset();
    
    // Hide sections
    document.getElementById('config-section').style.display = 'none';
    document.getElementById('results-section').style.display = 'none';
    document.getElementById('download-section').style.display = 'none';
    
    // Clear session
    currentSessionId = null;
    uploadedFiles = {};
    
    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
    
    // Show success message
    showToast('Interface reset. Ready for new document upload.', 'success');
}

// Update file label when file is selected
document.getElementById('document').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        const label = document.querySelector('label[for="document"] .file-label-text');
        if (label) {
            label.textContent = file.name;
        }
    }
});

// Upload form handler
document.getElementById('upload-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData();
    const documentFile = document.getElementById('document').files[0];
    
    if (!documentFile) {
        showError('Please select a document file');
        return;
    }
    
    formData.append('document', documentFile);
    
    showLoading('Uploading files...');
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok) {
            currentSessionId = data.session_id;
            uploadedFiles = {
                document: data.document,
                playbook: data.playbook
            };
            
            showToast('Document uploaded successfully!', 'success');
            
            // Show configuration section
            document.getElementById('config-section').style.display = 'block';
            document.getElementById('config-section').scrollIntoView({ behavior: 'smooth' });
        } else {
            showError(data.error || 'Upload failed');
        }
    } catch (error) {
        showError('Error uploading files: ' + error.message);
    } finally {
        hideLoading();
    }
});

// Process document
async function processDocument() {
    const aiProvider = document.getElementById('ai_provider').value;
    const model = document.getElementById('model').value;
    const createSummary = document.getElementById('create_summary').checked;
    const useTrackedChanges = document.getElementById('use_tracked_changes').checked;
    const contractTypeId = document.getElementById('contract_type')?.value || '';
    
    const outputType = useTrackedChanges ? 'tracked changes (counter redlines)' : 'comments';
    showLoading(`Analyzing redlines with AI and inserting ${outputType}... This may take a few minutes.`);
    
    try {
        const response = await fetch('/api/process', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                ai_provider: aiProvider,
                model: model,
                create_summary: createSummary,
                use_tracked_changes: useTrackedChanges,
                contract_type_id: contractTypeId || null
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showToast('Document processed successfully!', 'success');
            displayResults(data);
        } else {
            showError(data.error || 'Processing failed');
        }
    } catch (error) {
        showError('Error processing document: ' + error.message);
    } finally {
        hideLoading();
    }
}

// Analyze only
async function analyzeOnly() {
    const aiProvider = document.getElementById('ai_provider').value;
    const model = document.getElementById('model').value;
    const contractTypeId = document.getElementById('contract_type')?.value || '';
    
    showLoading('Analyzing redlines...');
    
    try {
        const response = await fetch('/api/analyze-only', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                ai_provider: aiProvider,
                model: model,
                contract_type_id: contractTypeId || null
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showToast('Analysis completed!', 'success');
            displayResults(data, false);
        } else {
            showError(data.error || 'Analysis failed');
        }
    } catch (error) {
        showError('Error analyzing document: ' + error.message);
    } finally {
        hideLoading();
    }
}

// Display results
function displayResults(data, showDownload = true) {
    const resultsSection = document.getElementById('results-section');
    const summaryDiv = document.getElementById('results-summary');
    const detailsDiv = document.getElementById('results-details');
    const downloadSection = document.getElementById('download-section');
    const downloadLinks = document.getElementById('download-links');
    
    const meta = data.metadata || {};
    // Summary with metadata
    summaryDiv.innerHTML = `
        <h3>Contract Info</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 0.75rem; margin-top: 0.75rem;">
            <div><strong>Party 1:</strong> ${meta.party_one || 'Not detected'}</div>
            <div><strong>Party 2:</strong> ${meta.party_two || 'Not detected'}</div>
            <div><strong>Contract Type:</strong> ${meta.contract_type || 'Default'}</div>
            <div><strong>Document:</strong> ${meta.document || 'Unknown'}</div>
            <div><strong>Playbook:</strong> ${meta.playbook || 'default_playbook.txt'}</div>
            <div><strong>AI:</strong> ${meta.ai_provider || 'n/a'} (${meta.model || ''})</div>
        </div>
        <h3 style="margin-top: 1rem;">Summary</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-top: 0.5rem;">
            <div>
                <div style="font-size: 2rem; font-weight: 700; color: var(--primary);">${data.redlines_count}</div>
                <div style="color: var(--text-secondary); font-size: 0.875rem;">Redlines Found</div>
            </div>
            <div>
                <div style="font-size: 2rem; font-weight: 700; color: var(--primary);">${data.analyses.length}</div>
                <div style="color: var(--text-secondary); font-size: 0.875rem;">Analyses Completed</div>
            </div>
        </div>
    `;
    
    // Details
    detailsDiv.innerHTML = '';
    
    if (data.analyses.length === 0) {
        detailsDiv.innerHTML = '<div class="analysis-item"><p style="text-align: center; color: var(--text-secondary);">No redlines found in the document.</p></div>';
    } else {
        data.analyses.forEach((analysis, index) => {
            const riskClass = `risk-${analysis.risk_level.toLowerCase()}`;
            const analysisDiv = document.createElement('div');
            analysisDiv.className = 'analysis-item';
            analysisDiv.innerHTML = `
                <h4>
                    Redline #${index + 1}
                    <span class="risk-badge ${riskClass}">${analysis.risk_level} Risk</span>
                </h4>
                <div class="analysis-detail">
                    <p><strong>Type:</strong> <span style="text-transform: capitalize;">${analysis.type}</span></p>
                    <p><strong>Text:</strong> ${analysis.text || 'N/A'}</p>
                    ${analysis.assessment ? `<p><strong>Assessment:</strong> ${analysis.assessment.replace(/\n/g, '<br>')}</p>` : ''}
                    ${analysis.response ? `<p><strong>Recommended Action:</strong> ${analysis.response}</p>` : ''}
                </div>
            `;
            detailsDiv.appendChild(analysisDiv);
        });
    }
    
    // Download links
    if (showDownload && data.output_file) {
        downloadLinks.innerHTML = '';
        
        if (data.output_file) {
            const outputLink = document.createElement('a');
            outputLink.href = `/api/download/${currentSessionId}/${data.output_file}`;
            outputLink.className = 'download-link';
            outputLink.textContent = `📄 Download Document with Comments`;
            outputLink.download = data.output_file;
            downloadLinks.appendChild(outputLink);
        }
        
        if (data.summary_file) {
            const summaryLink = document.createElement('a');
            summaryLink.href = `/api/download/${currentSessionId}/${data.summary_file}`;
            summaryLink.className = 'download-link';
            summaryLink.textContent = `📊 Download Summary Document`;
            summaryLink.download = data.summary_file;
            downloadLinks.appendChild(summaryLink);
        }
        
        downloadSection.style.display = 'block';
    } else {
        downloadSection.style.display = 'none';
    }
    
    // Show results section
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

// Admin functions
async function loadPlaybook() {
    showLoading('Loading playbook...');
    try {
        const response = await fetch('/api/admin/playbook');
        const data = await response.json();
        
        if (response.ok) {
            document.getElementById('playbook_content').value = data.content;
            showAdminMessage('Playbook loaded successfully', 'success');
        } else {
            showAdminMessage('Error loading playbook: ' + (data.error || 'Unknown error'), 'error');
        }
    } catch (error) {
        showAdminMessage('Error loading playbook: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

async function updatePlaybook() {
    const content = document.getElementById('playbook_content').value;
    
    if (!content.trim()) {
        showAdminMessage('Playbook content cannot be empty', 'error');
        return;
    }
    
    showLoading('Updating playbook...');
    try {
        const response = await fetch('/api/admin/playbook', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ content: content })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showAdminMessage('Playbook updated successfully', 'success');
        } else {
            showAdminMessage('Error updating playbook: ' + (data.error || 'Unknown error'), 'error');
        }
    } catch (error) {
        showAdminMessage('Error updating playbook: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

async function downloadPlaybook() {
    try {
        const response = await fetch('/api/admin/playbook/download');
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'default_playbook.txt';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            showAdminMessage('Playbook downloaded', 'success');
        } else {
            const data = await response.json();
            showAdminMessage('Error downloading playbook: ' + (data.error || 'Unknown error'), 'error');
        }
    } catch (error) {
        showAdminMessage('Error downloading playbook: ' + error.message, 'error');
    }
}

async function uploadPlaybookFile() {
    const fileInput = document.getElementById('playbook_upload');
    const file = fileInput.files[0];
    
    if (!file) {
        showAdminMessage('Please select a file to upload', 'error');
        return;
    }
    
    showLoading('Uploading playbook...');
    try {
        const formData = new FormData();
        formData.append('playbook', file);
        
        const response = await fetch('/api/admin/playbook/upload', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok) {
            // Reload the playbook content
            await loadPlaybook();
            fileInput.value = ''; // Clear file input
            showAdminMessage('Playbook uploaded and updated successfully', 'success');
        } else {
            showAdminMessage('Error uploading playbook: ' + (data.error || 'Unknown error'), 'error');
        }
    } catch (error) {
        showAdminMessage('Error uploading playbook: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

function showAdminMessage(message, type) {
    const messageDiv = document.getElementById('admin-message');
    messageDiv.textContent = message;
    messageDiv.className = `alert ${type}`;
    messageDiv.style.display = 'block';
    
    setTimeout(() => {
        messageDiv.style.display = 'none';
    }, 5000);
}

// Load playbook when admin tab is opened
let adminTabLoaded = false;
let adminAuthenticated = false;
const ADMIN_PASSWORD = 'redline2025';  // Change this password as needed

function switchTab(tab) {
    // Password protect admin tab
    if (tab === 'admin' && !adminAuthenticated) {
        const password = prompt('Enter admin password:');
        if (password !== ADMIN_PASSWORD) {
            if (password !== null) {  // User didn't cancel
                showToast('Incorrect password', 'error');
            }
            return;  // Don't switch to admin tab
        }
        adminAuthenticated = true;
    }
    
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
        content.style.display = 'none';
    });
    
    // Remove active class from all buttons
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show selected tab
    const tabContent = document.getElementById(`${tab}-tab`);
    if (tabContent) {
        tabContent.classList.add('active');
        tabContent.style.display = 'block';
    }
    
    // Activate button
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
        if (btn.dataset.tab === tab) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
    
    // Auto-load playbook when admin tab is opened
    if (tab === 'admin' && !adminTabLoaded) {
        loadPlaybook();
        adminTabLoaded = true;
    }
}

// Google Doc form handler
document.getElementById('google-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const docId = document.getElementById('doc_id').value;
    const aiProvider = document.getElementById('google_ai_provider').value;
    const model = document.getElementById('google_model').value;
    const contractTypeId = document.getElementById('google_contract_type')?.value || '';
    
    // Extract doc ID from URL if full URL provided
    let extractedDocId = docId;
    if (docId.includes('docs.google.com')) {
        const match = docId.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (match) {
            extractedDocId = match[1];
        }
    }
    
    showLoading('Processing Google Doc... This may take a few minutes.');
    
    try {
        const response = await fetch('/api/process-google', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                doc_id: extractedDocId,
                ai_provider: aiProvider,
                model: model,
                contract_type_id: contractTypeId || null
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            displayResults(data, false);
            alert('Comments have been added to your Google Doc!');
        } else {
            showError(data.error || 'Processing failed');
        }
    } catch (error) {
        showError('Error processing Google Doc: ' + error.message);
    } finally {
        hideLoading();
    }
});

// Initialize model dropdown on page load
document.addEventListener('DOMContentLoaded', function() {
    // Set Anthropic as default and update model dropdown
    const aiProvider = document.getElementById('ai_provider');
    const googleAiProvider = document.getElementById('google_ai_provider');
    
    if (aiProvider.value === 'anthropic') {
        const modelSelect = document.getElementById('model');
        modelSelect.innerHTML = '<option value="claude-3-haiku-20240307" selected>Claude 3 Haiku</option>';
    }
    
    if (googleAiProvider.value === 'anthropic') {
        const googleModelSelect = document.getElementById('google_model');
        googleModelSelect.innerHTML = '<option value="claude-3-haiku-20240307" selected>Claude 3 Haiku</option>';
    }
});

// Update model options based on provider
document.getElementById('ai_provider').addEventListener('change', function() {
    const modelSelect = document.getElementById('model');
    const provider = this.value;
    
    modelSelect.innerHTML = '';
    
    if (provider === 'openai') {
        modelSelect.innerHTML = `
            <option value="gpt-4">GPT-4</option>
            <option value="gpt-4-turbo">GPT-4 Turbo</option>
            <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
        `;
    } else if (provider === 'anthropic') {
        modelSelect.innerHTML = `
            <option value="claude-3-haiku-20240307" selected>Claude 3 Haiku</option>
        `;
    }
});

document.getElementById('google_ai_provider').addEventListener('change', function() {
    const modelSelect = document.getElementById('google_model');
    const provider = this.value;
    
    modelSelect.innerHTML = '';
    
    if (provider === 'openai') {
        modelSelect.innerHTML = `
            <option value="gpt-4">GPT-4</option>
            <option value="gpt-4-turbo">GPT-4 Turbo</option>
            <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
        `;
    } else if (provider === 'anthropic') {
        modelSelect.innerHTML = `
            <option value="claude-3-haiku-20240307" selected>Claude 3 Haiku</option>
        `;
    }
});

// Contract Types Management
let currentMarkdownContent = '';

// Load contract types into dropdowns
async function loadContractTypes() {
    try {
        const response = await fetch('/api/contract-types');
        const data = await response.json();
        
        if (data.success) {
            const contractTypeSelect = document.getElementById('contract_type');
            const googleContractTypeSelect = document.getElementById('google_contract_type');
            
            // Clear existing options
            if (contractTypeSelect) {
                contractTypeSelect.innerHTML = '<option value="">Default</option>';
            }
            if (googleContractTypeSelect) {
                googleContractTypeSelect.innerHTML = '<option value="">Default</option>';
            }
            
            // Add contract types
            data.contract_types.forEach(type => {
                const option = document.createElement('option');
                option.value = type.id;
                option.textContent = type.name;
                if (contractTypeSelect) contractTypeSelect.appendChild(option.cloneNode(true));
                if (googleContractTypeSelect) googleContractTypeSelect.appendChild(option);
            });
            
            // Also refresh the contract types list in admin
            displayContractTypesList(data.contract_types);
        }
    } catch (error) {
        console.error('Error loading contract types:', error);
    }
}

// Display contract types list in admin
function displayContractTypesList(types) {
    const listDiv = document.getElementById('contract-types-list');
    if (!listDiv) return;
    
    if (!types || types.length === 0) {
        listDiv.innerHTML = '<p>No contract types found. Add one above.</p>';
        return;
    }
    
    listDiv.innerHTML = types.map(type => `
        <div class="contract-type-item" style="padding: 12px; border: 1px solid var(--border-color); border-radius: 8px; margin-bottom: 12px;">
            <div style="display: flex; justify-content: space-between; align-items: start;">
                <div>
                    <strong>${type.name}</strong>
                    ${type.description ? `<p style="margin: 4px 0; color: var(--text-secondary);">${type.description}</p>` : ''}
                    <small style="color: var(--text-secondary);">Playbook: ${type.playbook}</small>
                </div>
                <button class="btn btn-secondary" onclick="deleteContractType('${type.id}')" style="margin-left: 12px;">
                    <span>Delete</span>
                </button>
            </div>
        </div>
    `).join('');
}

// Add new contract type
async function addContractType() {
    const name = document.getElementById('contract_type_name').value.trim();
    const description = document.getElementById('contract_type_description').value.trim();
    const playbook = document.getElementById('contract_type_playbook').value.trim() || 'default_playbook.txt';
    
    if (!name) {
        showToast('Contract type name is required', 'error');
        return;
    }
    
    showLoading('Adding contract type...');
    
    try {
        const response = await fetch('/api/contract-types', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                name: name,
                description: description,
                playbook: playbook
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            showToast('Contract type added successfully!', 'success');
            // Clear form
            document.getElementById('contract_type_name').value = '';
            document.getElementById('contract_type_description').value = '';
            document.getElementById('contract_type_playbook').value = '';
            // Reload contract types
            await loadContractTypes();
        } else {
            showToast(data.error || 'Failed to add contract type', 'error');
        }
    } catch (error) {
        showToast('Error adding contract type: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// Delete contract type
async function deleteContractType(typeId) {
    if (!confirm('Are you sure you want to delete this contract type?')) {
        return;
    }
    
    showLoading('Deleting contract type...');
    
    try {
        const response = await fetch(`/api/contract-types/${typeId}`, {
            method: 'DELETE'
        });
        
        const data = await response.json();
        
        if (data.success) {
            showToast('Contract type deleted successfully!', 'success');
            await loadContractTypes();
        } else {
            showToast(data.error || 'Failed to delete contract type', 'error');
        }
    } catch (error) {
        showToast('Error deleting contract type: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// Convert playbook from Word to Markdown
async function convertPlaybook() {
    const fileInput = document.getElementById('playbook_converter_file');
    const file = fileInput.files[0];
    
    if (!file) {
        showToast('Please select a Word document to convert', 'error');
        return;
    }
    
    if (!file.name.endsWith('.docx')) {
        showToast('Only .docx files are supported', 'error');
        return;
    }
    
    showLoading('Converting playbook to Markdown...');
    
    try {
        const formData = new FormData();
        formData.append('playbook', file);
        
        const response = await fetch('/api/admin/convert-playbook', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            currentMarkdownContent = data.markdown;
            document.getElementById('converted_markdown').value = data.markdown;
            document.getElementById('converter-result').style.display = 'block';
            showToast('Playbook converted successfully!', 'success');
        } else {
            showToast(data.error || 'Conversion failed', 'error');
        }
    } catch (error) {
        showToast('Error converting playbook: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
}

// Download converted markdown
function downloadConvertedMarkdown() {
    if (!currentMarkdownContent) {
        showToast('No markdown content to download', 'error');
        return;
    }
    
    const blob = new Blob([currentMarkdownContent], { type: 'text/markdown' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'converted_playbook.md';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Copy markdown to clipboard
async function copyMarkdownToClipboard() {
    if (!currentMarkdownContent) {
        showToast('No markdown content to copy', 'error');
        return;
    }
    
    try {
        await navigator.clipboard.writeText(currentMarkdownContent);
        showToast('Markdown copied to clipboard!', 'success');
    } catch (error) {
        showToast('Failed to copy to clipboard: ' + error.message, 'error');
    }
}


// processGoogleDoc function (for Google Doc processing)
async function processGoogleDoc() {
    const docId = document.getElementById('doc_id').value.trim();
    const aiProvider = document.getElementById('google_ai_provider').value;
    const model = document.getElementById('google_model').value;
    const contractTypeId = document.getElementById('google_contract_type')?.value || '';
    
    if (!docId) {
        showToast('Google Doc ID is required', 'error');
        return;
    }
    
    showLoading('Processing Google Doc... This may take a few minutes.');
    
    try {
        const response = await fetch('/api/process-google', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                doc_id: docId,
                ai_provider: aiProvider,
                model: model,
                contract_type_id: contractTypeId || null
            })
        });
        
        const data = await response.json();
        
        if (response.ok) {
            showToast('Google Doc processed successfully!', 'success');
            displayResults(data);
        } else {
            showError(data.error || 'Processing failed');
        }
    } catch (error) {
        showError('Error processing Google Doc: ' + error.message);
    } finally {
        hideLoading();
    }
}

// Load contract types on page load and when admin tab is opened
document.addEventListener('DOMContentLoaded', function() {
    loadContractTypes();
    
    // Reload contract types when admin tab is opened
    const adminTabBtn = document.querySelector('[data-tab="admin"]');
    if (adminTabBtn) {
        adminTabBtn.addEventListener('click', function() {
            loadContractTypes();
        });
    }
});

