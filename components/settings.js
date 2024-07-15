import { loadModels } from '../services/api.js';

export function initSettings() {
    const settingsTab = document.getElementById('settings');
    settingsTab.innerHTML = `
        <div class="mt-3">
            <label for="api-key" class="form-label">API Ключ:</label>
            <input type="password" id="api-key" class="form-control">
        </div>
    `;

    document.getElementById('api-key').onchange = saveApiKey;
    loadModels();
    loadSettings();
}

export function populateModelSelect(models) {
    const select = document.getElementById('prompt-model');
    select.innerHTML = '';
    models.forEach(model => {
        const option = document.createElement('option');
        option.value = model.id;
        option.textContent = model.name;
        select.appendChild(option);
    });
}

function saveApiKey() {
    const apiKey = document.getElementById('api-key').value;
    localStorage.setItem('apiKey', apiKey);
}

function loadSettings() {
    const apiKey = localStorage.getItem('apiKey') || '';
    document.getElementById('api-key').value = apiKey;
}
