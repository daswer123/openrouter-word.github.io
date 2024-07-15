import { getPrompts, savePrompts } from '../services/storage.js';
import { updatePromptSelect } from './home.js';

export function initPrompts() {
    const promptsTab = document.getElementById('prompts');
    promptsTab.innerHTML = `
        <div class="mt-3">
            <input type="text" id="prompt-name" class="form-control" placeholder="Название промпта">
            <select id="prompt-model" class="form-select mt-2"></select>
            <textarea id="prompt-text" class="form-control mt-2" rows="5" placeholder="Текст промпта"></textarea>
            <button id="save-prompt" class="btn btn-primary mt-2">Сохранить промпт</button>
        </div>
        <div id="prompt-list" class="mt-3"></div>
    `;

    $('#prompt-model').select2({
        width: '100%',
        placeholder: 'Выберите модель',
    }).on('select2:opening', function() {
        setTimeout(() => {
            $('.select2-search__field').get(0).focus();
        }, 0);
    });

    document.getElementById("save-prompt").onclick = savePrompt;
    updatePromptList();
    populateModelSelect();
}

function populateModelSelect() {
    const select = $('#prompt-model');
    const models = JSON.parse(localStorage.getItem('models') || '[]');
    models.forEach(model => {
        const option = new Option(model.name, model.id, false, false);
        select.append(option);
    });
    select.trigger('change');
}

function savePrompt() {
    const name = document.getElementById('prompt-name').value;
    const model = $('#prompt-model').val();
    const text = document.getElementById('prompt-text').value;
    if (name && model && text) {
        const prompts = getPrompts();
        prompts.push({ name, model, text });
        savePrompts(prompts);
        updatePromptList();
        updatePromptSelect();
        document.getElementById('prompt-name').value = '';
        $('#prompt-model').val(null).trigger('change');
        document.getElementById('prompt-text').value = '';
    }
}

function updatePromptList() {
    const list = document.getElementById('prompt-list');
    const prompts = getPrompts();
    list.innerHTML = '';
    prompts.forEach((prompt, index) => {
        const div = document.createElement('div');
        div.className = 'mt-2';
        div.innerHTML = `
            <strong>${prompt.name}</strong> (${prompt.model})
            <button class="btn btn-sm btn-primary edit-prompt" data-index="${index}">Редактировать</button>
            <button class="btn btn-sm btn-info duplicate-prompt" data-index="${index}">Дубликат</button>
            <button class="btn btn-sm btn-danger delete-prompt" data-index="${index}">Удалить</button>
        `;
        list.appendChild(div);
    });

    list.querySelectorAll('.delete-prompt').forEach(button => {
        button.onclick = () => deletePrompt(button.dataset.index);
    });

    list.querySelectorAll('.edit-prompt').forEach(button => {
        button.onclick = () => editPrompt(button.dataset.index);
    });

    list.querySelectorAll('.duplicate-prompt').forEach(button => {
        button.onclick = () => duplicatePrompt(button.dataset.index);
    });
}

function deletePrompt(index) {
    const prompts = getPrompts();
    prompts.splice(index, 1);
    savePrompts(prompts);
    updatePromptList();
    updatePromptSelect();
}

function editPrompt(index) {
    const prompts = getPrompts();
    const prompt = prompts[index];
    document.getElementById('prompt-name').value = prompt.name;
    $('#prompt-model').val(prompt.model).trigger('change');
    document.getElementById('prompt-text').value = prompt.text;
    document.getElementById('save-prompt').textContent = 'Обновить промпт';
    document.getElementById('save-prompt').onclick = () => {
        updatePrompt(index);
        document.getElementById('save-prompt').textContent = 'Сохранить промпт';
        document.getElementById('save-prompt').onclick = savePrompt;
    };
}

function updatePrompt(index) {
    const name = document.getElementById('prompt-name').value;
    const model = $('#prompt-model').val();
    const text = document.getElementById('prompt-text').value;
    if (name && model && text) {
        const prompts = getPrompts();
        prompts[index] = { name, model, text };
        savePrompts(prompts);
        updatePromptList();
        updatePromptSelect();
        document.getElementById('prompt-name').value = '';
        $('#prompt-model').val(null).trigger('change');
        document.getElementById('prompt-text').value = '';
    }
}

function duplicatePrompt(index) {
    const prompts = getPrompts();
    const prompt = prompts[index];
    document.getElementById('prompt-name').value = prompt.name + ' (копия)';
    $('#prompt-model').val(null).trigger('change');
    document.getElementById('prompt-text').value = prompt.text;
    document.getElementById('save-prompt').textContent = 'Сохранить промпт';
    document.getElementById('save-prompt').onclick = savePrompt;
}
