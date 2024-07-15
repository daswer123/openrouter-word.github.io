import { getPrompts, getLastSelectedPromptIndex, saveLastSelectedPromptIndex } from '../services/storage.js';
import { generateTextWithAI } from '../services/api.js';

let originalText = "";
let currentController = null;
let lastSelectedPrompt = null;

export function initHome() {
    const homeTab = document.getElementById('home');
    homeTab.innerHTML = `
        <select id="prompt-select" class="form-select mt-3"></select>
        <div id="selected-model" class="mt-2"></div>
        <button id="replace-text" class="btn btn-primary mt-3" disabled>Заменить выделенный текст</button>
        <button id="cancel-request" class="btn btn-secondary mt-3" style="display: none;">Отменить запрос</button>
    `;

    $('#prompt-select').select2({
        width: '100%',
        placeholder: 'Выберите промпт',
        templateResult: formatPrompt,
        templateSelection: formatPromptSelection,
    });

    document.getElementById("replace-text").onclick = replaceSelectedText;
    document.getElementById("cancel-request").onclick = cancelRequest;

    updatePromptSelect();
    setInterval(checkSelection, 1000);
}

function formatPrompt(prompt) {
    if (!prompt.id) {
        return prompt.text;
    }
    const prompts = getPrompts();
    const promptData = prompts[prompt.id];
    return $(`<span>${promptData.name} (${promptData.model})</span>`);
}

function formatPromptSelection(prompt) {
    if (!prompt.id) {
        return prompt.text;
    }
    const prompts = getPrompts();
    const promptData = prompts[prompt.id];
    return promptData.name;
}

export function updatePromptSelect() {
    const select = $('#prompt-select');
    const modelDiv = document.getElementById('selected-model');
    const prompts = getPrompts();
    select.empty();
    select.append(new Option('Выберите промпт', '', true, true));
    const lastSelectedPromptIndex = getLastSelectedPromptIndex();
    prompts.forEach((prompt, index) => {
        const option = new Option(prompt.name, index, false, index === lastSelectedPromptIndex);
        select.append(option);
    });
    select.trigger('change');
    const promptIndex = lastSelectedPromptIndex >= 0 && lastSelectedPromptIndex < prompts.length ? lastSelectedPromptIndex : null;
    lastSelectedPrompt = promptIndex !== null ? prompts[promptIndex] : null;
    modelDiv.textContent = lastSelectedPrompt ? `Модель: ${lastSelectedPrompt.model}` : '';
    document.getElementById("replace-text").disabled = !lastSelectedPrompt;
    select.on('change', function() {
        const newPromptIndex = $(this).val();
        lastSelectedPrompt = newPromptIndex ? prompts[newPromptIndex] : null;
        saveLastSelectedPromptIndex(newPromptIndex || -1);
        modelDiv.textContent = lastSelectedPrompt ? `Модель: ${lastSelectedPrompt.model}` : '';
        document.getElementById("replace-text").disabled = !lastSelectedPrompt;
    });
}

function checkSelection() {
    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        const replaceButton = document.getElementById("replace-text");
        if (selection.text.trim().length > 0) {
            replaceButton.disabled = false;
        } else {
            replaceButton.disabled = true;
        }
    });
}

async function replaceSelectedText() {
    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        originalText = selection.text;

        if (!lastSelectedPrompt) return;

        document.getElementById("replace-text").style.display = "none";
        document.getElementById("cancel-request").style.display = "inline-block";

        try {
            currentController = new AbortController();
            const fullPrompt = `${lastSelectedPrompt.text}\n\nUser provided the following text: "${originalText}"`;
            const generatedText = await generateTextWithAI(fullPrompt, lastSelectedPrompt.model, currentController.signal);
            selection.insertText(generatedText + '\n', Word.InsertLocation.replace);
        } catch (error) {
            if (error.name !== 'AbortError') {
                console.error('Error:', error);
            }
        } finally {
            document.getElementById("cancel-request").style.display = "none";
            document.getElementById("replace-text").style.display = "inline-block";
            currentController = null;
        }
    });
}

function cancelRequest() {
    if (currentController) {
        currentController.abort();
        document.getElementById("cancel-request").style.display = "none";
        document.getElementById("replace-text").style.display = "inline-block";
    }
}
