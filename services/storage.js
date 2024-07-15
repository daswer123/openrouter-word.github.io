export function getPrompts() {
    return JSON.parse(localStorage.getItem('prompts') || '[]');
}

export function savePrompts(prompts) {
    localStorage.setItem('prompts', JSON.stringify(prompts));
}

export function getLastSelectedPromptIndex() {
    return parseInt(localStorage.getItem('lastSelectedPromptIndex') || '-1', 10);
}

export function saveLastSelectedPromptIndex(index) {
    localStorage.setItem('lastSelectedPromptIndex', index.toString());
}
