import axios from 'axios';
import { populateModelSelect } from '../components/settings.js';

const OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1";

export async function loadModels() {
    try {
        const response = await axios.get(`${OPENROUTER_BASE_URL}/models`);
        const models = response.data.data;
        localStorage.setItem('models', JSON.stringify(models));
        populateModelSelect(models);
    } catch (error) {
        console.error('Error loading models:', error);
    }
}

export async function generateTextWithAI(prompt, model, signal) {
    const apiKey = localStorage.getItem('apiKey');

    try {
        const response = await fetch(`${OPENROUTER_BASE_URL}/chat/completions`, {
            method: 'POST',
            signal,
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json',
                'HTTP-Referer': 'https://www.office.com',
                'X-Title': 'Word AI Assistant',
            },
            body: JSON.stringify({
                model: model,
                messages: [{ role: "user", content: prompt }],
            }),
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0].message.content.trim();
    } catch (error) {
        console.error('Error in generateTextWithAI:', error);
        throw error;
    }
}
