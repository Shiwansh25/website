let excelData = null;
let workbook = null;
let apiKey = localStorage.getItem('gemini_api_key');

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const chatSection = document.getElementById('chatSection');
const chatMessages = document.getElementById('chatMessages');
const userInput = document.getElementById('userInput');
const sendButton = document.getElementById('sendButton');

// Check for API key on startup
if (!apiKey) {
    addMessage('Welcome! Before you can use the AI features, please enter your Google Gemini API key:', 'bot');
    const apiKeyInput = document.createElement('div');
    apiKeyInput.className = 'api-key-input';
    apiKeyInput.innerHTML = `
        <input type="password" id="apiKeyInput" placeholder="Enter your Google Gemini API key">
        <button id="saveApiKey">Save API Key</button>
    `;
    chatMessages.appendChild(apiKeyInput);

    document.getElementById('saveApiKey').addEventListener('click', () => {
        const key = document.getElementById('apiKeyInput').value.trim();
        if (key) {
            apiKey = key;
            localStorage.setItem('gemini_api_key', key);
            apiKeyInput.remove();
            addMessage('API key saved successfully! You can now ask questions about your Excel data.', 'bot');
        }
    });
}

// Drag and drop event listeners
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                 file.type === 'application/vnd.ms-excel')) {
        handleFile(file);
    }
});

// File input change event
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
});

// Handle file upload
function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(firstSheet);
            
            // Show chat section and hide upload section
            document.querySelector('.upload-section').style.display = 'none';
            chatSection.style.display = 'block';
            
            // Add success message
            addMessage('Excel file loaded successfully! You can now ask questions about your data.', 'bot');
        } catch (error) {
            addMessage('Error reading the Excel file. Please make sure it\'s a valid Excel file.', 'bot');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Add message to chat
function addMessage(text, sender) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender}`;
    messageDiv.textContent = text;
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Handle user input
async function handleUserInput() {
    const question = userInput.value.trim();
    if (!question) return;

    // Check for API key
    if (!apiKey) {
        addMessage('Please enter your Google Gemini API key to use the AI features:', 'bot');
        const apiKeyInput = document.createElement('div');
        apiKeyInput.className = 'api-key-input';
        apiKeyInput.innerHTML = `
            <input type="password" id="apiKeyInput" placeholder="Enter your Google Gemini API key">
            <button id="saveApiKey">Save API Key</button>
        `;
        chatMessages.appendChild(apiKeyInput);

        document.getElementById('saveApiKey').addEventListener('click', () => {
            const key = document.getElementById('apiKeyInput').value.trim();
            if (key) {
                apiKey = key;
                localStorage.setItem('gemini_api_key', key);
                apiKeyInput.remove();
                addMessage('API key saved successfully! You can now ask your question again.', 'bot');
            }
        });
        return;
    }

    // Add user message
    addMessage(question, 'user');
    userInput.value = '';

    // Show loading message
    const loadingMessage = document.createElement('div');
    loadingMessage.className = 'message bot';
    loadingMessage.textContent = 'Thinking...';
    chatMessages.appendChild(loadingMessage);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    try {
        // Process the question with AI
        const response = await processQuestionWithAI(question);
        // Remove loading message
        loadingMessage.remove();
        // Add AI response
        addMessage(response, 'bot');
    } catch (error) {
        // Remove loading message
        loadingMessage.remove();
        // Add error message with more details
        if (error.message.includes('401') || error.message.includes('403')) {
            addMessage('Invalid API key. Please enter a valid Google Gemini API key:', 'bot');
            const apiKeyInput = document.createElement('div');
            apiKeyInput.className = 'api-key-input';
            apiKeyInput.innerHTML = `
                <input type="password" id="apiKeyInput" placeholder="Enter your Google Gemini API key">
                <button id="saveApiKey">Save API Key</button>
            `;
            chatMessages.appendChild(apiKeyInput);

            document.getElementById('saveApiKey').addEventListener('click', () => {
                const key = document.getElementById('apiKeyInput').value.trim();
                if (key) {
                    apiKey = key;
                    localStorage.setItem('gemini_api_key', key);
                    apiKeyInput.remove();
                    addMessage('API key saved successfully! You can now ask your question again.', 'bot');
                }
            });
        } else {
            addMessage('Sorry, I encountered an error while processing your question. Please try again. Error: ' + error.message, 'bot');
        }
    }
}

// Process the question with AI
async function processQuestionWithAI(question) {
    if (!excelData) {
        return 'Please upload an Excel file first.';
    }

    // Prepare the data context for the AI
    const dataContext = prepareDataContext();
    
    // Prepare the prompt for the AI
    const prompt = `You are an AI assistant helping users analyze Excel data. Here is the context about the data:

${dataContext}

User Question: ${question}

Please provide a clear and concise answer based on the Excel data. If the question cannot be answered with the available data, please explain why.`;

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{
                        text: prompt
                    }]
                }],
                generationConfig: {
                    temperature: 0.7,
                    maxOutputTokens: 500,
                },
                safetySettings: [
                    {
                        category: "HARM_CATEGORY_HARASSMENT",
                        threshold: "BLOCK_NONE"
                    },
                    {
                        category: "HARM_CATEGORY_HATE_SPEECH",
                        threshold: "BLOCK_NONE"
                    },
                    {
                        category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                        threshold: "BLOCK_NONE"
                    },
                    {
                        category: "HARM_CATEGORY_DANGEROUS_CONTENT",
                        threshold: "BLOCK_NONE"
                    }
                ]
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(`HTTP error! status: ${response.status}, message: ${errorData.error?.message || 'Unknown error'}`);
        }

        const data = await response.json();
        if (!data.candidates || !data.candidates[0]?.content?.parts?.[0]?.text) {
            throw new Error('Invalid response format from Gemini API');
        }
        return data.candidates[0].content.parts[0].text;
    } catch (error) {
        console.error('Error calling Gemini API:', error);
        throw error;
    }
}

// Prepare data context for AI
function prepareDataContext() {
    if (!excelData || excelData.length === 0) {
        return 'No data available.';
    }

    const columns = Object.keys(excelData[0]);
    const sampleData = excelData.slice(0, 5); // Get first 5 rows as sample

    let context = `The Excel file contains ${excelData.length} rows with the following columns: ${columns.join(', ')}.\n\n`;
    context += 'Sample data (first 5 rows):\n';
    
    sampleData.forEach((row, index) => {
        context += `Row ${index + 1}: ${Object.entries(row).map(([key, value]) => `${key}: ${value}`).join(', ')}\n`;
    });

    return context;
}

// Event listeners for chat input
sendButton.addEventListener('click', handleUserInput);
userInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        handleUserInput();
    }
}); 