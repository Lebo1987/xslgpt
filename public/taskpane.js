// XSLGPT Excel Add-in Main JavaScript

Office.onReady((info) => {
    console.log('Office.js loaded, host:', info.host);
    if (info.host === Office.HostType.Excel) {
        document.getElementById('sideload-msg').style.display = 'none';
        document.getElementById('app-body').style.display = 'flex';
    }
});

// Initialize the add-in
Office.onReady(() => {
    console.log('Initializing XSLGPT add-in...');
    initializeAddin();
});

// Backend API configuration
const API_BASE_URL = window.location.origin; // Automatically detect server URL
console.log('API Base URL:', API_BASE_URL);

// Configuration options
const CONFIG = {
    ALWAYS_SET_NUMBER_FORMAT: false, // Set to true to always apply number formatting
    DEFAULT_DATE_FORMAT: 'mm/dd/yyyy', // Default date format
    DEFAULT_NUMBER_FORMAT: 'General', // Default number format for non-date formulas
    // Additional date formats for different regions
    DATE_FORMATS: {
        'US': 'mm/dd/yyyy',
        'EUROPE': 'dd/mm/yyyy',
        'ISO': 'yyyy-mm-dd',
        'SHORT': 'm/d/yy',
        'LONG': 'mmmm dd, yyyy'
    }
};

function initializeAddin() {
    console.log('Setting up add-in...');
    // Set up event listeners
    setupEventListeners();
    
    // Load prompt history
    loadPromptHistory();
    
    // Update button state
    updateGenerateButtonState();
    console.log('Add-in initialization complete');
}

function setupEventListeners() {
    console.log('Setting up event listeners...');
    // Generate button
    const generateBtn = document.getElementById('generate-btn');
    generateBtn.addEventListener('click', handleGenerateFormula);
    
    // Prompt input
    const promptInput = document.getElementById('prompt-input');
    promptInput.addEventListener('input', updateGenerateButtonState);
    promptInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.ctrlKey) {
            handleGenerateFormula();
        }
    });
    

    
    console.log('Event listeners set up successfully');
}

function updateGenerateButtonState() {
    const promptInput = document.getElementById('prompt-input');
    const generateBtn = document.getElementById('generate-btn');
    
    const hasPrompt = promptInput.value.trim().length > 0;
    const wasDisabled = generateBtn.disabled;
    generateBtn.disabled = !hasPrompt;
    
    // Log state changes for debugging
    if (wasDisabled !== generateBtn.disabled) {
        console.log('Generate button state changed:', {
            disabled: generateBtn.disabled,
            hasPrompt: hasPrompt,
            promptLength: promptInput.value.trim().length
        });
    }
}

async function handleGenerateFormula() {
    console.log('=== Starting formula generation ===');
    const promptInput = document.getElementById('prompt-input');
    const prompt = promptInput.value.trim();
    
    console.log('User prompt:', prompt);
    
    if (!prompt) {
        console.warn('Empty prompt provided');
        showError('Please enter a prompt');
        return;
    }
    
    // Show loading state
    setLoadingState(true);
    
    try {
        console.log('Calling backend API...');
        const result = await generateFormula(prompt);
        console.log('Backend API response:', result);
        
        if (result.success) {
            console.log('Formula received successfully:', result.formula);
            
            // Display formula in UI for feedback
            displayFormulaResult(result.formula, result.explanation);
            
            // Add to history
            addToPromptHistory(prompt, result.formula, result.explanation);
            
            // Clear input
            promptInput.value = '';
            updateGenerateButtonState();
            
            showSuccess('Formula generated successfully! Click "Insert Formula" to add it to Excel.');
            console.log('=== Formula generation completed successfully ===');
        } else {
            console.error('Backend returned error:', result.error);
            showError(result.error || 'Failed to generate formula');
        }
    } catch (error) {
        console.error('Error in handleGenerateFormula:', error);
        showError('An error occurred while generating the formula');
    } finally {
        setLoadingState(false);
    }
}

async function handleInsertFormula() {
    console.log('=== Starting formula insertion ===');
    const insertBtn = document.getElementById('insert-formula-btn');
    const formula = insertBtn.dataset.formula;
    
    if (!formula) {
        console.warn('No formula available to insert');
        showError('No formula available to insert');
        return;
    }
    
    console.log('Inserting formula into Excel:', formula);
    
    try {
        await insertFormulaIntoExcel(formula);
        console.log('Formula inserted successfully');
        showSuccess('Formula inserted into Excel!');
        
        // Disable the button after successful insertion
        if (insertBtn) {
            insertBtn.disabled = true;
            insertBtn.textContent = '✓ Inserted';
            insertBtn.style.background = '#107c10';
        }
    } catch (error) {
        console.error('Error inserting formula:', error);
        showError('Failed to insert formula into Excel. Please try again.');
    }
}

async function generateFormula(prompt) {
    console.log('Making API request to:', `${API_BASE_URL}/api/generate`);
    try {
        const response = await fetch(`${API_BASE_URL}/api/generate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                prompt: prompt
            })
        });
        
        console.log('API response status:', response.status);
        
        if (!response.ok) {
            const errorData = await response.json();
            console.error('API error response:', errorData);
            throw new Error(errorData.error || `HTTP ${response.status}`);
        }
        
        const data = await response.json();
        console.log('API response data:', data);
        
        if (!data.success) {
            return {
                success: false,
                error: data.error || 'Failed to generate formula'
            };
        }
        
        return {
            success: true,
            formula: data.formula,
            explanation: data.explanation
        };
        
    } catch (error) {
        console.error('Backend API error:', error);
        return {
            success: false,
            error: error.message
        };
    }
}

async function insertFormulaIntoExcel(formula) {
    console.log('=== Starting Excel formula insertion ===');
    console.log('Formula to insert:', formula);
    
    return new Promise((resolve, reject) => {
        try {
            // Check if we're in Excel Desktop
            if (typeof Excel === 'undefined') {
                const error = 'Excel object not available - not running in Excel Desktop';
                console.error(error);
                reject(new Error(error));
                return;
            }
            
            console.log('Excel object available, running Excel.run()...');
            
            Excel.run(async (context) => {
                try {
                    console.log('Getting selected range...');
                    const range = context.workbook.getSelectedRange();
                    
                    // Load the range address for logging
                    range.load('address');
                    await context.sync();
                    
                    console.log('Selected cell address:', range.address);
                    
                    console.log('Setting formula:', formula);
                    range.formulas = [[formula]];
                    
                    // Set number format based on configuration and formula type
                    if (CONFIG.ALWAYS_SET_NUMBER_FORMAT) {
                        // Always set number format
                        if (isDateFormula(formula)) {
                            console.log('Setting number format to date format');
                            range.numberFormat = [[CONFIG.DEFAULT_DATE_FORMAT]];
                        } else {
                            console.log('Setting number format to default format');
                            range.numberFormat = [[CONFIG.DEFAULT_NUMBER_FORMAT]];
                        }
                    } else {
                        // Only set number format for date formulas
                        if (isDateFormula(formula)) {
                            console.log('Detected date formula, setting number format to date');
                            range.numberFormat = [[CONFIG.DEFAULT_DATE_FORMAT]];
                        }
                    }
                    
                    console.log('Syncing changes...');
                    await context.sync();
                    
                    console.log('Formula inserted successfully into cell:', range.address);
                    resolve();
                    
                } catch (excelError) {
                    console.error('Error in Excel.run():', excelError);
                    reject(excelError);
                }
            });
            
        } catch (error) {
            console.error('Error setting up Excel.run():', error);
            reject(error);
        }
    });
}

function isDateFormula(formula) {
    // Convert to uppercase for case-insensitive matching
    const upperFormula = formula.toUpperCase();
    
    // List of Excel functions that return dates
    const dateFunctions = [
        'TODAY()',
        'NOW()',
        'DATE(',
        'EDATE(',
        'EOMONTH(',
        'WORKDAY(',
        'WORKDAY.INTL(',
        'NETWORKDAYS(',
        'NETWORKDAYS.INTL(',
        'YEAR(',
        'MONTH(',
        'DAY(',
        'WEEKDAY(',
        'WEEKNUM(',
        'ISOWEEKNUM(',
        'QUARTER(',
        'DATEDIF(',
        'DAYS(',
        'DAYS360(',
        'YEARFRAC(',
        'DATEVALUE(',
        'TIMEVALUE(',
        'HOUR(',
        'MINUTE(',
        'SECOND(',
        'TIME(',
        'NOW()',
        'TODAY()'
    ];
    
    // Check if the formula contains any date functions
    for (const dateFunc of dateFunctions) {
        if (upperFormula.includes(dateFunc)) {
            console.log('Date function detected:', dateFunc);
            return true;
        }
    }
    
    // Also check for common date patterns
    const datePatterns = [
        /TODAY\(\)/i,
        /NOW\(\)/i,
        /DATE\(/i,
        /EDATE\(/i,
        /EOMONTH\(/i,
        /WORKDAY\(/i,
        /NETWORKDAYS\(/i,
        /YEAR\(/i,
        /MONTH\(/i,
        /DAY\(/i,
        /WEEKDAY\(/i,
        /WEEKNUM\(/i,
        /ISOWEEKNUM\(/i,
        /QUARTER\(/i,
        /DATEDIF\(/i,
        /DAYS\(/i,
        /DAYS360\(/i,
        /YEARFRAC\(/i,
        /DATEVALUE\(/i,
        /TIMEVALUE\(/i,
        /HOUR\(/i,
        /MINUTE\(/i,
        /SECOND\(/i,
        /TIME\(/i
    ];
    
    for (const pattern of datePatterns) {
        if (pattern.test(formula)) {
            console.log('Date pattern detected in formula');
            return true;
        }
    }
    
    console.log('No date function detected in formula');
    return false;
}

function displayFormulaResult(formula, explanation) {
    console.log('Displaying formula result in UI:', { formula, explanation });
    
    // Create or update the result display
    let resultDiv = document.getElementById('formula-result');
    if (!resultDiv) {
        resultDiv = document.createElement('div');
        resultDiv.id = 'formula-result';
        resultDiv.className = 'formula-result';
        
        // Insert after the loading div
        const loadingDiv = document.getElementById('loading');
        loadingDiv.parentNode.insertBefore(resultDiv, loadingDiv.nextSibling);
    }
    
    resultDiv.innerHTML = `
        <div class="result-header">
            <h4>Generated Formula:</h4>
        </div>
        <div class="formula-display">
            <code>${escapeHtml(formula)}</code>
        </div>
        ${explanation ? `<div class="explanation">${escapeHtml(explanation)}</div>` : ''}
        <div class="insert-formula-container">
            <button id="insert-formula-btn" class="insert-formula-button" data-formula="${escapeHtml(formula)}">
                <span class="button-icon">📊</span>
                <span class="button-text">Insert Formula</span>
            </button>
        </div>
    `;
    
    // Add event listener to the new button
    const insertBtn = document.getElementById('insert-formula-btn');
    if (insertBtn) {
        insertBtn.addEventListener('click', handleInsertFormula);
    }
    
    // Auto-hide after 10 seconds
    setTimeout(() => {
        if (resultDiv && resultDiv.parentNode) {
            resultDiv.style.opacity = '0.7';
        }
    }, 10000);
}

function addToPromptHistory(prompt, formula, explanation) {
    console.log('Adding to prompt history:', { prompt, formula });
    const history = getPromptHistory();
    
    // Add new entry to the beginning
    history.unshift({
        prompt: prompt,
        formula: formula,
        explanation: explanation,
        timestamp: new Date().toISOString()
    });
    
    // Keep only the last 10 entries
    if (history.length > 10) {
        history.splice(10);
    }
    
    // Save to localStorage
    try {
        localStorage.setItem('xslgpt_prompt_history', JSON.stringify(history));
        console.log('Prompt history saved to localStorage');
    } catch (error) {
        console.error('Error saving prompt history:', error);
    }
    
    // Update UI and counter immediately
    updatePromptHistoryUI();
    
    // Also update counter directly for immediate feedback
    const historyCountElement = document.getElementById('history-count');
    if (historyCountElement) {
        historyCountElement.textContent = history.length;
        console.log('Immediately updated history counter to:', history.length);
    }
}

function getPromptHistory() {
    try {
        const history = localStorage.getItem('xslgpt_prompt_history');
        return history ? JSON.parse(history) : [];
    } catch (error) {
        console.error('Error loading prompt history:', error);
        return [];
    }
}

function loadPromptHistory() {
    console.log('Loading prompt history...');
    updatePromptHistoryUI();
    
    // Ensure counter is properly initialized
    const history = getPromptHistory();
    const historyCountElement = document.getElementById('history-count');
    if (historyCountElement) {
        historyCountElement.textContent = history.length;
        console.log('Initialized history counter to:', history.length);
    }
}

function updatePromptHistoryUI() {
    const historyContainer = document.getElementById('prompt-history');
    const historyCountElement = document.getElementById('history-count');
    const history = getPromptHistory();
    
    console.log('Updating history UI with', history.length, 'items');
    
    // Update the counter
    if (historyCountElement) {
        historyCountElement.textContent = history.length;
        console.log('Updated history counter to:', history.length);
    }
    
    if (history.length === 0) {
        historyContainer.innerHTML = '<p style="color: #605e5c; font-style: italic; text-align: center; padding: 20px;">No previous prompts yet</p>';
        return;
    }
    
    historyContainer.innerHTML = history.map((item, index) => `
        <div class="history-item" onclick="reusePrompt('${index}')">
            <div class="prompt-text">${escapeHtml(item.prompt)}</div>
            <div class="formula-text">${escapeHtml(item.formula)}</div>
        </div>
    `).join('');
}

function reusePrompt(index) {
    console.log('Reusing prompt at index:', index);
    const history = getPromptHistory();
    if (history[index]) {
        const item = history[index];
        document.getElementById('prompt-input').value = item.prompt;
        updateGenerateButtonState();
        console.log('Prompt reused:', item.prompt);
    }
}

function setLoadingState(isLoading) {
    console.log('Setting loading state:', isLoading);
    const generateBtn = document.getElementById('generate-btn');
    const loadingDiv = document.getElementById('loading');
    const promptInput = document.getElementById('prompt-input');
    
    if (isLoading) {
        generateBtn.disabled = true;
        loadingDiv.classList.remove('hidden');
        promptInput.disabled = true;
    } else {
        loadingDiv.classList.add('hidden');
        promptInput.disabled = false;
        updateGenerateButtonState();
    }
}

function showSuccess(message) {
    console.log('Success:', message);
    // You could implement a proper notification system here
}

function showError(message) {
    console.error('Error:', message);
    alert(message); // Simple alert for now
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function clearPromptHistory() {
    console.log('Clearing prompt history...');
    try {
        localStorage.removeItem('xslgpt_prompt_history');
        updatePromptHistoryUI();
        console.log('Prompt history cleared successfully');
    } catch (error) {
        console.error('Error clearing prompt history:', error);
    }
} 