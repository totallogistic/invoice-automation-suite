// ============================================================================
// Form Generator - Dynamic forms from JSON Schema
// ============================================================================

console.log('📋 Schema loaded:', SCHEMA);

const formContainer = document.getElementById('form-container');
const btnSave = document.getElementById('btn-save');
const btnValidate = document.getElementById('btn-validate');
const btnReset = document.getElementById('btn-reset');
const btnHelp = document.getElementById('btn-help');
const helpPanel = document.getElementById('keyboard-help');
const errorsDiv = document.getElementById('validation-errors');
const successDiv = document.getElementById('success-message');

// State
let formData = {};

// ============================================================================
// Form Generation
// ============================================================================

function renderForm() {
    const html = generateFieldsHTML(SCHEMA.properties || {}, SCHEMA.required || [], '');
    formContainer.innerHTML = `<div class="form-fields">${html}</div>`;
    
    // Attach event listeners
    attachEventListeners();
    
    // Load from localStorage if exists
    loadFromStorage();
}

function generateFieldsHTML(properties, required, prefix) {
    let html = '';
    
    for (const [key, schema] of Object.entries(properties)) {
        const fieldPath = prefix ? `${prefix}.${key}` : key;
        const isRequired = required.includes(key);
        const label = schema.title || key;
        const desc = schema.description || '';
        
        html += `<div class="form-group">`;
        html += `<label for="${fieldPath}">
            ${label}
            ${isRequired ? '<span class="required">*</span>' : ''}
        </label>`;
        
        if (schema.type === 'object') {
            html += `<div class="nested-object">`;
            html += generateFieldsHTML(
                schema.properties || {}, 
                schema.required || [], 
                fieldPath
            );
            html += `</div>`;
        } else if (schema.type === 'array') {
            html += generateArrayField(fieldPath, schema);
        } else {
            html += generateScalarField(fieldPath, schema);
        }
        
        if (desc) {
            html += `<small class="field-description">${desc}</small>`;
        }
        
        html += `</div>`;
    }
    
    return html;
}

function generateScalarField(path, schema) {
    const type = schema.type || 'string';
    let inputHTML = '';
    
    if (schema.enum) {
        // Enum → Select
        inputHTML = `<select name="${path}" id="${path}" class="form-input">
            <option value="">-- Seleccionar --</option>`;
        for (const option of schema.enum) {
            inputHTML += `<option value="${option}">${option}</option>`;
        }
        inputHTML += `</select>`;
    } else if (type === 'boolean') {
        // Boolean → Checkbox
        inputHTML = `<input type="checkbox" name="${path}" id="${path}" class="form-checkbox">`;
    } else if (type === 'integer' || type === 'number') {
        // Number
        const min = schema.minimum !== undefined ? `min="${schema.minimum}"` : '';
        const max = schema.maximum !== undefined ? `max="${schema.maximum}"` : '';
        inputHTML = `<input type="number" name="${path}" id="${path}" 
            class="form-input" ${min} ${max} step="${type === 'integer' ? '1' : 'any'}">`;
    } else if (schema.format === 'date') {
        // Date
        inputHTML = `<input type="date" name="${path}" id="${path}" class="form-input">`;
    } else {
        // String (default)
        const maxLength = schema.maxLength ? `maxlength="${schema.maxLength}"` : '';
        inputHTML = `<input type="text" name="${path}" id="${path}" 
            class="form-input" ${maxLength}>`;
    }
    
    return inputHTML;
}

function generateArrayField(path, schema) {
    const itemSchema = schema.items || {};
    const isObjectArray = itemSchema.type === 'object';
    
    return `
        <div class="array-container" data-path="${path}">
            <div class="array-items" id="${path}-items"></div>
            <button type="button" class="btn-add-item" data-path="${path}">
                ➕ Añadir ${schema.title || 'Item'}
            </button>
        </div>
    `;
}

// ============================================================================
// Event Listeners
// ============================================================================

function attachEventListeners() {
    // Input changes → Update state
    formContainer.querySelectorAll('input, select, textarea').forEach(input => {
        input.addEventListener('change', handleInputChange);
        input.addEventListener('input', handleInputChange);
    });
    
    // Array add buttons
    formContainer.querySelectorAll('.btn-add-item').forEach(btn => {
        btn.addEventListener('click', handleAddArrayItem);
    });
    
    // Keyboard navigation
    formContainer.addEventListener('keydown', handleKeyboardNav);
}

function handleInputChange(e) {
    const path = e.target.name;
    let value = e.target.value;
    
    // Type conversion
    if (e.target.type === 'checkbox') {
        value = e.target.checked;
    } else if (e.target.type === 'number') {
        value = parseFloat(value) || 0;
    }
    
    setValueAtPath(formData, path, value);
    saveToStorage();
}

function handleAddArrayItem(e) {
    const path = e.target.dataset.path;
    // TODO: Implement array item addition
    console.log('Add item to:', path);
}

function handleKeyboardNav(e) {
    // Enter → Next field
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
        e.preventDefault();
        focusNextField(e.target);
    }
}

// ============================================================================
// Data Management
// ============================================================================

function collectFormData() {
    const data = {};
    
    formContainer.querySelectorAll('input, select, textarea').forEach(input => {
        const path = input.name;
        if (!path) return;
        
        let value = input.value;
        
        if (input.type === 'checkbox') {
            value = input.checked;
        } else if (input.type === 'number') {
            value = parseFloat(value) || 0;
        } else if (input.type === 'date') {
            value = value; // Keep as string in YYYY-MM-DD format
        }
        
        setValueAtPath(data, path, value);
    });
    
    return data;
}

function setValueAtPath(obj, path, value) {
    const keys = path.split('.');
    let current = obj;
    
    for (let i = 0; i < keys.length - 1; i++) {
        const key = keys[i];
        if (!(key in current)) {
            current[key] = {};
        }
        current = current[key];
    }
    
    current[keys[keys.length - 1]] = value;
}

function getValueAtPath(obj, path) {
    const keys = path.split('.');
    let current = obj;
    
    for (const key of keys) {
        if (current === null || current === undefined) return undefined;
        current = current[key];
    }
    
    return current;
}

// ============================================================================
// Validation
// ============================================================================

async function validateForm() {
    const data = collectFormData();
    
    try {
        const response = await fetch(`/api/validate/${SCHEMA_NAME}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        
        const result = await response.json();
        
        if (result.valid) {
            showSuccess('✓ Formulario válido');
            return true;
        } else {
            showErrors(result.errors);
            return false;
        }
    } catch (e) {
        showErrors([`Error de validación: ${e.message}`]);
        return false;
    }
}

// ============================================================================
// Save & Download
// ============================================================================

async function saveForm() {
    const data = collectFormData();
    
    // Clear previous messages
    errorsDiv.style.display = 'none';
    successDiv.style.display = 'none';
    
    try {
        const response = await fetch(`/api/save/${SCHEMA_NAME}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showSuccess(`✓ Guardado: ${result.filename}`);
            
            // Auto-download
            setTimeout(() => {
                window.location.href = result.download_url;
            }, 500);
            
            // Clear localStorage
            localStorage.removeItem(`form_${SCHEMA_NAME}`);
        } else {
            showErrors(result.errors);
        }
    } catch (e) {
        showErrors([`Error al guardar: ${e.message}`]);
    }
}

function resetForm() {
    if (confirm('¿Seguro que quieres resetear el formulario?')) {
        formData = {};
        localStorage.removeItem(`form_${SCHEMA_NAME}`);
        renderForm();
        successDiv.style.display = 'none';
        errorsDiv.style.display = 'none';
    }
}

// ============================================================================
// Storage
// ============================================================================

function saveToStorage() {
    const data = collectFormData();
    localStorage.setItem(`form_${SCHEMA_NAME}`, JSON.stringify(data));
}

function loadFromStorage() {
    const saved = localStorage.getItem(`form_${SCHEMA_NAME}`);
    if (!saved) return;
    
    try {
        const data = JSON.parse(saved);
        
        // Populate form
        for (const [path, value] of Object.entries(flattenObject(data))) {
            const input = document.getElementById(path);
            if (!input) continue;
            
            if (input.type === 'checkbox') {
                input.checked = value;
            } else {
                input.value = value;
            }
        }
    } catch (e) {
        console.error('Error loading from storage:', e);
    }
}

function flattenObject(obj, prefix = '') {
    const result = {};
    
    for (const [key, value] of Object.entries(obj)) {
        const path = prefix ? `${prefix}.${key}` : key;
        
        if (value && typeof value === 'object' && !Array.isArray(value)) {
            Object.assign(result, flattenObject(value, path));
        } else {
            result[path] = value;
        }
    }
    
    return result;
}

// ============================================================================
// UI Helpers
// ============================================================================

function showErrors(errors) {
    errorsDiv.innerHTML = errors.map(e => `<div>• ${e}</div>`).join('');
    errorsDiv.style.display = 'block';
    successDiv.style.display = 'none';
}

function showSuccess(message) {
    successDiv.innerHTML = message;
    successDiv.style.display = 'block';
    errorsDiv.style.display = 'none';
}

function focusNextField(current) {
    const inputs = Array.from(formContainer.querySelectorAll('input, select, textarea'));
    const currentIndex = inputs.indexOf(current);
    const next = inputs[currentIndex + 1];
    if (next) next.focus();
}

function toggleHelp() {
    helpPanel.style.display = helpPanel.style.display === 'none' ? 'block' : 'none';
}

// ============================================================================
// Event Bindings
// ============================================================================

btnSave.addEventListener('click', saveForm);
btnValidate.addEventListener('click', validateForm);
btnReset.addEventListener('click', resetForm);
btnHelp.addEventListener('click', toggleHelp);

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl+S → Save
    if (e.ctrlKey && e.key === 's') {
        e.preventDefault();
        saveForm();
    }
    
    // Ctrl+R → Reset (prevent default browser reload)
    if (e.ctrlKey && e.key === 'r') {
        e.preventDefault();
        resetForm();
    }
    
    // Escape → Close help
    if (e.key === 'Escape') {
        helpPanel.style.display = 'none';
    }
    
    // ? → Toggle help
    if (e.key === '?' && !e.ctrlKey && !e.altKey) {
        e.preventDefault();
        toggleHelp();
    }
});

// ============================================================================
// Initialize
// ============================================================================

renderForm();
console.log('✓ Form generator ready');