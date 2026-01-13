// Дані конвертації - ОНОВЛЕНІ
const CONVERSION_TABLES = {
    "Українська мова": {
        7: 100, 8: 105, 9: 110, 10: 115, 11: 120, 12: 125,  
        13: 131, 14: 134, 15: 136, 16: 138, 17: 140, 18: 142,  
        19: 143, 20: 144, 21: 145, 22: 146, 23: 148, 24: 149,  
        25: 150, 26: 152, 27: 154, 28: 156, 29: 157, 30: 159,  
        31: 160, 32: 162, 33: 163, 34: 165, 35: 167, 36: 170,  
        37: 172, 38: 175, 39: 177, 40: 180, 41: 183, 42: 186,  
        43: 191, 44: 195, 45: 200
    },
    "Математика": {
        5: 100, 6: 108, 7: 115, 8: 123, 9: 131, 10: 134,  
        11: 137, 12: 140, 13: 143, 14: 145, 15: 147, 16: 148,  
        17: 149, 18: 150, 19: 151, 20: 152, 21: 155, 22: 159,  
        23: 163, 24: 167, 25: 170, 26: 173, 27: 176, 28: 180,  
        29: 184, 30: 189, 31: 194, 32: 200
    },
    "Історія України": {
        8: 100, 9: 105, 10: 111, 11: 116, 12: 120, 13: 124,  
        14: 127, 15: 130, 16: 132, 17: 134, 18: 136, 19: 138,  
        20: 140, 21: 141, 22: 142, 23: 143, 24: 144, 25: 145,  
        26: 146, 27: 147, 28: 148, 29: 149, 30: 150, 31: 151,  
        32: 152, 33: 154, 34: 156, 35: 158, 36: 160, 37: 163,  
        38: 166, 39: 168, 40: 169, 41: 170, 42: 172, 43: 173,  
        44: 175, 45: 177, 46: 179, 47: 181, 48: 183, 49: 185,  
        50: 188, 51: 191, 52: 194, 53: 197, 54: 200
    },
    "Англійська мова": {
        5: 100, 6: 109, 7: 118, 8: 125, 9: 131, 10: 134,  
        11: 137, 12: 140, 13: 143, 14: 145, 15: 147, 16: 148,  
        17: 149, 18: 150, 19: 151, 20: 152, 21: 153, 22: 155,  
        23: 157, 24: 159, 25: 162, 26: 166, 27: 169, 28: 173,  
        29: 179, 30: 185, 31: 191, 32: 200
    }
};

// Іконки для предметів
const SUBJECT_ICONS = {
    "Українська мова": "fas fa-language",
    "Математика": "fas fa-calculator",
    "Історія України": "fas fa-landmark",
    "Англійська мова": "fas fa-globe-europe"
};

// Глобальні змінні
let selectedSubject = "Українська мова";
let fileData = null;
let convertedData = null;
let workbook = null;
let headers = [];

// Ініціалізація при завантаженні сторінки
document.addEventListener('DOMContentLoaded', function() {
    console.log("Сторінка завантажена, ініціалізація...");
    initializeSubjects();
    setupEventListeners();
    updateConversionInfo();
    updateStatus("Готово до роботи! Оберіть предмет та завантажте файл.", "success");
});

// Ініціалізація предметів
function initializeSubjects() {
    console.log("Ініціалізація предметів...");
    const container = document.getElementById('subjectGrid');
    container.innerHTML = '';
    
    for (const [subject, icon] of Object.entries(SUBJECT_ICONS)) {
        const item = document.createElement('div');
        item.className = 'subject-item';
        item.innerHTML = `
            <i class="${icon}"></i>
            <div>${subject}</div>
        `;
        
        item.addEventListener('click', () => selectSubject(subject));
        container.appendChild(item);
    }
    
    // Вибираємо перший предмет за замовчуванням
    selectSubject("Українська мова");
}

// Вибір предмету
function selectSubject(subject) {
    console.log("Обрано предмет:", subject);
    
    // Знімаємо виділення з усіх предметів
    document.querySelectorAll('.subject-item').forEach(item => {
        item.classList.remove('active');
    });
    
    // Виділяємо вибраний предмет
    document.querySelectorAll('.subject-item').forEach(item => {
        if (item.textContent.includes(subject)) {
            item.classList.add('active');
        }
    });
    
    selectedSubject = subject;
    updateConversionInfo();
    updateStatus(`Обрано предмет: ${subject}`, "success");
}

// Оновлення інформації про конвертацію
function updateConversionInfo() {
    const container = document.getElementById('conversionInfo');
    const table = CONVERSION_TABLES[selectedSubject];
    
    if (!table) {
        container.innerHTML = '<div style="padding: 20px; text-align: center; color: #666;">Таблиця конвертації не знайдена</div>';
        return;
    }
    
    let html = '';
    const entries = Object.entries(table);
    
    // Показуємо всі значення
    entries.forEach(([testScore, nmtScore]) => {
        html += `
            <div class="conversion-item">
                <div class="score">${testScore}</div>
                <div class="nmt-score">→ ${nmtScore}</div>
            </div>
        `;
    });
    
    container.innerHTML = html;
}

// Налаштування обробників подій
function setupEventListeners() {
    console.log("Налаштування обробників подій...");
    
    // Завантаження файлу
    const fileInput = document.getElementById('fileInput');
    const fileLabel = document.getElementById('fileLabel');
    const uploadArea = document.getElementById('uploadArea');
    
    // Клік на тексті "Оберіть файл"
    if (fileLabel) {
        fileLabel.addEventListener('click', (e) => {
            e.stopPropagation();
            fileInput.click();
        });
    }
    
    // Клік на всій області завантаження
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // Зміна вибору файлу
    fileInput.addEventListener('change', handleFileSelect);
    
    // Перетягування файлів
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // Очищення файлу
    document.getElementById('clearFile').addEventListener('click', clearFile);
    
    // Конвертація
    document.getElementById('convertBtn').addEventListener('click', performConversion);
    
    console.log("Обробники подій налаштовані");
}

// Обробка перетягування файлів
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    const uploadArea = document.getElementById('uploadArea');
    uploadArea.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    const uploadArea = document.getElementById('uploadArea');
    uploadArea.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    const uploadArea = document.getElementById('uploadArea');
    uploadArea.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

// Обробка вибору файлу
function handleFileSelect(e) {
    console.log("Файл вибрано:", e.target.files[0]);
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
}

// Обробка файлу
async function handleFile(file) {
    if (!file) {
        updateStatus("Файл не вибрано", "error");
        return;
    }
    
    // Перевірка розширення
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xls') && !fileName.endsWith('.xlsx')) {
        updateStatus("Помилка: Файл має бути у форматі Excel (.xls або .xlsx)", "error");
        return;
    }
    
    updateStatus("Завантаження файлу...", "loading");
    
    try {
        const data = await readFileAsArrayBuffer(file);
        workbook = XLSX.read(data, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error("Файл не містить жодного аркуша");
        }
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (jsonData.length === 0) {
            throw new Error("Аркуш порожній");
        }
        
        fileData = jsonData;
        showFilePreview(jsonData, file.name);
        setupColumnSelectors(jsonData[0]);
        showColumnSettings();
        
        updateStatus(`Файл "${file.name}" успішно завантажено (${jsonData.length - 1} рядків даних)`, "success");
        
    } catch (error) {
        console.error("Помилка при читанні файлу:", error);
        updateStatus(`Помилка: ${error.message}`, "error");
        clearFile();
    }
}

// Читання файлу як ArrayBuffer
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(new Error("Помилка читання файлу"));
        reader.readAsArrayBuffer(file);
    });
}

// Показ попереднього перегляду файлу - ВИПРАВЛЕНА ВЕРСІЯ
function showFilePreview(data, fileName) {
    const preview = document.getElementById('filePreview');
    const uploadArea = document.getElementById('uploadArea');
    
    // Показуємо прев'ю, ховаємо область завантаження
    uploadArea.style.display = 'none';
    preview.style.display = 'block';
    preview.classList.add('fade-in');
    
    // Заповнення заголовків таблиці
    const headerRow = document.getElementById('tableHeader');
    headerRow.innerHTML = '';
    
    if (data.length > 0) {
        headers = data[0];
        
        // Обмежуємо кількість стовпців для відображення
        const maxColumns = 6; // Максимум 6 стовпців для прев'ю
        const columnsToShow = Math.min(headers.length, maxColumns);
        
        for (let i = 0; i < columnsToShow; i++) {
            const th = document.createElement('th');
            const headerText = headers[i] || `Стовпець ${i + 1}`;
            
            // Обрізаємо довгий текст
            th.textContent = headerText.length > 15 ? 
                headerText.substring(0, 12) + '...' : headerText;
            th.title = headerText; // Повний текст у підказці
            headerRow.appendChild(th);
        }
        
        // Якщо стовпців більше ніж maxColumns, показуємо індикацію
        if (headers.length > maxColumns) {
            const th = document.createElement('th');
            th.textContent = `...`;
            th.title = `Ще ${headers.length - maxColumns} стовпців`;
            th.style.color = '#666';
            th.style.fontStyle = 'italic';
            th.style.minWidth = '50px';
            headerRow.appendChild(th);
        }
    }
    
    // Заповнення тіла таблиці
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';
    
    // Показуємо максимум 8 рядків для попереднього перегляду
    const rowsToShow = Math.min(8, data.length - 1);
    const maxColumns = Math.min(headers.length, 6);
    
    for (let i = 1; i <= rowsToShow; i++) {
        const row = data[i];
        const tr = document.createElement('tr');
        
        // Додаємо дані для видимих стовпців
        for (let j = 0; j < maxColumns; j++) {
            const td = document.createElement('td');
            let cellValue = row && row[j] !== undefined ? String(row[j]) : '';
            
            // Обрізаємо довгий текст
            if (cellValue.length > 20) {
                cellValue = cellValue.substring(0, 17) + '...';
            }
            
            td.textContent = cellValue;
            td.title = row && row[j] !== undefined ? String(row[j]) : '';
            tr.appendChild(td);
        }
        
        // Додаємо індикацію, якщо стовпців більше
        if (headers.length > maxColumns) {
            const td = document.createElement('td');
            td.textContent = '...';
            td.style.color = '#666';
            td.style.fontStyle = 'italic';
            td.style.minWidth = '50px';
            tr.appendChild(td);
        }
        
        tableBody.appendChild(tr);
    }
    
    // Статистика файлу
    const stats = document.getElementById('fileStats');
    const totalRows = data.length - 1;
    stats.innerHTML = `
        <div style="display: flex; flex-wrap: wrap; gap: 10px; justify-content: space-between;">
            <div><strong>Файл:</strong> ${fileName}</div>
            <div><strong>Рядків:</strong> ${totalRows}</div>
            <div><strong>Стовпців:</strong> ${headers.length}</div>
        </div>
    `;
}

// Налаштування вибору стовпців
function setupColumnSelectors(headersArray) {
    const studentSelect = document.getElementById('studentColumn');
    const scoreSelect = document.getElementById('scoreColumn');
    
    studentSelect.innerHTML = '';
    scoreSelect.innerHTML = '';
    
    headersArray.forEach((header, index) => {
        const studentOption = document.createElement('option');
        studentOption.value = index;
        const headerText = header || `Стовпець ${index + 1}`;
        studentOption.textContent = headerText.length > 30 ? 
            headerText.substring(0, 27) + '...' : headerText;
        studentOption.title = headerText;
        
        const scoreOption = studentOption.cloneNode(true);
        
        studentSelect.appendChild(studentOption);
        scoreSelect.appendChild(scoreOption);
        
        // Автовибір за ключовими словами
        const headerLower = String(header).toLowerCase();
        if (studentSelect.value === '' && 
            (headerLower.includes('піб') || headerLower.includes('ім') || 
             headerLower.includes('учн') || headerLower.includes('студ') || 
             headerLower.includes('name') || headerLower.includes('фіо') ||
             headerLower.includes('повне'))) {
            studentSelect.value = index;
        }
        if (scoreSelect.value === '' && 
            (headerLower.includes('бал') || headerLower.includes('оцін') || 
             headerLower.includes('резул') || headerLower.includes('score') || 
             headerLower.includes('mark') || headerLower.includes('points') ||
             headerLower.includes('оценка'))) {
            scoreSelect.value = index;
        }
    });
    
    // Якщо не знайдено автоматично, вибираємо перший стовпець
    if (studentSelect.value === '' && headersArray.length > 0) {
        studentSelect.value = 0;
    }
    if (scoreSelect.value === '' && headersArray.length > 1) {
        scoreSelect.value = Math.min(1, headersArray.length - 1);
    }
}

// Показ налаштувань стовпців
function showColumnSettings() {
    const columnSettings = document.getElementById('columnSettings');
    columnSettings.style.display = 'block';
    columnSettings.classList.add('fade-in');
}

// Очищення файлу
function clearFile() {
    fileData = null;
    workbook = null;
    convertedData = null;
    headers = [];
    
    document.getElementById('fileInput').value = '';
    const filePreview = document.getElementById('filePreview');
    filePreview.style.display = 'none';
    filePreview.classList.remove('fade-in');
    
    document.getElementById('uploadArea').style.display = 'flex';
    document.getElementById('columnSettings').style.display = 'none';
    document.getElementById('resultsCard').style.display = 'none';
    
    updateStatus("Файл очищено", "info");
}

// Виконання конвертації
function performConversion() {
    const studentColumn = parseInt(document.getElementById('studentColumn').value);
    const scoreColumn = parseInt(document.getElementById('scoreColumn').value);
    
    if (isNaN(studentColumn) || isNaN(scoreColumn)) {
        updateStatus("Будь ласка, виберіть обидва стовпці", "error");
        return;
    }
    
    if (!fileData || fileData.length === 0) {
        updateStatus("Спочатку завантажте файл", "error");
        return;
    }
    
    updateStatus("Виконується конвертація балів...", "loading");
    
    // Деактивуємо кнопку під час обробки
    const convertBtn = document.getElementById('convertBtn');
    convertBtn.disabled = true;
    convertBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Обробка...';
    
    // Виконуємо асинхронно для уникнення блокування інтерфейсу
    setTimeout(() => {
        try {
            const result = convertScores(studentColumn, scoreColumn);
            showResults(result);
            updateStatus(`Конвертацію завершено! Оброблено ${result.success} з ${result.total} рядків`, "success");
        } catch (error) {
            console.error("Помилка конвертації:", error);
            updateStatus(`Помилка: ${error.message}`, "error");
        } finally {
            // Активуємо кнопку знову
            convertBtn.disabled = false;
            convertBtn.innerHTML = '<i class="fas fa-sync-alt"></i> Конвертувати бали';
        }
    }, 50);
}

// Функція конвертації балів
function convertScores(studentColumnIndex, scoreColumnIndex) {
    if (!selectedSubject) {
        throw new Error("Виберіть предмет для конвертації");
    }
    
    const conversionTable = CONVERSION_TABLES[selectedSubject];
    if (!conversionTable) {
        throw new Error("Таблиця конвертації не знайдена для обраного предмета");
    }
    
    // Отримуємо мінімальне та максимальне значення з таблиці конвертації
    const testScores = Object.keys(conversionTable).map(Number);
    const minScore = Math.min(...testScores);
    const maxScore = Math.max(...testScores);
    
    const results = [];
    let removedRows = 0;
    const totalRows = fileData.length - 1; // Мінус заголовки
    
    // Обробляємо рядки даних (починаючи з 1, бо 0 - заголовки)
    for (let i = 1; i < fileData.length; i++) {
        const row = fileData[i];
        
        if (row && row.length > Math.max(studentColumnIndex, scoreColumnIndex)) {
            const studentName = row[studentColumnIndex] || '';
            const testScore = row[scoreColumnIndex];
            
            try {
                // Конвертуємо бал у число
                const scoreNum = parseFloat(testScore);
                if (isNaN(scoreNum)) {
                    removedRows++;
                    continue;
                }
                
                const scoreInt = Math.round(scoreNum);
                
                // Перевіряємо, чи є бал у таблиці конвертації
                if (conversionTable[scoreInt] !== undefined) {
                    results.push({
                        originalIndex: i,
                        ПІБ: String(studentName).trim(),
                        'Бал НМТ': conversionTable[scoreInt]
                    });
                } else {
                    // Якщо бал поза діапазоном таблиці
                    if (scoreInt < minScore) {
                        // Використовуємо мінімальне значення
                        results.push({
                            originalIndex: i,
                            ПІБ: String(studentName).trim(),
                            'Бал НМТ': conversionTable[minScore]
                        });
                    } else if (scoreInt > maxScore) {
                        // Використовуємо максимальне значення
                        results.push({
                            originalIndex: i,
                            ПІБ: String(studentName).trim(),
                            'Бал НМТ': conversionTable[maxScore]
                        });
                    } else {
                        // Шукаємо найближче значення
                        let closestScore = minScore;
                        let minDiff = Math.abs(scoreInt - minScore);
                        
                        for (const key of testScores) {
                            const diff = Math.abs(scoreInt - key);
                            if (diff < minDiff) {
                                minDiff = diff;
                                closestScore = key;
                            }
                        }
                        
                        results.push({
                            originalIndex: i,
                            ПІБ: String(studentName).trim(),
                            'Бал НМТ': conversionTable[closestScore]
                        });
                    }
                }
            } catch (error) {
                console.warn(`Помилка в рядку ${i}:`, error);
                removedRows++;
            }
        } else {
            removedRows++;
        }
    }
    
    // Сортуємо за оригінальним порядком
    results.sort((a, b) => a.originalIndex - b.originalIndex);
    
    // Видаляємо originalIndex з результатів
    const finalResults = results.map(item => ({
        ПІБ: item.ПІБ,
        'Бал НМТ': item['Бал НМТ']
    }));
    
    convertedData = finalResults;
    
    return {
        success: results.length,
        total: totalRows,
        removed: removedRows,
        data: finalResults
    };
}

// Показ результатів
function showResults(result) {
    const resultsCard = document.getElementById('resultsCard');
    const resultsStats = document.getElementById('resultsStats');
    const resultsPreview = document.getElementById('resultsPreview');
    const downloadButtons = document.getElementById('downloadButtons');
    
    // Статистика
    resultsStats.innerHTML = `
        <div class="stat-item">
            <div class="stat-value">${result.success}</div>
            <div class="stat-label">Успішно</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">${result.total}</div>
            <div class="stat-label">Всього</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">${result.removed}</div>
            <div class="stat-label">Пропущено</div>
        </div>
    `;
    
    // Попередній перегляд результатів
    let previewHTML = '<div class="results-preview">';
    previewHTML += '<table>';
    previewHTML += '<tr><th>ПІБ</th><th>Бал НМТ</th></tr>';
    
    // Показуємо максимум 10 рядків
    const rowsToShow = Math.min(10, result.data.length);
    for (let i = 0; i < rowsToShow; i++) {
        const row = result.data[i];
        const pib = row.ПІБ.length > 30 ? row.ПІБ.substring(0, 27) + '...' : row.ПІБ;
        previewHTML += `<tr><td title="${row.ПІБ}">${pib}</td><td>${row['Бал НМТ']}</td></tr>`;
    }
    
    if (result.data.length > 10) {
        previewHTML += `<tr><td colspan="2" style="text-align: center; color: #666; padding: 10px; font-style: italic;">
            ... та ще ${result.data.length - 10} рядків
        </td></tr>`;
    }
    
    previewHTML += '</table></div>';
    resultsPreview.innerHTML = previewHTML;
    
    // Кнопки завантаження
    downloadButtons.innerHTML = `
        <button class="download-btn excel" onclick="downloadExcel()">
            <i class="fas fa-file-excel"></i> Завантажити Excel
        </button>
        <button class="download-btn" onclick="downloadCSV()">
            <i class="fas fa-file-csv"></i> Завантажити CSV
        </button>
    `;
    
    resultsCard.style.display = 'block';
    resultsCard.classList.add('fade-in');
}

// Завантаження Excel
function downloadExcel() {
    if (!convertedData || convertedData.length === 0) {
        updateStatus("Немає даних для експорту", "error");
        return;
    }
    
    try {
        // Створюємо новий workbook
        const ws = XLSX.utils.json_to_sheet(convertedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Результати НМТ');
        
        // Налаштовуємо ширину стовпців
        if (!ws['!cols']) ws['!cols'] = [];
        ws['!cols'][0] = { wch: 40 }; // Ширина для ПІБ
        ws['!cols'][1] = { wch: 15 }; // Ширина для балів
        
        // Генеруємо назву файлу
        const subjectName = selectedSubject.replace(/\s+/g, '_').replace(/[^\wа-яА-ЯіІїЇєЄґҐ_]/g, '');
        const fileName = `результати_нмт_${subjectName}.xlsx`;
        
        // Записуємо файл
        XLSX.writeFile(wb, fileName);
        
        updateStatus(`Файл "${fileName}" завантажено успішно`, "success");
    } catch (error) {
        console.error("Помилка при створенні Excel:", error);
        updateStatus(`Помилка: ${error.message}`, "error");
    }
}

// Завантаження CSV
function downloadCSV() {
    if (!convertedData || convertedData.length === 0) {
        updateStatus("Немає даних для експорту", "error");
        return;
    }
    
    try {
        // Конвертуємо в CSV з роздільником крапка з комою
        const headers = ['ПІБ', 'Бал НМТ'];
        const csvRows = [
            headers.join(';'),
            ...convertedData.map(row => `${row.ПІБ};${row['Бал НМТ']}`)
        ];
        
        const csvContent = '\uFEFF' + csvRows.join('\n'); // BOM для UTF-8
        
        // Створюємо Blob
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        
        // Генеруємо назву файлу
        const subjectName = selectedSubject.replace(/\s+/g, '_').replace(/[^\wа-яА-ЯіІїЇєЄґҐ_]/g, '');
        const fileName = `результати_нмт_${subjectName}.csv`;
        
        // Створюємо URL для завантаження
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', fileName);
        link.style.visibility = 'hidden';
        
        // Додаємо посилання на сторінку та клікаємо
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Звільняємо пам'ять
        setTimeout(() => URL.revokeObjectURL(url), 100);
        
        updateStatus(`Файл "${fileName}" завантажено успішно`, "success");
    } catch (error) {
        console.error("Помилка при створенні CSV:", error);
        updateStatus(`Помилка: ${error.message}`, "error");
    }
}

// Оновлення статусу
function updateStatus(message, type = "info") {
    const statusBar = document.getElementById('statusBar');
    const statusMessage = document.getElementById('statusMessage');
    
    if (!statusBar || !statusMessage) {
        console.log(`Статус [${type}]: ${message}`);
        return;
    }
    
    // Очищаємо класи
    statusBar.className = 'status-bar';
    statusMessage.className = 'status-message';
    
    // Додаємо клас за типом
    statusBar.classList.add(`status-${type}`);
    
    let icon = 'fas fa-info-circle';
    if (type === 'success') icon = 'fas fa-check-circle';
    if (type === 'error') icon = 'fas fa-exclamation-circle';
    if (type === 'loading') icon = 'fas fa-spinner fa-spin';
    
    statusMessage.innerHTML = `
        <i class="${icon} status-icon"></i>
        <span>${message}</span>
    `;
    
    // Автоматичне очищення повідомлень про успіх через 5 секунд
    if (type === 'success') {
        setTimeout(() => {
            if (statusMessage.innerHTML.includes(message)) {
                statusMessage.innerHTML = '';
                statusBar.className = 'status-bar';
            }
        }, 5000);
    }
}

// Експорт функцій для глобального використання
window.selectSubject = selectSubject;
window.clearFile = clearFile;
window.performConversion = performConversion;
window.downloadExcel = downloadExcel;
window.downloadCSV = downloadCSV;