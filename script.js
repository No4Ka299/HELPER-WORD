// Глобальные переменные
let originalText = '';
let formattedText = '';

// Элементы DOM
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const uploadArea = document.getElementById('uploadArea');
const textInput = document.getElementById('textInput');
const previewContent = document.getElementById('previewContent');
const formatBtn = document.getElementById('formatBtn');
const exportBtn = document.getElementById('exportBtn');

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
});

// Установка обработчиков событий
function setupEventListeners() {
    // Обработчики для загрузки файла
    browseBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileUpload);
    
    // Обработчики для области перетаскивания
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.backgroundColor = '#e3f2fd';
        uploadArea.style.borderColor = '#2980b9';
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.style.backgroundColor = '#f8f9fa';
        uploadArea.style.borderColor = '#3498db';
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.backgroundColor = '#f8f9fa';
        uploadArea.style.borderColor = '#3498db';
        
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
    
    // Обработчик изменения текста
    textInput.addEventListener('input', updatePreview);
    
    // Обработчики кнопок
    formatBtn.addEventListener('click', applyFormatting);
    exportBtn.addEventListener('click', exportToDocx);
}

// Обработка загрузки файла
function handleFileUpload(e) {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
}

// Обработка файла (любого типа)
async function handleFile(file) {
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop().toLowerCase();
    
    try {
        if (fileExtension === 'txt') {
            // Для текстовых файлов
            const text = await file.text();
            textInput.value = text;
            originalText = text;
            updatePreview();
        } else if (fileExtension === 'docx') {
            // Для DOCX файлов используем библиотеку mammoth
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.convertToHtml({arrayBuffer: arrayBuffer});
            textInput.value = result.value;
            originalText = result.value;
            updatePreview();
        } else if (fileExtension === 'doc') {
            // Для DOC файлов показываем сообщение
            alert('Формат .doc не поддерживается. Пожалуйста, конвертируйте файл в .docx или .txt');
        } else {
            // Для других файлов пробуем прочитать как текст
            const text = await file.text();
            textInput.value = text;
            originalText = text;
            updatePreview();
        }
    } catch (error) {
        console.error('Ошибка при обработке файла:', error);
        alert('Ошибка при чтении файла. Пожалуйста, проверьте формат файла.');
    }
}

// Обновление превью
function updatePreview() {
    const text = textInput.value;
    previewContent.textContent = text || 'Текст для предварительного просмотра появится здесь...';
}

// Применение форматирования
function applyFormatting() {
    const text = textInput.value;
    if (!text) {
        alert('Пожалуйста, введите текст для форматирования');
        return;
    }
    
    originalText = text;
    
    // Получение параметров форматирования
    const settings = getFormattingSettings();
    
    // Применение форматирования
    formattedText = formatText(text, settings);
    
    // Обновление превью с форматированным текстом
    updateFormattedPreview(formattedText, settings);
    
    alert('Форматирование применено!');
}

// Получение параметров форматирования
function getFormattingSettings() {
    return {
        fontFamily: document.getElementById('fontFamily').value,
        fontSize: parseFloat(document.getElementById('fontSize').value),
        lineSpacing: parseFloat(document.getElementById('lineSpacing').value),
        indentSize: parseFloat(document.getElementById('indentSize').value),
        marginTop: parseFloat(document.getElementById('marginTop').value),
        marginBottom: parseFloat(document.getElementById('marginBottom').value),
        marginLeft: parseFloat(document.getElementById('marginLeft').value),
        marginRight: parseFloat(document.getElementById('marginRight').value),
        paragraphSpacing: parseFloat(document.getElementById('paragraphSpacing').value)
    };
}

// Форматирование текста
function formatText(text, settings) {
    // Разделение текста на абзацы
    const paragraphs = text.split(/\n\s*\n/);
    
    // Форматирование каждого абзаца
    const formattedParagraphs = paragraphs.map(paragraph => {
        // Удаление лишних пробелов и переносов строк
        const cleanParagraph = paragraph.replace(/\n/g, ' ').trim();
        
        if (cleanParagraph) {
            // Добавление отступа для абзаца (кроме заголовков)
            // Проверяем, является ли абзац заголовком (по наличию двоеточия в начале или по другим признакам)
            if (!isHeading(cleanParagraph)) {
                return `<p style="text-indent: ${settings.indentSize}cm; margin-bottom: ${settings.paragraphSpacing}cm; line-height: ${settings.lineSpacing}; font-family: '${settings.fontFamily}', serif; font-size: ${settings.fontSize}pt;">${cleanParagraph}</p>`;
            } else {
                return `<p style="margin-bottom: ${settings.paragraphSpacing}cm; line-height: ${settings.lineSpacing}; font-family: '${settings.fontFamily}', serif; font-size: ${settings.fontSize}pt; font-weight: bold;">${cleanParagraph}</p>`;
            }
        }
        return '';
    }).filter(p => p !== '');
    
    return formattedParagraphs.join('\n');
}

// Проверка, является ли текст заголовком
function isHeading(text) {
    // Простая проверка: если текст короткий и заканчивается на двоеточие или является заголовком по стилю
    const trimmed = text.trim();
    return trimmed.length < 100 && (trimmed.endsWith(':') || /^[А-Я]/.test(trimmed));
}

// Обновление превью с форматированным текстом
function updateFormattedPreview(formattedText, settings) {
    previewContent.innerHTML = `
        <div class="formatted-preview" style="
            font-family: '${settings.fontFamily}', serif; 
            font-size: ${settings.fontSize}pt; 
            line-height: ${settings.lineSpacing};
            margin-top: ${settings.marginTop}cm;
            margin-bottom: ${settings.marginBottom}cm;
            margin-left: ${settings.marginLeft}cm;
            margin-right: ${settings.marginRight}cm;
        ">
            ${formattedText}
        </div>
    `;
}

// Экспорт в DOCX
async function exportToDocx() {
    if (!formattedText && !originalText) {
        alert('Сначала примените форматирование к тексту');
        return;
    }
    
    // Получение параметров форматирования
    const settings = getFormattingSettings();
    
    try {
        // Создание документа с помощью библиотеки docx
        const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak, SectionType, ShadingType } = docx;
        
        // Разделение текста на абзацы
        const paragraphs = (formattedText || originalText).split(/\n\s*\n/);
        
        const docParagraphs = paragraphs.map(paragraph => {
            const cleanParagraph = paragraph.replace(/\n/g, ' ').trim();
            
            if (cleanParagraph) {
                // Проверяем, является ли абзац заголовком
                if (isHeading(cleanParagraph)) {
                    return new Paragraph({
                        text: cleanParagraph,
                        heading: HeadingLevel.HEADING_1,
                        spacing: {
                            after: Math.round(settings.paragraphSpacing * 20 * 72/2.54) || 0, // преобразование см в twips
                        },
                        alignment: AlignmentType.LEFT,
                    });
                } else {
                    return new Paragraph({
                        text: cleanParagraph,
                        spacing: {
                            after: Math.round(settings.paragraphSpacing * 20 * 72/2.54) || 0, // преобразование см в twips
                        },
                        indent: {
                            firstLine: Math.round(settings.indentSize * 20 * 72/2.54) || 0, // преобразование см в twips
                        },
                        spacing: {
                            line: settings.lineSpacing * 240, // множитель для междустрочного интервала
                        },
                        alignment: AlignmentType.JUSTIFIED,
                    });
                }
            }
            return new Paragraph(''); // Пустой абзац для разделения
        });
        
        // Создание документа
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: Math.round(settings.marginTop * 20 * 72/2.54), // преобразование см в twips
                            right: Math.round(settings.marginRight * 20 * 72/2.54),
                            bottom: Math.round(settings.marginBottom * 20 * 72/2.54),
                            left: Math.round(settings.marginLeft * 20 * 72/2.54),
                        },
                    },
                },
                children: docParagraphs,
            }],
        });
        
        // Генерация и скачивание файла
        const { Blob } = docx;
        const packer = new docx.Packer();
        const blob = await packer.toBlob(doc);
        
        // Создание ссылки для скачивания
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'отформатированная_курсовая.docx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        alert('Файл успешно экспортирован!');
    } catch (error) {
        console.error('Ошибка при экспорте:', error);
        alert('Ошибка при экспорте файла. Пожалуйста, попробуйте еще раз.');
    }
}

// Функция для тестирования форматирования (временно)
function testFormatting() {
    const sampleText = `ВВЕДЕНИЕ

Актуальность темы исследования. В современном мире информационные технологии играют все более важную роль. 

Цель работы. Целью данной курсовой работы является изучение влияния информационных технологий на развитие современного общества.

Задачи работы: 
1. Изучить теоретические аспекты информационных технологий.
2. Проанализировать современные тенденции в области ИТ.
3. Рассмотреть практические примеры применения ИТ.

Структура работы. Курсовая работа состоит из введения, двух глав, заключения и списка использованных источников.`;
    
    textInput.value = sampleText;
    originalText = sampleText;
    updatePreview();
}