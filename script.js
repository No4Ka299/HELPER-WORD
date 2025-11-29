// Глобальные переменные
let uploadedText = '';

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
});

// Инициализация всех обработчиков событий
function initializeEventListeners() {
    // Загрузка файла
    const fileInput = document.getElementById('fileInput');
    const uploadLabel = document.querySelector('.upload-label');
    
    fileInput.addEventListener('change', handleFileUpload);
    
    // Перетаскивание файлов
    uploadLabel.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('drag-over');
    });
    
    uploadLabel.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.classList.remove('drag-over');
    });
    
    uploadLabel.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileUpload({ target: { files: e.dataTransfer.files } });
        }
    });
    
    // Изменение текста в textarea
    document.getElementById('textInput').addEventListener('input', function() {
        uploadedText = this.value;
        updatePreview();
    });
    
    // Кнопка форматирования и скачивания
    document.getElementById('formatBtn').addEventListener('click', formatAndDownload);
    
    // Кнопка предварительного просмотра
    document.getElementById('previewBtn').addEventListener('click', updatePreview);
}

// Обработка загрузки файла
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        let content = e.target.result;
        
        // Определяем тип файла по расширению
        const fileName = file.name.toLowerCase();
        
        if (fileName.endsWith('.txt')) {
            uploadedText = content;
        } else if (fileName.endsWith('.docx')) {
            // Для .docx файлов нужно использовать специальную библиотеку
            // В реальном приложении использовалась бы библиотека для чтения docx
            // Пока что покажем сообщение
            alert('Формат .docx требует дополнительной обработки. В демонстрации используется текст.');
            uploadedText = content.substring(0, 1000) + '... [текст обрезан для демонстрации]';
        } else if (fileName.endsWith('.doc')) {
            // Для .doc файлов тоже нужна специальная обработка
            alert('Формат .doc требует дополнительной обработки. В демонстрации используется текст.');
            uploadedText = content.substring(0, 1000) + '... [текст обрезан для демонстрации]';
        } else {
            uploadedText = content;
        }
        
        // Обновляем textarea
        document.getElementById('textInput').value = uploadedText;
        
        // Обновляем превью
        updatePreview();
    };
    
    // Читаем файл в зависимости от типа
    if (file.type === 'text/plain' || file.name.toLowerCase().endsWith('.txt')) {
        reader.readAsText(file, 'UTF-8');
    } else {
        reader.readAsText(file, 'UTF-8'); // В демонстрации читаем как текст
    }
}

// Обновление предварительного просмотра
function updatePreview() {
    const text = uploadedText || document.getElementById('textInput').value;
    const previewContent = document.getElementById('previewContent');
    
    if (text) {
        // Применяем форматирование для предварительного просмотра
        const formattedText = formatTextForPreview(text);
        previewContent.textContent = formattedText;
    } else {
        previewContent.textContent = 'Текст для предварительного просмотра появится здесь после загрузки...';
    }
}

// Форматирование текста для предварительного просмотра
function formatTextForPreview(text) {
    // Простое форматирование для отображения в превью
    // В реальном приложении тут будет более сложная логика
    return text.substring(0, 1000) + (text.length > 1000 ? '... [далее текст сокращен для просмотра]' : '');
}

// Форматирование и скачивание DOCX
async function formatAndDownload() {
    const text = uploadedText || document.getElementById('textInput').value;
    if (!text.trim()) {
        alert('Пожалуйста, загрузите или введите текст для форматирования.');
        return;
    }
    
    // Показываем индикатор загрузки
    const btn = document.getElementById('formatBtn');
    const originalText = btn.textContent;
    btn.innerHTML = '<span class="loading"></span> Обработка...';
    btn.disabled = true;
    
    try {
        // Получаем настройки форматирования
        const settings = getFormattingSettings();
        
        // Создаем документ с помощью библиотеки docx
        const doc = createFormattedDocument(text, settings);
        
        // Генерируем файл
        const blob = await generateDocxBlob(doc);
        
        // Скачиваем файл
        downloadFile(blob, 'formatted_coursework.docx');
        
        // Восстанавливаем кнопку
        btn.textContent = originalText;
        btn.disabled = false;
        
        alert('Документ успешно сформирован и скачивается!');
    } catch (error) {
        console.error('Ошибка при создании документа:', error);
        alert('Произошла ошибка при создании документа. Пожалуйста, попробуйте снова.');
        
        // Восстанавливаем кнопку
        btn.textContent = originalText;
        btn.disabled = false;
    }
}

// Получение настроек форматирования
function getFormattingSettings() {
    return {
        fontFamily: document.getElementById('fontFamily').value,
        fontSize: parseInt(document.getElementById('fontSize').value),
        lineSpacing: parseFloat(document.getElementById('lineSpacing').value),
        indentSize: parseFloat(document.getElementById('indentSize').value),
        pageMargins: document.getElementById('pageMargin').value,
        autoParagraph: document.getElementById('autoParagraph').checked
    };
}

// Создание форматированного документа
function createFormattedDocument(text, settings) {
    // Используем библиотеку docx для создания документа
    const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak, SectionType } = docx;
    
    // Разбиваем текст на абзацы
    const paragraphs = text.split(/\n\s*\n/).filter(p => p.trim() !== '');
    
    // Создаем массив параграфов документа
    const docParagraphs = paragraphs.map(paragraphText => {
        // Удаляем лишние переносы строк внутри абзаца и форматируем
        const cleanText = paragraphText.replace(/\n/g, ' ').trim();
        
        // Создаем параграф с настройками форматирования
        return new Paragraph({
            text: cleanText,
            spacing: {
                line: settings.lineSpacing * 240, // docx использует специальную единицу измерения
            },
            indent: {
                firstLine: settings.indentSize * 567, // 1 см ≈ 567 единиц в docx
            },
            heading: HeadingLevel.HEADING_1, //временно отключим заголовки
        });
    });
    
    // Определяем отступы страницы
    const margins = parsePageMargins(settings.pageMargins);
    
    // Создаем документ
    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: margins.top * 15, // docx использует специальные единицы
                        right: margins.right * 15,
                        bottom: margins.bottom * 15,
                        left: margins.left * 15,
                    }
                },
                type: SectionType.CONTINUOUS
            },
            children: docParagraphs,
        }],
        styles: {
            default: {
                heading1: {
                    run: {
                        size: settings.fontSize * 2,
                        font: settings.fontFamily,
                        bold: true,
                    },
                    paragraph: {
                        spacing: {
                            line: settings.lineSpacing * 240,
                        },
                    },
                },
                heading2: {
                    run: {
                        size: (settings.fontSize - 2) * 2,
                        font: settings.fontFamily,
                        bold: true,
                    },
                    paragraph: {
                        spacing: {
                            line: settings.lineSpacing * 240,
                        },
                    },
                },
                paragraph: {
                    run: {
                        size: settings.fontSize * 2,
                        font: settings.fontFamily,
                    },
                    paragraph: {
                        spacing: {
                            line: settings.lineSpacing * 240,
                        },
                        indent: {
                            firstLine: settings.indentSize * 567,
                        },
                    },
                },
            },
        },
    });
    
    return doc;
}

// Парсинг строки отступов страницы
function parsePageMargins(marginString) {
    // Пример: "3, 2, 2, 2" -> левый, правый, верхний, нижний
    const values = marginString.split(',').map(v => parseFloat(v.trim()) || 2);
    return {
        left: values[0] || 2,
        right: values[1] || 2,
        top: values[2] || 2,
        bottom: values[3] || 2
    };
}

// Генерация Blob из документа
function generateDocxBlob(document) {
    return docx.Packer.toBlob(document);
}

// Скачивание файла
function downloadFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Дополнительные утилиты для форматирования текста

// Функция для автоматического форматирования абзацев
function autoFormatParagraphs(text) {
    // Разбиваем текст на строки
    const lines = text.split('\n');
    const formattedLines = [];
    let currentParagraph = '';
    
    for (const line of lines) {
        const trimmedLine = line.trim();
        
        if (trimmedLine === '') {
            // Пустая строка - завершаем текущий абзац
            if (currentParagraph) {
                formattedLines.push(currentParagraph);
                currentParagraph = '';
            }
        } else {
            // Добавляем строку к текущему абзацу
            if (currentParagraph) {
                currentParagraph += ' ' + trimmedLine;
            } else {
                currentParagraph = trimmedLine;
            }
        }
    }
    
    // Добавляем последний абзац, если он есть
    if (currentParagraph) {
        formattedLines.push(currentParagraph);
    }
    
    return formattedLines.join('\n\n');
}