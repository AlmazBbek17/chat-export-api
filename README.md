Gemini Chat Export API (Python)
API для экспорта чатов Gemini в Word (.docx) с поддержкой таблиц, кода и LaTeX формул.
Установка на Vercel

Создайте репозиторий на GitHub
Загрузите все файлы из этого проекта
Зайдите на vercel.com
Нажмите "Add New" → "Project"
Выберите ваш репозиторий
Нажмите "Deploy"

После деплоя получите URL типа: https://your-project.vercel.app
Использование в расширении
javascriptasync function exportGeminiChatToWord(messages) {
  try {
    const response = await fetch('https://your-project.vercel.app/api/export-chat', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        messages: messages, // массив объектов { role: 'user' | 'model', content: 'текст' }
        title: 'Чат с Gemini ' + new Date().toLocaleDateString('ru-RU')
      })
    });

    if (!response.ok) {
      throw new Error('Ошибка экспорта');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'gemini-chat.docx';
    a.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error('Ошибка:', error);
    alert('Не удалось экспортировать чат');
  }
}
Формат данных
Отправляйте POST запрос на /api/export-chat с телом:
json{
  "messages": [
    {
      "role": "user",
      "content": "Привет!"
    },
    {
      "role": "model",
      "content": "Здравствуйте! Чем могу помочь?"
    }
  ],
  "title": "Чат с Gemini" // опционально
}
Структура проекта
chat-export-api/
├── api/
│   └── export-chat.js    # Serverless функция
├── package.json          # Зависимости
├── .gitignore           # Игнорируемые файлы
└── README.md            # Эта инструкция