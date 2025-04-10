# AprilFrontier
Этот макрос для Microsoft Word автоматически проверяет соответствие научной статьи требованиям оформления (в формате конференций или сборников трудов)
под апрельскую конференцию на основании [информационного письма](https://docs.google.com/document/d/1W2YMnxND9VNnD3Hr3fKju6pvtajEgsQh/edit).

---

## ✅ Возможности макроса

`Sub ПроверкаСоответствияТребованиям`

Проверяет следующее:

- 📏 **Шрифт**: Times New Roman, размер 12 pt
- 📐 **Межстрочный интервал**: одинарный
- 🧾 **Формат страницы**:
  - Размер A4
  - Поля: верх/низ — 15 мм, лево/право — 20 мм
  - Ориентация — книжная
- 🧱 **Абзацный отступ**: 0,75 см в стиле "Обычный"
- 🧹 **Оформление документа**:
  - Отсутствие колонтитулов
  - Отсутствие нумерации страниц
  - Отсутствие сносок

По результатам макрос выводит окно со списком найденных нарушений или сообщает об успешной проверке.

[📄 Информационное письмо (требования к оформлению)](https://docs.google.com/document/d/1W2YMnxND9VNnD3Hr3fKju6pvtajEgsQh/edit)

---

## 🚀 Установка и запуск

1. Открой Microsoft Word.
2. Нажми `Alt + F11` — откроется редактор VBA.
3. Вставь код макроса в модуль.
4. Закрой редактор.
5. Нажми `Alt + F8`, выбери `ПроверкаСоответствияТребованиям` и запусти.

---

## 🤝 Контрибьюция

Хочешь улучшить макрос?

- Добавить проверки структуры (УДК, аннотация, ключевые слова, научный руководитель)
- Проверку названия таблиц и рисунков
- Генерацию отчёта в отдельный файл

Будем рады твоему pull request! Делай форк, добавляй улучшения, открывай обсуждения.

### Как сделать Pull Request (PR)

1. Форкните репозиторий (кнопка "Fork" в правом верхнем углу на GitHub).

2. Клонируйте форк к себе:
   ```bash
   git clone https://github.com/ВАШ_ЛОГИН/ИМЯ_РЕПОЗИТОРИЯ.git
   cd ИМЯ_РЕПОЗИТОРИЯ
   ```
3. Создайте новую ветку:
   ```bash
   git checkout -b имя-вашей-ветки
   ```
4. Внесите изменения в .bas файл (или другие нужные файлы).

5. Зафиксируйте изменения:
   ```bash
   git add .
   git commit -m "Описание изменений"
   git push origin имя-вашей-ветки
   ```

6. Перейдите на GitHub, нажмите "Compare & pull request", выберите нужную ветку и создайте PR в `main`.
7. Напишите мне в VK/TG, если хотите оперативнее.

Советы:
- Проверьте, что код работает.
- Убедитесь, что нет лишних изменений или временных файлов.
- Коммиты должны быть понятными.

---

## 🧩 Работа с `.bas` файлами (VBA-модули Word)

Файл с расширением `.bas` — это модуль макроса, экспортированный из Microsoft Word (VBA). С ним удобно работать в Git, отслеживать изменения и обмениваться кодом.

---

### 📥 Как импортировать макрос из `.bas` файла в Word

1. Открой Microsoft Word.
2. Нажми `Alt + F11`, чтобы открыть редактор VBA.
3. В меню выбери: **File → Import File…**
4. Укажи файл `ПроверкаСоответствия.bas` (или другой).
5. Макрос появится в списке модулей (обычно `Module1`).

Теперь его можно запускать через `Alt + F8`.

---

### 📤 Как сохранить/экспортировать макрос в `.bas` файл

1. В редакторе VBA (`Alt + F11`) найди нужный модуль (например, `Module1`).
2. Правый клик → **Export File…**
3. Сохрани файл с расширением `.bas` (например, `ПроверкаСоответствия.bas`).
4. Добавь этот файл в свой Git-репозиторий.

---

### ✍️ Как редактировать `.bas` файлы

- Лучше всего редактировать макросы прямо в Word (VBA редактор).
- Если нужно — `.bas` можно открыть и в обычном текстовом редакторе (например, VS Code, Notepad++).
- Избегай ручного изменения структуры файла (например, `Attribute VB_Name =`), если не уверен в действиях.

---

### 🔄 Зачем использовать `.bas`

- Удобно хранить и отслеживать изменения в Git.
- Можно делиться отдельными макросами без `.docm` или `.dotm` файлов.
- Удобно подключать в разные документы без дублирования.


## 📄 Лицензия

MIT License. Используй, улучшай, делись с коллегами.

