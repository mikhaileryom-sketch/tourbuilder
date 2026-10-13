# TourBuilder — маркетплейс для организации мероприятий

## Что это
Платформа для артистов, площадок и организаторов. Работает как Telegram Web App.
Основные фишки: визуальный календарь занятости, ИИ-подбор артистов, гибкие цены, ролевая система.

## Стек
- Frontend: HTML + CSS + Vanilla JS (один файл index.html)
- Backend: Google Apps Script (см. code-reference.js)
- DB: Google Sheets
- Hosting: GitHub Pages
- Авторизация: Telegram Web App

## Ссылки
- App: https://mikhaileryom-sketch.github.io/tourbuilder/
- Бот: https://t.me/Quantumlover_bot
- GitHub: https://github.com/mikhaileryom-sketch/tourbuilder

## Структура папки
- index.html — фронтенд
- code-reference.js — копия Code.gs (бэкенд, живёт в Apps Script в облаке)
- README.md — этот файл
- photos/ — ассеты (placeholder.jpg и др.)

## Роли пользователей
- artist — артист (работает)
- organizer — организатор (работает)
- venue — площадка (в разработке)
- viewer — зритель (план)

## Дизайн-система
- Фон: #F5F5F7
- Карточки: #FFFFFF
- Основной текст: #1A1A1A
- Вторичный текст: #6E6E73
- Акцент (кнопки): #EC4899
- Соты артистов: золотой #D4AF37 (в процессе смены с розового)

## Важное про бэкенд
Code.gs работает в Google Apps Script. У Harness нет прямого доступа к нему.
Все правки бэкенда — в файл code-reference.js, вручную переносим в Apps Script и делаем Deploy.

## Текущие задачи
См. файл TASKS.md или последнее ТЗ в чате.