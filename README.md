# 😺 Mono Bank Google Sheets Integration

Скрипт для автоматичного завантаження транзакцій з Mono Bank в Google Sheets. Дозволяє відстежувати ваші витрати та доходи в зручному форматі таблиці.

## 🚀 Можливості

- Завантаження транзакцій з Mono Bank
- Підтримка декількох рахунків
- Автоматична категоризація транзакцій за MCC кодами
- Можливість налаштування власних правил категоризації
- Завантаження транзакцій за вказаний період

## 📋 Токен API Mono Bank

([Можна отримати тут](https://web.monobank.ua/))

## 🛠 Налаштування

1. Створіть нову Google таблицю або використайте [готовий темплейт](https://docs.google.com/spreadsheets/d/1MgAEmkjN7cwfI8STMvDsXJb_J0XWmt7kUf87qqAtrdM/edit?gid=1583520598#gid=1583520598), який містить:
   - **Усі транзакції** - основний лист для відображення всіх транзакцій з ваших рахунків
   - **Правила** - лист для налаштування автоматичної категоризації транзакцій за описом
2. Відкрийте редактор скриптів: `Розширення > Apps Script`
3. Скопіюйте код з репозиторію в редактор
4. Додайте секрети скрипта в меню `Project Settings > Script Properties`:
   - `MONO_TOKEN` - ваш токен API Mono Bank ([отримати тут](https://api.monobank.ua/))
   - `MONO_BLACK` - ID вашого основного рахунку 
   - `MONO_WHITE` - ID додаткового рахунку
   > ⚠️ ID рахунків можна отримати, зробивши запит до [API Mono Bank](https://api.monobank.ua/docs/index.html#tag/Kliyentski-personalni-dani) в розділі "Інформація про клієнта". ID знаходиться в полі `accounts[].id`
   
   > ⚠️ Важливо: Зберігайте токен та ID рахунків як секрети скрипта для безпеки ваших даних

## 📱 Використання

### Меню Mono

Після налаштування в таблиці з'явиться меню "😺 Mono" з наступними опціями:

- **💳 Завантажити нові транзакції** - завантажує транзакції з моменту останнього оновлення
- **📅 Завантажити за період...** - дозволяє вказати конкретний період для завантаження
- **📃 Застосувати правила** - застосовує правила категоризації до існуючих транзакцій
- **❗️ Створити/перестворити табличку** - створює необхідну структуру таблиці

### Завантаження за період

При виборі "Завантажити за період":
1. Введіть початкову дату та час у форматі `DD.MM.YYYY HH:mm`
2. Введіть кінцеву дату та час у тому ж форматі
3. Максимальний період - 90 днів


## 🔧 Налаштування правил категоризації

В листі "Правила" можна налаштувати автоматичну категоризацію транзакцій за різними параметрами:
1. Додайте нове правило в таблицю
2. Вкажіть умову та результат
3. Правила застосовуються автоматично при завантаженні нових транзакцій

## 📄 Ліцензія

MIT License

## 🙏 Подяки

Цей проект створено на основі:
- [mono-google-sheets-integration](https://github.com/Kostiancheck/mono-google-sheets-integration) від Kostiancheck
- [Merchant-Category-Codes](https://github.com/Oleksios/Merchant-Category-Codes) від Oleksios для MCC кодів
