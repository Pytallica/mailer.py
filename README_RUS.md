▎Сертификаты и отправка по электронной почте

Этот проект предназначен для создания персонализированных сертификатов на основе шаблона в формате PDF и их отправки по электронной почте. Он использует библиотеки pandas, aspose.pdf и smtplib для выполнения своих задач.

▎Описание

Скрипт выполняет следующие функции:

1. Чтение данных: Считывает имена участников и их адреса электронной почты из файла Excel.

2. Создание сертификатов: Создает PDF-сертификаты на основе шаблона, заменяя текст "ФИО" на имя участника.

3. Отправка по электронной почте: Отправляет созданные сертификаты на указанные адреса электронной почты.

▎Установка

Для работы скрипта необходимо установить следующие библиотеки:

pip install pandas aspose-pdf


▎Использование

1. Подготовьте файл data1.xlsx с двумя столбцами:

   • Первый столбец: полные имена участников.

   • Второй столбец: адреса электронной почты.

2. Создайте текстовый файл text.txt с текстом, который будет включен в сертификат.

3. Замените пути к файлам:

   • Путь к шаблону сертификата в переменной template_path.

   • Путь к файлу Excel в переменной excel_file.

4. Убедитесь, что вы заменили учетные данные и SMTP-сервер в функции send_email.

5. Запустите скрипт:

python ваш_скрипт.py


▎Примечания

• Убедитесь, что у вас есть доступ к SMTP-серверу и правильные учетные данные для отправки электронной почты.

• Шаблон сертификата должен содержать текст "ФИО", который будет заменен на имя участника.

• Скрипт автоматически создает сертификаты и отправляет их по электронной почте.

▎Лицензия

Этот проект лицензирован под MIT License.

▎Контрибьюция

Если вы хотите внести изменения или улучшения, пожалуйста, создайте форк репозитория и отправьте пулл-запрос.

---

Если у вас возникли вопросы или проблемы, не стесняйтесь обращаться за помощью!