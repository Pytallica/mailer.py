import pandas as pd
import aspose.pdf as ap
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import msvcrt


with open('text.txt', 'r', encoding='utf-8') as file:
    NAME = file.read() 

def create_certificate(template_path, output_path, name):

    document = ap.Document(template_path)

    txtAbsorber = ap.text.TextFragmentAbsorber("<<<<<<<<<<<<<<<<<<<<<ФИО>>>>>>>>>>>>>>>>>>>>") # Необходимо подогнать данную строку под Ваш шаблон, заполнив ей всю ширину страницы ;)

    document.pages.accept(txtAbsorber)

    textFragmentCollection = txtAbsorber.text_fragments

    for txtFragment in textFragmentCollection:
        txtFragment.text = name 

    document.save(output_path)

def send_email(to_email, subject, body, attachment):
    msg = MIMEMultipart()
    msg['From'] = 'xxxxxxxxxx@xxx.xx'  # Замените на ваш адрес
    msg['To'] = to_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment, 'rb') as f:
        attach = MIMEText(f.read(), 'base64', 'utf-8')
        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment))
        attach.add_header('Content-Type', 'application/pdf')
        msg.attach(attach)

    with smtplib.SMTP('smtp.yandex.ru',587) as server:  # Замените на SMTP-сервер вашего почтового провайдера
        server.starttls()
        server.login('xxxxxxxxxx', 'xxxxxxxxxxx')  # Замените на Ваши учетные данные
        server.send_message(msg)

excel_file = './xxxxxxxx.xlsx'  # Путь к Вашему файлу Excel
df = pd.read_excel(excel_file)

template_path = '123_OCR.pdf'  # Путь к Вашему шаблону сертификата

for index, row in df.iterrows():
    full_name = (row.iloc[0]) # Предполагаем, что имя находится в первом столбце
    
    if full_name.count(' ') == 2:
        full_name_x = full_name.center(68) # Необходимо редактирование под свою ширину
        
    else:
        full_name_x = full_name.center(79) # Необходимо редактирование под свою ширину
    email = row.iloc[1]     # Предполагаем, что почта находится во втором столбце
    
    certificate_filename = f"{email}certificate_{index + 1}.pdf"  # Имя файла сертификата
    create_certificate(template_path, certificate_filename, full_name_x)
    names = f"""Уважаемый (ая) {full_name}!\n\n{NAME}""" # Необходимо редактирование под свои данные

    send_email(email, "Сертификат хххххххххх", names, certificate_filename)

print("Сертификаты успешно созданы и отправлены!")
