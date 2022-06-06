import subprocess
import tkinter
from tkinter import *
from pyad import pyad
import secrets
import string
from tkinter import ttk
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def new_user():

    user_data = {}

    ########Подключение к серверу и создание пользователя
    pyad.set_defaults(ldap_server="KKRAD03.kubis.ru",
                      username="f.burov",
                      password="MSUMCFZZ342511m")

    #####генерирование пароля
    def create_password():

        length = 8
        chars = string.ascii_letters + string.digits + '!@#$%^*()'
        while True:
            user_password = ''.join(secrets.choice(chars) for i in range(length))
            if (sum(c.islower() for c in user_password) == 2
                    and sum(c.isupper() for c in user_password) == 2
                    and sum(c.isdigit() for c in user_password) == 2):
                break

        user_data['password'] = user_password

    create_password()

    ####создание пользователя в AD
    def create_user():

        if domen.get() == 1:
            domen_mail = '@solber.ru'
        elif domen.get() == 0:
            domen_mail = '@kubis.ru'

        if Propusk.get() == 0:
            propusk = 'Да'
        elif Propusk.get() == 1:
            propusk = 'Нет'

        if Office_chosen.get() == "Москва":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Users MSK, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Краснодар":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Users KRSN, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Казань":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Kazan, ou=Users RegOffices, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Новосибирск":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Novosibirsk, ou=Users RegOffices, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Омск":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Omsk, ou=Users RegOffices, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Ростов-на-Дону":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Rostov-on-Don, ou=Users RegOffices, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Санкт-Петербург":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=St Petersburg, ou=Users RegOffices, dc=kubis, dc=ru")
        elif Office_chosen.get() == "Воронеж":
            ou = pyad.adcontainer.ADContainer.from_dn("ou=Voronezh, ou=Users RegOffices, dc=kubis, dc=ru")

        ############Сохранение данных введенных пользователем
        Original_FIO = Text_FIO.get('1.0', END).split()
        Original_FIO = Original_FIO[0] + ' ' + Original_FIO[1] + ' ' + Original_FIO[2]
        cyrillic = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'
        latin = 'a|b|v|g|d|e|yo|zh|z|i|y|k|l|m|n|o|p|r|s|t|u|f|kh|tc|ch|sh|shch||y|' \
                '|e|yu|ya|A|B|V|G|D|E|E|Zh|Z|I|I|K|L|M|N|O|P|R|S|T|U|F|Kh|Tc|Ch|Sh|Shch||Y||E|Yu|Ya'.split('|')
        Translit_FIO = Original_FIO.translate({ord(k): v for k, v in zip(cyrillic, latin)})
        Translit_FIO = Translit_FIO.split()
        Account_name = Translit_FIO[1] + " " + Translit_FIO[0]
        givenName = Translit_FIO[1]
        sn = Translit_FIO[0]
        userPrincipalName = Translit_FIO[1][0] + '.' + Translit_FIO[0]
        sAMAccountName = Translit_FIO[1][0] + '.' + Translit_FIO[0]
        telephone = '+7 (495) 783-30-67'
        mail = Translit_FIO[1][0] + '.' + Translit_FIO[0] + domen_mail
        userAccountControl = '66050'
        title = Text_Title.get('1.0', END)
        mobile_telephone = Text_Telephone.get('1.0', END)
        ##########Создание пользователя в AD
        new_user = pyad.aduser.ADUser.create(Account_name, ou, password=user_data['password'],
                                             optional_attributes={'sn': sn,
                                                                  'givenName': givenName,
                                                                  'displayName': Account_name,
                                                                  'telephoneNumber': telephone,
                                                                  'mail': mail.lower(),
                                                                  'userPrincipalName': userPrincipalName.lower(),
                                                                  'sAMAccountName': sAMAccountName.lower(),
                                                                  'userAccountControl': userAccountControl,
                                                                  'UserPrincipalName': sAMAccountName.lower() + '@kubis.ru',
                                                                  'title': title,
                                                                  'l': Original_FIO})

        user_data['Original_FIO'] = Original_FIO
        user_data['mobile_telephone'] = mobile_telephone
        user_data['title'] = title
        user_data['domen_mail'] = domen_mail
        user_data['sAMAccountName'] = sAMAccountName
        user_data['propusk'] = propusk

    create_user()

    ############Отправка оповещения на почту
    def send_message_servicedesk():

        addr_from = "f.burov@kubis.ru"  # Адресат
        addr_to = "f.burov@kubis.ru"  # Получатель
        password = "MSUMCFZZ342511m"  # Пароль

        msg = MIMEMultipart()  # Создаем сообщение
        msg['From'] = addr_from  # Адресат
        msg['To'] = addr_to  # Получатель
        msg['Subject'] = 'На сервере был создан новый пользователь - ' + user_data['Original_FIO']  # Тема сообщения

        html = f"""\
        <html>
            <head></head>
            <body>
                <table border="4">
                            <tr> Филиал - {Office_chosen.get()} </tr>
                            <tr> ФИО - {user_data['Original_FIO']} </tr>
                            <tr> Номер мобильного телефона - {user_data['mobile_telephone']} </tr>
                            <tr> Отдел - {Text_Department.get('1.0', END)} </tr>
                            <tr> Должность - {user_data['title']} </tr>
                            <tr> ФИО руководителя - {Text_Director.get('1.0', END)} </tr>
                            <tr> ФИО сотрудника с аналогичной должностью - {Text_Prava.get('1.0', END)} </tr>
                            <tr> Дата выхода на работу - {Text_Date.get('1.0', END)} </tr>
                            <tr> Где расположено рабочее место - {Text_Mesto.get('1.0', END)} </tr>
                            <tr> Домен - {user_data['domen_mail']} </tr>
                            <tr> Логин - {user_data['sAMAccountName'].lower()} </tr>
                            <tr> Пароль - {user_data['password']} </tr>
                            <tr> Заказ пропуска - {user_data['propusk']} </tr>
                </table>
            </body>
        </html>
        """
        msg.attach(MIMEText(html, 'html', 'utf-8'))  # Добавляем в сообщение HTML-фрагмент

        server = smtplib.SMTP('mobile.kubis.ru', 587)  # Создаем объект SMTP
        server.set_debuglevel(True)  # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
        server.starttls()  # Начинаем шифрованный обмен по TLS
        server.login(addr_from, password)  # Получаем доступ
        server.send_message(msg)  # Отправляем сообщение
        server.quit()

    send_message_servicedesk()

    ###Закрытие файла passwords через computer manager на KKZSQL04
    def close_passwords():

        command = '$s = New-CIMSession –Computername KKZSQL04; Get-SMBOpenFile -CIMSession $s | where {$_.Path –like "*Passwords.xlsx"} | Close-SMBOpenFile -CIMSession $s -Force'
        proc = subprocess.Popen(['powershell', command])
        proc.wait()

    close_passwords()

    ###########сохранение данных в passwords
    def save_passwords():
        #########Запись данных пользователя в Passwords
        def align_center(x):
            return ['text-align: center' for x in x]

        data = pd.read_excel('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords.xlsx')

        new_row = {'Фамилия Имя Отчество': user_data['Original_FIO'],
                   'Должность': user_data['title'],
                   'Мобильный телефон': user_data['mobile_telephone'],
                   'Логин': user_data['sAMAccountName'].lower(),
                   'Пароль': user_data['password']}

        data = data.append(new_row, ignore_index=True)
        pd.set_option('display.max_colwidth', None)
        with pd.ExcelWriter('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords.xlsx', engine='xlsxwriter') as wb:
            data.style.apply(align_center, axis=0).to_excel(wb, sheet_name='Passwords', index=False)
            sheet = wb.sheets['Passwords']
            sheet.set_column(0, 0, 40)
            sheet.set_column(1, 1, 60)
            sheet.set_column(2, 2, 20)
            sheet.set_column(3, 3, 20)
            sheet.set_column(4, 4, 20)
            sheet.set_column(5, 5, 15)
            sheet.set_column(6, 6, 15)
            sheet.set_column(7, 7, 40)

    save_passwords()

    ###########Создание нового окна
    def create_window():
        def close_window():
            new_root.destroy()
            Office_chosen.set('')
            Text_FIO.delete(1.0, END)
            Text_Telephone.delete(1.0, END)
            Text_Department.delete(1.0, END)
            Text_Title.delete(1.0, END)
            Text_Director.delete(1.0, END)
            Text_Prava.delete(1.0, END)
            Text_Date.delete(1.0, END)
            Text_Mesto.delete(1.0, END)
        def close_programm():
            root.destroy()

        new_root = Toplevel()
        new_root.title("Вывод")
        new_root.geometry("365x100")
        Label_Conclusion = Label(new_root, text='Пользователь успешно создан')
        Label_Conclusion.place(x=90, y=22)
        Button_close = Button(new_root, text = "Завершить работу", width=25,height=1,fg='black',font='arial 8',command=close_programm)
        Button_close.place(x=190, y=52)
        Button_close_window = Button(new_root, text='Создать еще пользователя', width=25,height=1,fg='black',font='arial 8',command=close_window)
        Button_close_window.place(x=10, y=52)
        new_root.mainloop()

    create_window()


    ####Редактирование данных

def edit_user():

    new_root = Toplevel()
    new_root.title("Редактирование")
    new_root.geometry("600x450")
    Label_FIO_Edit = Label(new_root, text='Введите ФИО:')
    Label_FIO_Edit.place(x=10, y=22)
    Label_Edit = Label(new_root, text='Что необходимо изменить?')
    Label_Edit.place(x=10, y=52)

    Text_Edit_FIO = Text(new_root, width=40, height=1, font='Arial 14')
    Text_Edit_FIO.place(x=150, y=20)
    Text_Edit = Text(new_root, width=37, height=14, font='Arial 14')
    Text_Edit.place(x=183, y=50)

    def close_window():
        new_root.destroy()

#####Отправить сообщение после редактированию

    def send_edit_message():

        Edit_FIO = Text_Edit_FIO.get('1.0', END)
        Edit = Text_Edit.get('1.0', END)

        addr_from = "f.burov@kubis.ru"  # Адресат
        addr_to = "servicedesk@kubis.ru"  # Получатель
        password = "MSUMCFZZ342511m"  # Пароль

        msg = MIMEMultipart()  # Создаем сообщение
        msg['From'] = addr_from  # Адресат
        msg['To'] = addr_to  # Получатель
        msg['Subject'] = f'По пользователю {Edit_FIO} внесены изменения'   # Тема сообщения

        html = f"""\
                <html>
                    <head></head>
                    <body>
                        <table border="4">
                                    <tr> ФИО -  {Edit_FIO} </tr>
                                    <tr> Необходимо изменить - {Edit} </tr>
                        </table>
                    </body>
                </html>
                """
        msg.attach(MIMEText(html, 'html', 'utf-8'))  # Добавляем в сообщение HTML-фрагмент

        server = smtplib.SMTP('mobile.kubis.ru', 587)  # Создаем объект SMTP
        server.set_debuglevel(True)  # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
        server.starttls()  # Начинаем шифрованный обмен по TLS
        server.login(addr_from, password)  # Получаем доступ
        server.send_message(msg)  # Отправляем сообщение
        server.quit()

    Button_Send = Button(new_root, text="Отправить изменения", width=25, height=1, fg='black', font='arial 10', command=send_edit_message)
    Button_Send.place(x=350, y=390)

    Button_Close = Button(new_root, text="Назад", width=10, height=1, fg='black', font='arial 10', command=close_window)
    Button_Close.place(x=240, y=390)

    new_root.mainloop()



root = Tk()
root.title("Создание Нового Пользователя")
root.geometry("600x450")

Label_Office = Label(text='Филиал:')
Label_Office.place(x=10, y=22)
Label_FIO = Label(text='Введите ФИО:')
Label_FIO.place(x=10, y=52)
Label_Telephone = Label(text='Номер мобильного телефона:')
Label_Telephone.place(x=10, y=82)
Label_Department = Label(text='Введите Отдел:')
Label_Department.place(x=10, y=112)
Label_Title = Label(text='Введите должность:')
Label_Title.place(x=10, y=142)
Label_Director = Label(text='ФИО руководителя:')
Label_Director.place(x=10, y=172)
Label_Prava = Label(text='ФИО сотрудника с аналогичной должностью')
Label_Prava.place(x=10, y=202)
Label_Date = Label(text='Дата выхода на работу:')
Label_Date.place(x=10, y=232)
Label_Mesto = Label(text='Где расположено рабочее место сотрудника')
Label_Mesto.place(x=10, y=262)
Label_Propusk = Label(text='Необходим ли заказ пропуска?')
Label_Propusk.place(x=10, y=342)

Office_list = tkinter.StringVar()
Office_chosen = ttk.Combobox(root, width=40, textvariable=Office_list, state='readonly')
Office_chosen["values"] = (
            'Москва', 'Краснодар', 'Казань', 'Новосибирск', 'Омск', 'Ростов-на-Дону', 'Санкт-Петербург', 'Воронеж')
Office_chosen.grid(column=1, row=5)
Office_chosen.place(x=150, y=20)
Office_chosen.current()
Text_FIO = Text(root, width=40, height=1, font='Arial 14')
Text_FIO.place(x=150, y=50)
Text_Telephone = Text(root, width=29, height=1, font='Arial 14')
Text_Telephone.place(x=271, y=80)
Text_Department = Text(root, width=29, height=1, font='Arial 14')
Text_Department.place(x=271, y=110)
Text_Title = Text(root, width=29, height=1, font='Arial 14')
Text_Title.place(x=271, y=140)
Text_Director = Text(root, width=29, height=1, font='Arial 14')
Text_Director.place(x=271, y=170)
Text_Prava = Text(root, width=29, height=1, font='Arial 14')
Text_Prava.place(x=271, y=200)
Text_Date = Text(root, width=29, height=1, font='Arial 14')
Text_Date.place(x=271, y=230)
Text_Mesto = Text(root, width=29, height=3, font='Arial 14')
Text_Mesto.place(x=271, y=260)

domen = IntVar()
domen.set(0)
Domen_Solber = Radiobutton(text='@solber.ru', variable=domen, value=1)
Domen_Kubis = Radiobutton(text='@kubis.ru', variable=domen, value=0)
Domen_Kubis.place(x=10, y=285)
Domen_Solber.place(x=10, y=305)

Propusk = IntVar()
Propusk.set(0)
Propusk_Net = Radiobutton(text='Нет', variable=Propusk, value=1)
Propusk_Da = Radiobutton(text='Да', variable=Propusk, value=0)
Propusk_Net.place(x=271, y=342)
Propusk_Da.place(x=321, y=342)

Button_action = Button(root, text="Создать учетную запись", width=20, height=1, fg='black', font='arial 14', command=new_user)
Button_action.place(x=60, y=390)

Button_editing = Button(root, text = "Отредактировать данные", width=20,height=1,fg='black',font='arial 14',command=edit_user)
Button_editing.place(x=310,y=390)

root.mainloop()
