from pyad import pyad
import random
import string
import pandas as pd
from NewUser.data import name,my_password

class New_User:

    def create_user(self):
        ########Подключение к серверу и создание пользователя
        pyad.set_defaults(ldap_server="KKRAD03.kubis.ru",
                          username=name,
                          password=my_password)
        length = 8
        chars = string.ascii_lowercase + string.ascii_uppercase + string.digits + '!@#$%^*()'
        self.user_password = ''.join(random.choice(chars) for i in range(length))

        if self.domen.get() == 1:
            self.domen_mail = '@solber.ru'
        elif self.domen.get() == 0:
            self.domen_mail = '@kubis.ru'

        if self.Propusk.get() == 0:
            self.propusk = 'Да'
        elif self.Propusk.get() == 1:
            self.propusk = 'Нет'

        if self.Office_chosen.get() == "Москва":
           self. ou = pyad.adcontainer.ADContainer.from_dn("ou=Users MSK, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Краснодар":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Users KRSN, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Казань":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Kazan, ou=Users RegOffices, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Новосибирск":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Novosibirsk, ou=Users RegOffices, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Омск":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Omsk, ou=Users RegOffices, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Ростов-на-Дону":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Rostov-on-Don, ou=Users RegOffices, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Санкт-Петербург":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=St Petersburg, ou=Users RegOffices, dc=kubis, dc=ru")
        elif self.Office_chosen.get() == "Воронеж":
            self.ou = pyad.adcontainer.ADContainer.from_dn("ou=Voronezh, ou=Users RegOffices, dc=kubis, dc=ru")

        self.Original_FIO = self.Text_FIO.get('1.0', 'END').split()
        self.Original_FIO = self.Original_FIO[0] + ' ' + self.Original_FIO[1] + ' ' + self.Original_FIO[2]
        cyrillic = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'
        latin = 'a|b|v|g|d|e|yo|zh|z|i|y|k|l|m|n|o|p|r|s|t|u|f|kh|tc|ch|sh|shch||y|' \
                '|e|yu|ya|A|B|V|G|D|E|E|Zh|Z|I|I|K|L|M|N|O|P|R|S|T|U|F|Kh|Tc|Ch|Sh|Shch||Y||E|Yu|Ya'.split('|')
        self.Translit_FIO = self.Original_FIO.translate({ord(k): v for k, v in zip(cyrillic, latin)})
        self. Translit_FIO = self.Translit_FIO.split()
        self.Account_name = self.Translit_FIO[1] + " " + self.Translit_FIO[0]
        self.givenName = self.Translit_FIO[1]
        self.sn = self.Translit_FIO[0]
        self.userPrincipalName = self.Translit_FIO[1][0] + '.' + self.Translit_FIO[0]
        self.sAMAccountName = self.Translit_FIO[1][0] + '.' + self.Translit_FIO[0]
        self.telephone = '+7 (495) 783-30-67'
        self.mail = self.Translit_FIO[1][0] + '.' + self.Translit_FIO[0] + self.domen_mail
        self.userAccountControl = '66050'
        self.title = self.Text_Title.get('1.0', self.END)
        self.mobile_telephone = self.Text_Telephone.get('1.0', self.END)

    ##########Создание пользователя в AD

        self.new_user = pyad.aduser.ADUser.create(self.Account_name, self.ou, password=self.user_password, optional_attributes={'sn':self.sn,
                                                                                                        'givenName':self.givenName,
                                                                                                        'displayName':self.Account_name,
                                                                                                        'telephoneNumber':self.telephone,
                                                                                                        'mail':self.mail.lower(),
                                                                                                        'userPrincipalName':self.userPrincipalName.lower(),
                                                                                                        'sAMAccountName':self.sAMAccountName.lower(),
                                                                                                        'userAccountControl':self.userAccountControl,
                                                                                                        'UserPrincipalName':self.sAMAccountName.lower()+'@kubis.ru',
                                                                                                        'title':self.title,
                                                                                                        'l':self.Original_FIO})



    # def save_user(self):
    #         #########Запись данных пользователя в Passwords
    #
    #     def align_center(x):
    #             return ['text-align: center' for x in x]
    #
    #     data_alternative = pd.read_excel('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords_.xlsx')
    #     data = pd.read_excel('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords.xlsx')
    #     new_row = {'Фамилия Имя Отчество': self.Original_FIO,
    #                    'Должность': self.title,
    #                    'Мобильный телефон': self.mobile_telephone,
    #                    'Логин': self.sAMAccountName.lower(),
    #                    'Пароль': self.user_password}
    #     if data.closed == True:
    #         data = data.append(new_row, ignore_index=True)
    #         pd.set_option('display.max_colwidth', None)
    #         with pd.ExcelWriter('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords.xlsx', engine='xlsxwriter') as wb:
    #             data.style.apply(align_center, axis=0).to_excel(wb, sheet_name='Passwords', index=False)
    #             sheet = wb.sheets['Passwords']
    #             sheet.set_column(0, 0, 40)
    #             sheet.set_column(1, 1, 60)
    #             sheet.set_column(2, 2, 20)
    #             sheet.set_column(3, 3, 20)
    #             sheet.set_column(4, 4, 20)
    #             sheet.set_column(5, 5, 15)
    #             sheet.set_column(6, 6, 15)
    #             sheet.set_column(7, 7, 40)
    #     if data.closed == False:
    #         data_alternative = data_alternative.append(new_row, ignore_index=True)
    #         pd.set_option('display.max_colwidth', None)
    #         with pd.ExcelWriter('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords_.xlsx',
    #                             engine='xlsxwriter') as wb:
    #             data_alternative.style.apply(align_center, axis=0).to_excel(wb, sheet_name='Passwords', index=False)
    #             sheet = wb.sheets['Passwords']
    #             sheet.set_column(0, 0, 40)
    #             sheet.set_column(1, 1, 60)
    #             sheet.set_column(2, 2, 20)
    #             sheet.set_column(3, 3, 20)
    #             sheet.set_column(4, 4, 20)
    #             sheet.set_column(5, 5, 15)
    #             sheet.set_column(6, 6, 15)
    #             sheet.set_column(7, 7, 40)

        ############Отправка оповещения на почту
    # def send_meggage(self):
    #     addr_from = name_domen  # Адресат
    #     addr_to = "servicedesk@kubis.ru"  # Получатель
    #     password = my_password  # Пароль
    #
    #     msg = MIMEMultipart()  # Создаем сообщение
    #     msg['From'] = addr_from  # Адресат
    #     msg['To'] = addr_to  # Получатель
    #     msg['Subject'] = 'На сервере был создан новый пользователь - ' + self.Original_FIO  # Тема сообщения
    #
    #     html = f"""\
    #     <html>
    #         <head></head>
    #         <body>
    #             <table border="4">
    #                         <tr> Филиал - {self.Office_chosen.get()} </tr>
    #                         <tr> ФИО - {self.Original_FIO} </tr>
    #                         <tr> Номер мобильного телефона - {self.mobile_telephone} </tr>
    #                         <tr> Отдел - {self.Text_Department.get('1.0', self.END)} </tr>
    #                         <tr> Должность - {self.title} </tr>
    #                         <tr> ФИО руководителя - {self.Text_Director.get('1.0', self.END)} </tr>
    #                         <tr> ФИО сотрудника с аналогичной должностью - {self.Text_Prava.get('1.0', self.END)} </tr>
    #                         <tr> Дата выхода на работу - {self.Text_Date.get('1.0', self.END)} </tr>
    #                         <tr> Где расположено рабочее место - {self.Text_Mesto.get('1.0', self.END)} </tr>
    #                         <tr> Домен - {self.domen_mail} </tr>
    #                         <tr> Логин - {self.sAMAccountName.lower()} </tr>
    #                         <tr> Пароль - {self.user_password} </tr>
    #                         <tr> Заказ пропуска - {self.propusk} </tr>
    #             </table>
    #         </body>
    #     </html>
    #     """
    #     msg.attach(MIMEText(html, 'html', 'utf-8'))  # Добавляем в сообщение HTML-фрагмент
    #
    #     server = smtplib.SMTP('mobile.kubis.ru', 587)  # Создаем объект SMTP
    #     server.set_debuglevel(True)  # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
    #     server.starttls()  # Начинаем шифрованный обмен по TLS
    #     server.login(addr_from, password)  # Получаем доступ
    #     server.send_message(msg)  # Отправляем сообщение
    #     server.quit()

    # ###########Создание нового окна
    # def create_window(self):
    #     def close_window(self):
    #         new_root.destroy()
    #         Office_chosen.set('')
    #         Text_FIO.delete(1.0, END)
    #         Text_Telephone.delete(1.0, END)
    #         Text_Department.delete(1.0, END)
    #         Text_Title.delete(1.0, END)
    #         Text_Director.delete(1.0, END)
    #         Text_Prava.delete(1.0, END)
    #         Text_Date.delete(1.0, END)
    #         Text_Mesto.delete(1.0, END)
    #     def close_programm():
    #         root.destroy()
    #
    #     new_root = Toplevel()
    #     new_root.title("Вывод")
    #     new_root.geometry("365x100")
    #     Label_Conclusion = Label(new_root, text='Пользователь успешно создан')
    #     Label_Conclusion.place(x=90, y=22)
    #     Button_close = Button(new_root, text = "Завершить работу", width=25,height=1,fg='black',font='arial 8',command=close_programm)
    #     Button_close.place(x=190, y=52)
    #     Button_close_window = Button(new_root, text='Создать еще пользователя', width=25,height=1,fg='black',font='arial 8',command=close_window)
    #     Button_close_window.place(x=10, y=52)
    #     new_root.mainloop()
    #
    # create_window()
    #
    #
    # #####Редактирование данных
    #
    # def edit_User(self):
    #
    #     new_root = Toplevel()
    #     new_root.title("Редактирование")
    #     new_root.geometry("600x450")
    #     Label_FIO_Edit = Label(new_root, text='Введите ФИО:')
    #     Label_FIO_Edit.place(x=10, y=22)
    #     Label_Edit = Label(new_root, text='Что необходимо изменить?')
    #     Label_Edit.place(x=10, y=52)
    #
    #     Text_Edit_FIO = Text(new_root, width=40, height=1, font='Arial 14')
    #     Text_Edit_FIO.place(x=150, y=20)
    #     Text_Edit = Text(new_root, width=37, height=14, font='Arial 14')
    #     Text_Edit.place(x=183, y=50)
    #
    #     def close_window():
    #         new_root.destroy()
    #
    # #####Отправить сообщение после редактированию
    #
    #     def Send_Edit_Message():
    #
    #         Edit_FIO = Text_Edit_FIO.get('1.0', END)
    #         Edit = Text_Edit.get('1.0', END)
    #
    #         addr_from = "f.burov@kubis.ru"  # Адресат
    #         addr_to = "servicedesk@kubis.ru"  # Получатель
    #         password = "MSUMCFZZ342511m"  # Пароль
    #
    #         msg = MIMEMultipart()  # Создаем сообщение
    #         msg['From'] = addr_from  # Адресат
    #         msg['To'] = addr_to  # Получатель
    #         msg['Subject'] = f'По пользователю {Edit_FIO} внесены изменения'   # Тема сообщения
    #
    #         html = f"""\
    #                 <html>
    #                     <head></head>
    #                     <body>
    #                         <table border="4">
    #                                     <tr> ФИО -  {Edit_FIO} </tr>
    #                                     <tr> Необходимо изменить - {Edit} </tr>
    #                         </table>
    #                     </body>
    #                 </html>
    #                 """
    #         msg.attach(MIMEText(html, 'html', 'utf-8'))  # Добавляем в сообщение HTML-фрагмент
    #
    #         server = smtplib.SMTP('mobile.kubis.ru', 587)  # Создаем объект SMTP
    #         server.set_debuglevel(True)  # Включаем режим отладки - если отчет не нужен, строку можно закомментировать
    #         server.starttls()  # Начинаем шифрованный обмен по TLS
    #         server.login(addr_from, password)  # Получаем доступ
    #         server.send_message(msg)  # Отправляем сообщение
    #         server.quit()
    #
    #     Button_Send = Button(new_root, text="Отправить изменения", width=25, height=1, fg='black', font='arial 10', command=Send_Edit_Message)
    #     Button_Send.place(x=350, y=390)
    #
    #     Button_Close = Button(new_root, text="Назад", width=10, height=1, fg='black', font='arial 10', command=close_window)
    #     Button_Close.place(x=240, y=390)
    #
    #     new_root.mainloop()

result = New_User()
result.create_user()