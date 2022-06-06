import tkinter
from tkinter import *
from tkinter import ttk
from pyad import pyad
import random
import string
import pandas as pd
from NewUser.data import name,my_password

class New_User:

    def interface(self):
        self.root = Tk()
        self.root.title("Создание Нового Пользователя")
        self.root.geometry("600x450")

        self.Label_Office = Label(text='Филиал:')
        self.Label_Office.place(x=10, y=22)
        self.Label_FIO = Label(text='Введите ФИО:')
        self.Label_FIO.place(x=10, y=52)
        self.Label_Telephone = Label(text='Номер мобильного телефона:')
        self.Label_Telephone.place(x=10, y=82)
        self.Label_Department = Label(text='Введите Отдел:')
        self.Label_Department.place(x=10, y=112)
        self.Label_Title = Label(text='Введите должность:')
        self.Label_Title.place(x=10, y=142)
        self.Label_Director = Label(text='ФИО руководителя:')
        self.Label_Director.place(x=10, y=172)
        self.Label_Prava = Label(text='ФИО сотрудника с аналогичной должностью')
        self.Label_Prava.place(x=10, y=202)
        self.Label_Date = Label(text='Дата выхода на работу:')
        self.Label_Date.place(x=10, y=232)
        self.Label_Mesto = Label(text='Где расположено рабочее место сотрудника')
        self.Label_Mesto.place(x=10, y=262)
        self.Label_Propusk = Label(text='Необходим ли заказ пропуска?')
        self.Label_Propusk.place(x=10, y=342)

        self.Office_list = tkinter.StringVar()
        self.Office_chosen = ttk.Combobox(self.root, width=40, textvariable=self.Office_list, state='readonly')
        self.Office_chosen["values"] = (
        'Москва', 'Краснодар', 'Казань', 'Новосибирск', 'Омск', 'Ростов-на-Дону', 'Санкт-Петербург', 'Воронеж')
        self.Office_chosen.grid(column=1, row=5)
        self.Office_chosen.place(x=150, y=20)
        self.Office_chosen.current()
        self.Text_FIO = Text(self.root, width=40, height=1, font='Arial 14')
        self.Text_FIO.place(x=150, y=50)
        self.Text_Telephone = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Telephone.place(x=271, y=80)
        self.Text_Department = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Department.place(x=271, y=110)
        self.Text_Title = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Title.place(x=271, y=140)
        self.Text_Director = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Director.place(x=271, y=170)
        self.Text_Prava = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Prava.place(x=271, y=200)
        self.Text_Date = Text(self.root, width=29, height=1, font='Arial 14')
        self.Text_Date.place(x=271, y=230)
        self.Text_Mesto = Text(self.root, width=29, height=3, font='Arial 14')
        self.Text_Mesto.place(x=271, y=260)

        self.domen = IntVar()
        self.domen.set(0)
        self.Domen_Solber = Radiobutton(text='@solber.ru', variable=self.domen, value=1)
        self.Domen_Kubis = Radiobutton(text='@kubis.ru', variable=self.domen, value=0)
        self.Domen_Kubis.place(x=10, y=285)
        self.Domen_Solber.place(x=10, y=305)

        self.Propusk = IntVar()
        self.Propusk.set(0)
        self.Propusk_Net = Radiobutton(text='Нет', variable=self.Propusk, value=1)
        self.Propusk_Da = Radiobutton(text='Да', variable=self.Propusk, value=0)
        self.Propusk_Net.place(x=271, y=342)
        self.Propusk_Da.place(x=321, y=342)

        self.Button_action = Button(self.root, text="Создать учетную запись", width=20, height=1, fg='black', font='arial 14',
                               command=New_User().interface())
        self.Button_action.place(x=60, y=390)

        self.root.mainloop()

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

        def save_user(self):

        #########Запись данных пользователя в Passwords

            def align_center(x):
                    return ['text-align: center' for x in x]

            data_alternative = pd.read_excel('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords_.xlsx')
            data = pd.read_excel('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords.xlsx')
            new_row = {'Фамилия Имя Отчество': self.Original_FIO,
                           'Должность': self.title,
                           'Мобильный телефон': self.mobile_telephone,
                           'Логин': self.sAMAccountName.lower(),
                           'Пароль': self.user_password}
            if data.closed == True:
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
            if data.closed == False:
                data_alternative = data_alternative.append(new_row, ignore_index=True)
                pd.set_option('display.max_colwidth', None)
                with pd.ExcelWriter('//kkzsql04/KUBIS_DATA/IT/Совместно/SERVICEDESK/Passwords_.xlsx',
                                    engine='xlsxwriter') as wb:
                    data_alternative.style.apply(align_center, axis=0).to_excel(wb, sheet_name='Passwords', index=False)
                    sheet = wb.sheets['Passwords']
                    sheet.set_column(0, 0, 40)
                    sheet.set_column(1, 1, 60)
                    sheet.set_column(2, 2, 20)
                    sheet.set_column(3, 3, 20)
                    sheet.set_column(4, 4, 20)
                    sheet.set_column(5, 5, 15)
                    sheet.set_column(6, 6, 15)
                    sheet.set_column(7, 7, 40)





