#Copyright Slawa Ms 2021
import xlrd
import datetime
import sys
import xml.etree.ElementTree as ET

#Проверка версии
Version= '2.25'
newVersion = ''

#Из файла настроек 'arshinXML.ini'
Poveritel = 'Поверитель'
ExcelName = 'Для загрузки в Аршин.xls'
FIO = ''
#pathEtalon = 'etalon.dat'  #Путь к файлу "Реестр эталонов"
#pathGPE = 'gpe.dat' #Путь к файлу "Реестр ГПЭ"
#pathSIE = 'sie.dat' #Путь к файлу "Реестр СИ в качестве эталона"


try:
    ini = open('arshinXML.ini')
    for line in ini:
        if line.find('Version=') >= 0:  #Путь к версии файла
            pathVersion = line[line.find('Version=') + 8:].strip()
        if line.find('Поверитель=') >= 0 and line[line.find('Поверитель=') + 11:].strip() != '':  #Фамилия поверителя из файла настроек (для названия файла выгрузки)
            Poveritel = line[line.find('Поверитель=') + 11:].strip()
        if line.find('ФИО=') >= 0 and line[line.find('ФИО=') + 4:].strip() != '':  #Фамилия Имя Отчество поверителя из файла настроек
            FIO = line[line.find('ФИО=') + 4:].strip()
        if line.find('ExcelName=') >= 0 and line[line.find('ExcelName=') + 10:].strip() != '':    #Имя файла Execel для загрузки в Аршин
            ExcelName = line[line.find('ExcelName=') + 10:].strip()
        if line.find('pathEtalon=') >= 0 and line[line.find('pathEtalon=') + 11:].strip() != '':    #Путь к файлу с эталонами
            pathEtalon = line[line.find('pathEtalon=') + 11:].strip()            
        if line.find('pathGPE=') >= 0 and line[line.find('pathGPE=') + 8:].strip() != '':    #Путь к файлу с ГПЭ
            pathGPE = line[line.find('pathGPE=') + 8:].strip()       
        if line.find('pathSIE=') >= 0 and line[line.find('pathSIE=') + 8:].strip() != '':    #Путь к файлу с СИ в качестве эталона
            pathSIE = line[line.find('pathSIE=') + 8:].strip()   
except:
    pass    #Значения по умолчанию

#Фамилия И.О. поверителя (если необходимо и нет в файле ini)
if FIO == '':
    FIO = input('Введите Фамилия И.О. поверителя: ')



#Проверка версии программы
try:
    ver = open(pathVersion).read().split('\n') #Разбиваем на строки
    for line in ver:
        if line.find('Version=') >= 0:
            newVersion = line[line.find('Version=') + 8:].strip()
except:
    print('Ошибка доступа к серверу. Проверьте сетевые соединения.')
    sys.exit()

if newVersion != Version:
    print('Необходимо загрузить новую версию программы. Запустите файл LoadNewVersion.bat')
    sys.exit() 



err = 0 #Число ошибок
priborNum = 0 #Число выгруженных СИ
dt = datetime.datetime.now()  #Актуальная дата
log = ['Тестирование файла '+ dt.strftime('%d.%m.%Y %H:%M')] #лог ошибок
reestrEtalon = [] #Список Эталонов единиц величин
reestrGPE = [] #Список Государственных первичных эталонов
reestrSIE = [] #СИ в качестве эталона

#Открываем файл с Эталонами единиц величин
try:
    eta = open(pathEtalon).read().split('\n') #Разбиваем на строки
    for i in eta:
        if i.strip() != '':
            reestrEtalon.append(i.strip())
            
except:
    print('Ошибка открытия файла с эталонами ' + pathEtalon)
    sys.exit()

#Открываем файл с Государственными первичными эталонами
try:
    eta = open(pathGPE).read().split('\n') #Разбиваем на строки
    for i in eta:
        if i.strip() != '':
            reestrGPE.append(i.strip())
except:
    print('Ошибка открытия файла с эталонами ' + pathGPE)
    sys.exit()

#Открываем файл с СИ в качестве эталона
try:
    eta = open(pathSIE).read().split('\n') #Разбиваем на строки
    for i in eta:
        if i.strip() != '':
            reestrSIE.append(i.strip())
except:
    print('Ошибка открытия файла с эталонами ' + pathSIE)
    sys.exit()



    
#Тестируем ввод Температуры, Давления, Влажности (должны быть строковые данные)
def TestTempPressHym(TPH):
    for i in range(len(TPH)):
        if TPH[i].isalpha() == True or TPH[i] == '%' or TPH[i] == '-':
            return True
    return False










def TestCell(row, numTest):
    
    if numTest == 1: #Тестируем Столбец 'B', 'E', 'I' и 'J', если все пустые, то строка пустая
        if str(row[1]).strip() == '' and str(row[4]).strip() == '' and str(row[8]).strip() == '' and str(row[9]).strip() == '':
            return 1    
        else:
            return 0
    
    if numTest == 2: #Тестируем Столбец 'A', Если "Э" или "2", то запись не отправляем
        pole = numCorrect(str(row[0]).strip().upper())
        if pole == 'Э' or pole == '2':
            return 2
        else:
            return 0

    if numTest == 3: #Тестируем Столбец 'A', поле должно быть "пустое", "МА", "Э" (не отправлять) или "2" (не отправлять)
        pole = numCorrect(str(row[0]).strip().upper())
        if pole == '' or pole == 'МА' or pole == 'Э' or pole == '2':
            return 0
        else:
            return 3

    if numTest == 5: #Тестируем Столбец 'I', Владелец СИ
        if str(row[8]).strip() == '' or len(str(row[8]).strip())> 512:
            return 5

    if numTest == 6: #Тестируем Столбец 'J', Дата поверки СИ
        if row[9] != '':
            try:
                year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[9]), 0)         #Извлекаем дату из Excel //Значения от 43466 = 01.01.2019 до 51501 = 31.12.2040
            except:
                return 6 #Ошибка - в поле не дата/время
        else:
            return 6 #Ошибка - поле пустое




    if numTest == 7: #Тестируем Столбец 'K', Поверка действительна до (Дата след. поверки)
        if str(row[14]).strip().upper() == 'НЕТ':
            if row[10] != '': #Если СИ НЕпригодно, то след. поверки не должно быть
                return 70
        else:
            if str(row[12]).strip().upper() == 'ПЕРИОДИЧЕСКАЯ' or str(row[12]).strip().upper() == '': 
                if row[10] == '':                                       #Если поверка Периодическая и СИ Пригодно, то след. поверка должна быть
                    return 71            

        if row[10] != '':     
            try:
                year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[10]), 0)        #Извлекаем дату из Excel
            except: 
                return 72 #Ошибка - в поле не дата/время 

        if row[10] != '' and row[9] != '':     
            try:
                year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[9]), 0)        #Извлекаем дату из Excel
                year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[10]), 0)       #Извлекаем дату из Excel
                if row[10] <= row[9]:
                    return 73             #Дата след. поверки СИ должна быть больше Даты поверки СИ
            except: 
                return 74  #Ошибка - в поле не дата/время




    if numTest == 8: #Тестируем Столбец 'L', Документ на методику поверки (до 128 символов)
        if str(row[11]).strip() == '' or len(str(row[11]).strip()) > 128:
            return 8    

    if numTest == 9: #Тестируем Столбец 'AI', Состав СИ на поверку
        if len(str(row[34]).strip())> 1024:
            return 9

    if numTest == 10: #Тестируем Столбец 'AL', Прочие сведения
        if len(str(row[37]).strip())> 1024:
            return 10

    if numTest == 11: #Тестируем Столбец 'N', Использование результатов калибровки, поле должно быть "Да", "Нет" или пустое
        pole = str(row[13]).strip().upper()
        if pole == '' or pole == 'ДА' or pole == 'НЕТ':
            return 0
        else:
            return 11





    if numTest == 12: #Тестируем Столбец 'AE', Температура (до 128 символов)
        temp = str(row[30]).strip()
        if temp == '' or len(temp) > 128 or TestTempPressHym(temp) == False:
            return 12
 
    if numTest == 13: #Тестируем Столбец 'AF', Атмосферное давление (до 128 символов)
        press = str(row[31]).strip()
        if press == '' or len(press) > 128 or TestTempPressHym(press) == False:
            return 13
    
    if numTest == 14: #Тестируем Столбец 'AG', Относительная влажность (до 128 символов)
        hym = str(row[32]).strip()
        if hym == '' or len(hym) > 128 or TestTempPressHym(hym) == False:
            return 14






    if numTest == 15: #Тестируем Столбец 'AI', Состав СИ, предоставленного на поверку (до 1024 символов)
        if len(str(row[34]).strip()) > 1024:
            return 15

    if numTest == 16: #Тестируем Столбец 'AJ', Поверка в сокращенном объеме, поле должно быть "Да", "Нет" или пустое
        pole = str(row[35]).strip().upper()
        if pole == 'ДА':   #+Данные не отправляем
            return 16      #+Данные не отправляем

    if numTest == 17: #Тестируем Столбец 'O', Пригодность СИ, поле должно быть "Да", "Нет" или пустое
        pole = str(row[14]).strip().upper()
        if pole == '' or pole == 'ДА':
            znak = numCorrect(str(row[16]).strip().upper())
            if znak == '' or znak == '2' or znak == 'В ПАСПОРТЕ' or znak == 'НА СИ':
                return 0             
            else:
                return 171 #Если СИ Пригодно, то столбец 'Q' Знак поверки, должен быть заполнен правильно
        elif pole == 'НЕТ':
            reason = str(row[17]).strip()
            if reason  == '' or len(reason) > 1024:
                return 172 #Если СИ НЕпригодно, то столбец 'R' Причины непригодности, должен быть заполнен, но не больше 1024 символов
            else: return 0    
        else:
            return 17

    if numTest == 18: #Тестируем Столбец 'A', поле пустое (Сведения о единичном СИ)
        pole = numCorrect(str(row[0]).strip().upper())
        if pole == '':
            if str(row[1]).strip() == '':
                return 181 #Номер СИ по Госреестру
            if str(row[5]).strip() == '' and str(row[6]).strip() == '':
                return 182 #Заводской № или Инвентарный №
            #if str(row[4]).strip() == '':  #При пустом поле, подставляем "Нет модификации"
            #    return 183 #Модификация СИ
        else:
            return 0
         
    if numTest == 19: #Тестируем Столбец 'A', поле = "МА" (Метрологическая аттестация)
        pole = numCorrect(str(row[0]).strip().upper())
        if pole == 'МА':
            if str(row[2]).strip() == '':
                return 19 #Наименование СИ
            if str(row[5]).strip() == '' and str(row[6]).strip() == '':
                return 182 #Заводской № или Инвентарный №
            #if str(row[4]).strip() == '':  #При пустом поле, подставляем "Нет модификации"
            #    return 183 #Модификация СИ
        else:
            return 0                
    
    if numTest == 20: #Тестируем Столбец 'H', Год выпуска СИ
        if str(row[7]).strip() != '':
            try:
                x = numCorrect(str(row[7]).strip())
                int(x) #Год выпуска Си только цифрами
            except:
                return 20
                
    if numTest == 21: #Тестируем столбец "U" (Стандартный образец)            
        if str(row[20]).strip() != '':  #Тип СО (ячейка заполнена)  'U'
            STO = str(row[20]).split('|') #Разбиваем СО через | 
            stoYear = str(row[21]).split('|') #Разбиваем Год выпуска СО через |
            if len(STO) != len(stoYear):
                return 210
            for i in range(len(stoYear)):
                try:
                    x = int(stoYear[i].strip()) #Год записан цифрами без ошибок
                    #print(x)  
                    if x < 1981 or x > 2030: #Год выпуска СО
                        return 211
                except:
                    return 212 #Ошибка в записи года выпуска СО               

    if numTest == 22: #Тестируем столбец "Z" (СИ, применяемое при поверке)            
        if str(row[25]).strip() != '':  #Рег. номер СИ (ячейка заполнена)  'Z'
            regNumSI = str(row[25]).split('|') #Разбиваем рег. номера СИ через |, поле "Z" 
            zavodNumSI = str(row[26]).split('|') #Разбиваем заводские номера СИ через |, поле "AA"
            inventNumSI = str(row[27]).split('|') #Разбиваем инвент. номера СИ через |, поле "AB"
            if len(regNumSI) == len(zavodNumSI) or len(regNumSI) == len(inventNumSI):
                return 0
            else:
                return 22 #Не у всех СИ есть заводские (инвентарные) номера 

    if numTest == 23: #Тестируем столбцы "S", "T", "U", "Y", "Z", "AC", "AD",
        if str(row[18]).strip() == '' and str(row[19]).strip() == '' and str(row[20]).strip() == '' and str(row[24]).strip() == '' and str(row[25]).strip() == '' and str(row[28]).strip() == '' and str(row[29]).strip() == '':
            return 23 #Должно быть заполнено хотя бы одно из полей "S", "T", "U", "Y", "Z", "AC", "AD"

    if numTest == 24: #Дополнительные методы, столбец "AD"
        if str(row[29]).strip() != '':
            dop = str(row[29]).strip().upper()
            if dop != 'ПОВЕРКА ИМИТАЦИОННЫМ МЕТОДОМ' or dop != 'САМОПОВЕРКА' or dop != 'ПОВЕРКА РАСЧЕТНЫМ МЕТОДОМ' or dop != 'ПОВЕРКА С ИСПОЛЬЗОВАНИЕМ ПЕРВИЧНОЙ РЕФЕРЕНТНОЙ МЕТОДИКИ ИЗМЕРЕНИЙ':
                return 24



    if numTest == 25: #Тестируем Столбец 'S'. Эталоны ГПЭ
        if str(row[18]).strip() != '':  #Государственные первичные эталоны (ячейка заполнена)
            Eta = str(row[18]).split('|') #Разбиваем эталоны через |
            for i in Eta:
                if i.strip() not in reestrGPE:        
                    return 25            
    
    if numTest == 26: #Тестируем Столбец 'T'. Эталоны единиц величин
        if str(row[19]).strip() != '':  #Эталоны единицы величины (ячейка заполнена)
            Eta = str(row[19]).split('|') #Разбиваем эталоны через |
            for i in Eta:
                if i.strip() not in reestrEtalon:        
                    return 26


    if numTest == 27: #Тестируем Столбец 'Y'. СИ в качестве эталона
        if str(row[24]).strip() != '':  #СИ в качестве эталона (ячейка заполнена)
            Eta = str(row[24]).split('|') #Разбиваем эталоны через |
            for i in Eta:
                if i.strip() not in reestrSIE:        
                    return 27




                
    return 0 #Ошибки нет


def ErrorCodeText(code):
    return {
            #0: 
            3: 'Колонка "A" должна быть пустой (Единичное СИ) или содержать надпись "МА", "Э" или "2" (не отправлять)',
            5: 'В колонке "I" должен быть Владелец СИ (не больше 512 символов)',
            6: 'В колонке "J" должна быть дата поверки СИ',

                70: 'Колонка "K" (Дата след. поверки) должна быть пустая, если СИ НЕпригодно',
                71: 'В колонке "K" должна быть дата след. поверки СИ, если поверка Периодическая и СИ Пригодно',
                72: 'В колонке "K" должна быть Дата',
                73: 'Дата след. поверки СИ (колонка "K") должна быть больше Даты поверки СИ (колонка "J")',
                
                74: 'В колонках "J" и "K" должны быть Даты',

            8: 'В колонке "L" должен быть Документ на методику поверки (до 128 символов)',
            9: 'Состав СИ на поверку, В колонке "AT" должно быть не более 1024 символов',
            10: 'Прочие сведения, В колонке "AL" должно быть не более 1024 символов',
            11: 'Использование результатов калибровки. Колонка "N" должна быть "Да", "Нет" или пустая',
            12: 'В колонке "AE" должна быть Температура (до 128 символов), например 23 C',
            13: 'В колонке "AF" должно быть Атмосферное давление (до 128 символов), например 101,9 кПа',
            14: 'В колонке "AG" должна Относительная влажность (до 128 символов), например 73%',
            15: 'В колонке "AI", Состав СИ, предоставленного на поверку (до 1024 символов)',
            16: 'Поверка в сокращенном объеме. Колонка "AJ" должна быть "Да", "Нет" или пустая',
                161: 'Краткая характеристика объема поверки. Колонка "AK" должна быть заполнена',
            17: 'Пригодность СИ. Колонка "O" должна быть "Да", "Нет" или пустая',
                171: 'Если СИ Пригодно, то колонка "Q" Знак поверки, должна быть заполнена правильно',
                172: 'Если СИ НЕпригодно, то колонка "R" Причины непригодности, должна быть заполнена, но не больше 1024 символов',
                
            181: 'Необходимо заполнить Номер СИ по Госреестру(Колонка "B")',    
            182: 'Необходимо заполнить или Заводской № СИ (Колонка "F") или Инвентарный № СИ (Колонка "G")',
            #183: 'Необходимо заполнить "Модификация СИ" (Колонка "E")',
            19: 'Необходимо заполнить "Наименование СИ" (Колонка "C")',
            20: 'Год выпуска СИ должен быть только цифрами (Колонка "H")',
            
            210: 'Стандартный образец и его год выпуска (Колонка "U" и колонка "V"). Количество должно быть одинаковым',
            211: 'Год выпуска СО вне диапазона 1981-2030. (Колонка "V")',
            212: 'Ошибка в записи года выпуска СО. (Колонка "V")',
            22: 'Не у всех СИ есть заводские (инвентарные) номера. (Колонки "Z", "AA" и "AB")',
            23: 'Должно быть заполнено хотя бы одно из полей "S", "T", "U", "Y", "Z", "AC", "AD"',
            24: 'Проверьте правильность заполнения поля Дополнительные методы, колонка "AD"',
            25: 'В колонке "S" должен быть Эталон из реестра ГПЭ',
            26: 'В колонке "T" должен быть Эталон из реестра единиц величин',
            27: 'В колонке "Y" должен быть Эталон из реестра СИ',

           # 43: 'не тестируется',
            
            
            
           }[code]





#Проверка записи о СИ на ошибки
def TestOnError(row, numrow):
    codes = [] #Список кодов ошибок
 
    #Тест 1. Проверка, пустая строка или с записью о поверке
    code = TestCell(row, 1) 
    if code > 0: return -1    #Это пустая строка, запись не отправляем    

    #Тест 2. Проверка поля 'A', если "Э" или "2", то запись не отправляем
    code = TestCell(row, 2) 
    if code > 0: return -1    
    
    #Тест 3. #Проверка поля 'A', поле должно быть пустое (Единичное СИ), "МА", "Э" (не отправлять) или "2" (не отправлять)
    code = TestCell(row, 3) 
    if code > 0: codes.append(code)

    #Тест 5. Владелец СИ (поле "I")
    code = TestCell(row, 5) 
    if code > 0: codes.append(code)

    #Тест 6. Дата поверки СИ (поле "J")
    code = TestCell(row, 6) 
    if code > 0: codes.append(code)

    #Тест 7. Поверка действительна до (поле "K")
    code = TestCell(row, 7) 
    if code > 0: codes.append(code)

    #Тест 8. Документ на методику поверки (поле "L")
    code = TestCell(row, 8) 
    if code > 0: codes.append(code)

    #Тест 9. Состав СИ на поверку (поле "AI")
    code = TestCell(row, 9) 
    if code > 0: codes.append(code)
 
    #Тест 10. Прочие сведения (поле "AL")
    code = TestCell(row, 10) 
    if code > 0: codes.append(code)   

    #Тест 11. Использование результатов калибровки (поле "N")
    code = TestCell(row, 11) 
    if code > 0: codes.append(code)

    #Тест 12. Температура (поле "AE")
    code = TestCell(row, 12) 
    if code > 0: codes.append(code)
    
    #Тест 13. Атмосферное давление (поле "AF")
    code = TestCell(row, 13) 
    if code > 0: codes.append(code)
    
    #Тест 14. Относительная влажность (поле "AG")
    code = TestCell(row, 14) 
    if code > 0: codes.append(code)

    #Тест 15. Состав СИ, предоставленного на поверку (поле "AI")
    code = TestCell(row, 15) 
    if code > 0: codes.append(code)

    #Тест 16. Поверка в сокращенном объеме (поле "AJ")
    code = TestCell(row, 16) 
    #if code > 0: codes.append(code) 
    if code > 0: return -1 #+Если сокращенная поверка, то данные не отправляем

    #Тест 17. Пригодность СИ (поле "O")
    code = TestCell(row, 17) 
    if code > 0: codes.append(code)

    #Тест 18. Сведения о единичном СИ (поле "A")
    code = TestCell(row, 18) 
    if code > 0: codes.append(code)

    #Тест 19. Метрологическая аттестация (поле "A"), Наименование СИ ((поле "C"))
    code = TestCell(row, 19) 
    if code > 0: codes.append(code)

    #Тест 20. Год выпуска СИ (поле "H")
    code = TestCell(row, 20) 
    if code > 0: codes.append(code)

    #Тест 21. Стандартный образец и год выпуска
    code = TestCell(row, 21) 
    if code > 0: codes.append(code) 

    #Тест 22. СИ, применяемые при поверке
    code = TestCell(row, 22) 
    if code > 0: codes.append(code) 

    #Тест 23. Средства поверки
    code = TestCell(row, 23) 
    if code > 0: codes.append(code)

    #Тест 24. Дополнительные методы
    code = TestCell(row, 24) 
    if code > 0: codes.append(code)
    
    #Тест 25. Эталоны ГПЭ 'S'
    code = TestCell(row, 25) 
    if code > 0: codes.append(code)
    
    #Тест 26. Эталоны единиц величин  'T'
    code = TestCell(row, 26) 
    if code > 0: codes.append(code)
    
    #Тест 27. СИ в качестве эталона  'Y'
    code = TestCell(row, 27) 
    if code > 0: codes.append(code)    
    
    
    
    

    if codes == []:
        return 0    #Ошибок нет
    else:
        for i in codes:
            log.append('Ошибка в строке '+ str(numrow)+ '. '+ ErrorCodeText(i)) #Выводим в лог-файл
        global err
        err += 1;
        return -1
    
        

#===================================================================================================
def numCorrect(nc): #Если номер заканчивается на '.0', то обрезаем (особенности ячеек Excel)
    if nc[-2:] == '.0':
        return nc[0:-2]
    else:
        return nc



def miInfo_Record(miInfo, row):
    singleMI = ET.SubElement(miInfo, 'gost:singleMI')
    pole = numCorrect(str(row[0]).strip().upper())
    if pole == '':    #Единичное СИ
        mitypeNumber = ET.SubElement(singleMI, 'gost:mitypeNumber')
        mitypeNumber.text = str(row[1]).strip() #Номер в Госреестре
    elif pole == 'МА':
        crtmitypeTitle = ET.SubElement(singleMI, 'gost:crtmitypeTitle')
        crtmitypeTitle.text = str(row[2]).strip() #Наименование СИ
    
    if str(row[5]).strip() != '':  #Заводской номер СИ (если заполнено)
        manufactureNum = ET.SubElement(singleMI, 'gost:manufactureNum')
        manufactureNum.text = numCorrect(str(row[5]).strip()) #Заводской номер СИ
    elif str(row[6]).strip() != '':  #Инвентарный номер СИ (если заполнено)
        inventoryNum = ET.SubElement(singleMI, 'gost:inventoryNum')
        inventoryNum.text = numCorrect(str(row[6]).strip()) #Инвентарный номер СИ

    if str(row[7]).strip() != '':
        try:
            x = numCorrect(str(row[7]).strip())
            int(x) #Год выпуска Си только цифрами        
            manufactureYear = ET.SubElement(singleMI, 'gost:manufactureYear')
            manufactureYear.text = numCorrect(str(row[7]).strip()) #Год выпуска СИ
        except: pass
                


    
    modification = ET.SubElement(singleMI, 'gost:modification')
    if str(row[4]).strip() != '':
        modification.text = numCorrect(str(row[4]).strip()) #Модификация СИ
    else:
        modification.text = 'Нет модификации' #Нет модификации
    
    

    
def means_Record(means, row):

    if str(row[18]).strip() != '':  #Государственные первичные эталоны (ячейка заполнена)  'S'
        npe = ET.SubElement(means, 'gost:npe')
        numEta = str(row[18]).split('|') #Разбиваем эталоны через |
        for i in numEta:
            if i.strip() != '':
                number = ET.SubElement(npe, 'gost:number')
                number.text = i.strip() #Номер ГПЭ по реестру   

    if str(row[19]).strip() != '':  #Эталоны единицы величины (ячейка заполнена)  'T'
        uve = ET.SubElement(means, 'gost:uve')
        numEta = str(row[19]).split('|') #Разбиваем эталоны через |
        for i in numEta:
            if i.strip() != '':        
                number = ET.SubElement(uve, 'gost:number')      
                number.text = i.strip() #Номер эталона по реестру    
        




    if str(row[20]).strip() != '':  #Тип СО (ячейка заполнена)  'U'
        ses = ET.SubElement(means, 'gost:ses')
        numSO = str(row[20]).split('|') #Разбиваем СО через |
        yearSO = str(row[21]).split('|') #Разбиваем год выпуска СО через |
        fabNumSO = str(row[22]).split('|') #Разбиваем заводские номера СО через |
        charSO = str(row[23]).split('|') #Разбиваем характеристики СО через |
        for i in range(len(numSO)):
            if numSO[i].strip() != '':        
                se = ET.SubElement(ses, 'gost:se')
                #Номер типа СО по реестру
                typeNum = ET.SubElement(se, 'gost:typeNum')
                typeNum.text = numSO[i].strip()
                #Год выпуска СО
                manufactureYear = ET.SubElement(se, 'gost:manufactureYear')
                if i < len(yearSO):
                    manufactureYear.text = yearSO[i].strip()
                #Заводской номер СО
                if i < len(fabNumSO):
                    if fabNumSO[i].strip() != '':
                        manufactureNum = ET.SubElement(se, 'gost:manufactureNum')
                        manufactureNum.text = numCorrect(fabNumSO[i].strip())
                #Метрологические характеристики СО
                if i < len(charSO):
                    if charSO[i].strip() != '':
                        metroChars = ET.SubElement(se, 'gost:metroChars')
                        metroChars.text = charSO[i].strip()                

                        
    if str(row[24]).strip() != '':  #Средство измерения, применяемое в качестве эталона (ячейка заполнена)  'Y'
        mieta = ET.SubElement(means, 'gost:mieta')
        numEta = str(row[24]).split('|') #Разбиваем эталоны через |
        for i in numEta:
            if i.strip() != '':        
                number = ET.SubElement(mieta, 'gost:number')
                number.text = i.strip() #Номер СИ по перечню СИ             

    #Средства измерения, применяемые при поверке (ячейка заполнена)  'Z'
    if str(row[25]).strip() != '':  #Регистрационный номер типа СИ (ячейка заполнена)  'Z'
        mis = ET.SubElement(means, 'gost:mis')
        numSI = str(row[25]).split('|') #Разбиваем СИ через |
        fabNumSI = str(row[26]).split('|') #Разбиваем заводские номера СИ через |
        inventSI = str(row[27]).split('|') #Разбиваем инвентарные номера СИ через |
        for i in range(len(numSI)):
            if numSI[i].strip() != '':        
                mi = ET.SubElement(mis, 'gost:mi')
                #Регистрационный номер типа СИ
                typeNum = ET.SubElement(mi, 'gost:typeNum')
                typeNum.text = numSI[i].strip()
                #Заводской номер СИ или Инвентарный номер СИ
                if i < len(fabNumSI) or i < len(inventSI):
                    if i < len(fabNumSI) and fabNumSI[i].strip() != '':
                        manufactureNum = ET.SubElement(mi, 'gost:manufactureNum')
                        manufactureNum.text = numCorrect(fabNumSI[i].strip())    
                    if i < len(inventSI) and inventSI[i].strip() != '':
                        inventoryNum = ET.SubElement(mi, 'gost:inventoryNum')
                        inventoryNum.text = numCorrect(inventSI[i].strip())    
    
    

    if str(row[28]).strip() != '':  #Вещество (материал), применяемый при поверке (ячейка заполнена)  'AC'
        reagent = ET.SubElement(means, 'gost:reagent')
        numEta = str(row[28]).split('|') #Разбиваем реагенты через |
        for i in numEta:
            if i.strip() != '':        
                number = ET.SubElement(reagent, 'gost:number')
                number.text = i.strip() #Номер вещества (материала) по реестру 

    if str(row[29]).strip() != '':  #Дополнительные методы (ячейка заполнена)  'AD'
        dop = str(row[29]).strip().upper() 
        num = 0
        if dop == 'ПОВЕРКА ИМИТАЦИОННЫМ МЕТОДОМ': num = 1
        if dop == 'САМОПОВЕРКА': num = 2
        if dop == 'ПОВЕРКА РАСЧЕТНЫМ МЕТОДОМ': num = 3
        if dop == 'ПОВЕРКА С ИСПОЛЬЗОВАНИЕМ ПЕРВИЧНОЙ РЕФЕРЕНТНОЙ МЕТОДИКИ ИЗМЕРЕНИЙ': num = 4
        if num > 0:
            oMethod = ET.SubElement(means, 'gost:oMethod')
            oMethod.text = num            

                
def ResultSI_Record(result, row):
    miInfo = ET.SubElement(result, 'gost:miInfo')
    miInfo_Record(miInfo, row)
    
    signCipher = ET.SubElement(result, 'gost:signCipher')
    signCipher.text = 'БС' #Условный шифр знака поверки              

    if str(row[8]).strip() != '':
        miOwner = ET.SubElement(result, 'gost:miOwner')
        miOwner.text = str(row[8]).strip()   #Владелец СИ   (I[8])

    vrfDate = ET.SubElement(result, 'gost:vrfDate')        
    year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[9]), 0)         #Извлекаем дату из Excel
    vrfDate.text = datetime.datetime(year, month, day).strftime('%Y-%m-%d') + '+02:00' #Дата поверки СИ
                
    if str(row[12]).strip().upper() == 'ПЕРИОДИЧЕСКАЯ' or str(row[12]).strip().upper() == '':  # Если поверка Периодическая (поле "M" пустое или "Периодическая")
        if str(row[14]).strip().upper() != 'НЕТ':   
            validDate = ET.SubElement(result, 'gost:validDate')
            year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[10]), 0)           #Извлекаем дату из Excel
            validDate.text = datetime.datetime(year, month, day).strftime('%Y-%m-%d') + '+02:00'  #Поверка действительна до
        type = ET.SubElement(result, 'gost:type')
        type.text = '2'  #Периодическая поверка
    else:
        if str(row[10]).strip() != '': #Если есть Дата след. поверки
            validDate = ET.SubElement(result, 'gost:validDate')
            year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(row[10]), 0)           #Извлекаем дату из Excel
            validDate.text = datetime.datetime(year, month, day).strftime('%Y-%m-%d') + '+02:00'  #Поверка действительна до       
        type = ET.SubElement(result, 'gost:type')
        type.text = '1'  #Первичная поверка    

    calibration = ET.SubElement(result, 'gost:calibration')
    pole = str(row[13]).strip().upper()
    if pole == '' or pole == 'НЕТ':
        calibration.text = 'false'  #Признак калибровки - Нет
    else:
        calibration.text = 'true'   #Признак калибровки - Да

    #Пригодность СИ
    pole = str(row[14]).strip().upper()
    if pole == '' or pole == 'ДА':    #Пригодно
        applicable = ET.SubElement(result, 'gost:applicable')
        if str(row[15]).strip() != '': #Номер наклейки
            stickerNum = ET.SubElement(applicable, 'gost:stickerNum')
            stickerNum.text = numCorrect(str(row[15]).strip())  #Номер наклейки          
        znak = numCorrect(str(row[16]).strip().upper())
        signPass = ET.SubElement(applicable, 'gost:signPass')
        signMi = ET.SubElement(applicable, 'gost:signMi')
        if znak == '':
            signPass.text = 'false' #Знак поверки в паспорте - нет           
            signMi.text = 'false'   #Знак поверки на СИ - нет
        elif znak == '2':    
            signPass.text = 'true' #Знак поверки в паспорте - есть           
            signMi.text = 'true'   #Знак поверки на СИ - есть      
        elif znak == 'В ПАСПОРТЕ':    
            signPass.text = 'true' #Знак поверки в паспорте - есть           
            signMi.text = 'false'  #Знак поверки на СИ - нет      
        elif znak == 'НА СИ':    
            signPass.text = 'false' #Знак поверки в паспорте - нет           
            signMi.text = 'true'    #Знак поверки на СИ - есть      
        else: pass
    elif pole == 'НЕТ':       #Непригодно
        inapplicable = ET.SubElement(result, 'gost:inapplicable')
        reasons = ET.SubElement(inapplicable, 'gost:reasons')
        reasons.text = str(row[17]).strip() #Причины непригодности
    else: pass
    
     
    
    docTitle = ET.SubElement(result, 'gost:docTitle')   
    docTitle.text = str(row[11]).strip() #Документ Методика поверки 
    
    if FIO != '':
        metrologist = ET.SubElement(result, 'gost:metrologist')
        metrologist.text = FIO  #Фамилия Имя Отчество поверителя

    means = ET.SubElement(result, 'gost:means')
    means_Record(means, row)    



    conditions = ET.SubElement(result, 'gost:conditions')    #Условия проведения поверки        
    temperature = ET.SubElement(conditions, 'gost:temperature')   
    temperature.text = numCorrect(str(row[30]).strip())    #Температура, Поле 'AE'
    pressure = ET.SubElement(conditions, 'gost:pressure')   
    pressure.text = numCorrect(str(row[31]).strip())    #Давление, Поле 'AF'    
    hymidity = ET.SubElement(conditions, 'gost:hymidity')   
    hym = str(row[32]).strip()                          #Относительная влажность, Поле 'AG'    
    try:
        x = float(hym) #Если поле в Excel не текстовое, а %%, т.е. 76% передает как 0,76
        if x < 1:
            x *= 100
        hym = numCorrect(str(x)) + '%'
    except: pass
    #hymidity.text = numCorrect(str(row[32]).strip())    #Относительная влажность, Поле 'AG'    
    hymidity.text = hym
    
    
    
    
    if str(row[33]).strip() != '':
        other = ET.SubElement(conditions, 'gost:other')   
        other.text = str(row[33]).strip()    #Другие факторы, Поле 'AH'    
    
    if str(row[34]).strip() != '':
        structure = ET.SubElement(result, 'gost:structure')   
        structure.text = str(row[34]).strip()    #Состав СИ, предоставленного на поверку, Поле 'AI'       

    # if str(row[35]).strip().upper() == 'ДА':  #
        # brief_procedure = ET.SubElement(result, 'gost:brief_procedure')   
        # characteristics = ET.SubElement(brief_procedure, 'gost:characteristics')
        # characteristics.text = str(row[36]).strip()        
        
    if str(row[37]).strip() != '':  #Прочие сведения
        additional_info = ET.SubElement(result, 'gost:additional_info')
        additional_info.text = numCorrect(str(row[37]).strip()) #Прочие сведения


       
    
    #ranges = ET.SubElement(result, 'gost:ranges')
    #values = ET.SubElement(result, 'gost:values')
    #channels = ET.SubElement(result, 'gost:channels')   
    #blocks = ET.SubElement(result, 'gost:blocks')
    #protocol = ET.SubElement(result, 'gost:protocol')
    


def WriteFileXML(appXML):
    dt = datetime.datetime.now()  #Дата сохранения файла
    FileName = Poveritel + dt.strftime(' %Y-%m-%d %H%M%S') + '.xml' #Имя файла выгрузки результатов xml
    with open(FileName, "wb") as file:
        file.write(ET.tostring(appXML, encoding='UTF-8', xml_declaration=True))





# ======== Основной модуль ======================================================================================================= Основной модуль
rb = xlrd.open_workbook(ExcelName,formatting_info=False) #Открываем файл
sheet = rb.sheet_by_index(0) #Первый Лист 1
 



rootXML = ET.Element('gost:application')   #Корень XML-файла
rootXML.set('xmlns:gost', "urn://fgis-arshin.gost.ru/module-verifications/import/2020-06-19")






FirstRowNum = 1 #Строка начала проверки
for rownum in range(FirstRowNum, sheet.nrows): #Перебираем строки Листа 1
    row = sheet.row_values(rownum) #row - содержимое строки, rownum - номер строки
    if TestOnError(row, rownum + 1) == 0:
        resultSI = ET.SubElement(rootXML, 'gost:result')
        ResultSI_Record(resultSI, row) 
        priborNum += 1



if err == 0:
    print('Ошибок не обнаружено.')
    print('Сформировано ' + str(priborNum) + ' записей.')
    #Записываем XML в файл
    WriteFileXML(rootXML)   
else:
    print('Обнаружено ' + str(err) + ' ошибок. Подробности в файле ИсправитьОшибки.txt')
    #Записываем лог ошибок    
    with open("Исправить ошибки.txt", "w") as file:
        log.insert(1, 'Обнаружено '+ str(err)+ ' ошибок.')
        print(*log, file=file, sep="\n")









    
    
    
    