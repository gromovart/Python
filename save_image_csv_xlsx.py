#подключаем библиотеку os
import os
#подключаем бибилиотеку request
import requests
#подключаем бибилиотеку BeautifulSoup
from bs4 import BeautifulSoup
#подключаем бибилиотеку csv
import csv
import openpyxl
filename = input('Укажите имя файла вместе с абсолютным путем:')
wb = openpyxl.load_workbook(filename)
ws = wb.active #Активация нужного листа
sheet = wb['Лист1']

def create_list(row_quantity, col_val):
  a = []
  for i in range(2, row_quantity):
    a.append(sheet.cell(row=i, column=col_val).value)
  return a

def read_csv():
  with open('fk.csv', newline='') as f:
    reader = csv.reader(f)
    for row in reader:
      print(row)



def get_file(url):
    r = requests.get(url, stream=True, timeout=(20))
    return r
 
#Функция обрезает ссылку и создает папку image. Затем выделяет из ссылки кусок который будет именем файла - артикул товара.
def get_name(url):
    name = url.split('/')[-1]
    folder = 'images'
    if not os.path.exists(folder):
        os.makedirs(folder)
    path = os.path.abspath(folder)
    return path + '/' + name
 
#Функция запускает сохранение изобоажений 
def save_image(name, file_object):
    with open(name, 'bw') as f:
        for chunk in file_object.iter_content(8192):
            f.write(chunk)

#Функция запуска цепочки сохранения изображений из файла csv
def main_fun():
  for url in create_list(50, 4):
	  save_image(get_name(url), get_file(url))


def main():
  pass
  main_fun()


if __name__ == '__main__':
    main()