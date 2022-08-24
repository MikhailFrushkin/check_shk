import csv

from pyzbar import pyzbar
import cv2
import os
import pandas as pd


def draw_barcode(decoded, image):
    image = cv2.rectangle(image, (decoded.rect.left, decoded.rect.top),
                          (decoded.rect.left + decoded.rect.width, decoded.rect.top + decoded.rect.height),
                          color=(0, 255, 0),
                          thickness=5)
    return image


def decode(image, name):
    decoded_objects = pyzbar.decode(image)
    if len(decoded_objects) == 0:
        print('Штрихкод не распознан на фото: {}'.format(name))
        bad_list.append(name)
    for obj in decoded_objects:
        print(f"Обнаружен штрих-код:\n{obj}")
        image = draw_barcode(obj, image)
        print("Тип:", obj.type)
        print("Данные:", obj.data)
        print()
        data.append(str(obj.data)[2:-1])
    return image


def replace_bad_shk(bad):
    file_source = ''
    file_destination = 'bad/'

    get_files = bad

    for g in get_files:
        os.replace(file_source + g, file_destination + g)


def to_exsel(good_shk):
    """запись в csv"""
    try:
        with open('result.csv', 'w', newline='', encoding='utf-8-sig') as file:
            file_writer = csv.writer(file, delimiter=";", lineterminator="\r")
            file_writer.writerow(["Code"])

            for i in good_shk:
                file_writer.writerow([i])
    except Exception as ex:
        print(ex)
    try:
        df = pd.read_csv('result.csv', encoding='utf-8-sig', delimiter=";")
        print(df)
        df['Code'] = df['Code'].astype(str)
        writer = pd.ExcelWriter('result.xlsx')

        df.style.apply(align_left, axis=0).to_excel(writer, sheet_name='Sheet1', index=False, na_rep='NaN')
        writer.sheets['Sheet1'].set_column(0, 2, 20)
        writer.save()
    except Exception as ex:
        print(ex)
    finally:
        os.remove('result.csv')


def align_left(x):
    return ['text-align: left' for x in x]


if __name__ == "__main__":
    from glob import glob

    data = list()
    bad_list = list()
    barcodes = glob("*.jpg")
    for barcode_file in barcodes:
        try:
            img = cv2.imread(barcode_file)
            img2 = decode(img, barcode_file)
            # cv2.imshow("img", img2)
            cv2.waitKey(0)
        except Exception as ex:
            print(ex)
    # print('Список не распознанных фото: {}'.format(','.join(bad_list)))
    # print(data)
    replace_bad_shk(bad_list)
    to_exsel(data)
