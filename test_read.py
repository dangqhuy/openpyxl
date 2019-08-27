import openpyxl
from PIL import Image
print(openpyxl.__version__)

if __name__ == '__main__':
    wb = openpyxl.load_workbook('/home/huydang/Downloads/excelsample/English Resume-Form resume061-241196.xlsx', data_only=True)
    print(wb)
    sheet_obj = wb['Sheet1']
    print(sheet_obj._images)
    for index, img in enumerate(sheet_obj._images):
        print(img.ref)
        _img = Image.open(img.ref)
        _img.save('%s.png' % index , 'PNG')