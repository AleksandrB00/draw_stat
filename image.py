from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook

class Excel():
    def __init__(self, file_name, cell_number):
        self.file_name = file_name
        self.cell_number = cell_number
        
    def get_cell_value(self):
        wb = load_workbook(filename=self.file_name, read_only=True)
        ws = wb['OD матрица']
        value = str(ws[self.cell_number].value)
        return value


class Get_Image(Excel):
    def __init__(self, file_name, cell_number, coordinate_x, coordinate_y):
        super().__init__(file_name, cell_number)
        self.coordinate_x = coordinate_x
        self.coordinate_y = coordinate_y

    def draw_image(self):
        image = Image.open('Шаблон.jpg')
        font = ImageFont.truetype("arial.ttf", 25)
        drawer = ImageDraw.Draw(image)
        drawer.text((self.coordinate_x, self.coordinate_y), Excel.get_cell_value(self), font=font, fill='black')
        image.save('Шаблон.jpg')
        image.show()
        

Value1 = Get_Image('2.xlsx', 'B4', 1118, 309)
Value1.draw_image()

Value2 = Get_Image('2.xlsx', 'B5', 500, 309)
Value2.draw_image()

Value3 = Get_Image('2.xlsx', 'E6', 1118, 900)
Value3.draw_image()
