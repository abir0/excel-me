import sys, os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import colors
from openpyxl.utils import get_column_letter


def main(arg):

    image = read_image(arg)
    # Make the image smaller
    image = resize_image(image, 240, 352)   # height, width for 240p
    list, width, height = image_to_list(image)
    #image.show()

    filename = get_filename(arg)
    wb = Workbook()
    sheet = wb.active

    # Apend the pixel values to worksheet
    for row in list:
        sheet.append(row)

    # Make conditional formatting rules
    color_scale_rule_red = ColorScaleRule(start_type="num",
                                      start_value=0,
                                      start_color="000000",
                                      end_type="num",
                                      end_value=255,
                                      end_color="FF0000")
    color_scale_rule_green = ColorScaleRule(start_type="num",
                                      start_value=0,
                                      start_color="000000",
                                      end_type="num",
                                      end_value=255,
                                      end_color="00FF00")
    color_scale_rule_blue = ColorScaleRule(start_type="num",
                                      start_value=0,
                                      start_color="000000",
                                      end_type="num",
                                      end_value=255,
                                      end_color="0000FF")

    # Add the rules to worksheet
    for i in range(1, width + 1):
        col = get_column_letter(i)
        string = col + "1" + ":" + col + str(height)
        #print("Adding color rule to cells", string)
        if i % 3 == 1:
            sheet.conditional_formatting.add(string, color_scale_rule_red)
        elif i % 3 == 2:
            sheet.conditional_formatting.add(string, color_scale_rule_green)
        elif i % 3 == 0:
            sheet.conditional_formatting.add(string, color_scale_rule_blue)

    wb.save(filename=filename)


def image_to_list(image):
    """Convert image to tabular list."""
    arr = np.array(image)
    w, h, p = arr.shape[0], arr.shape[1], arr.shape[2]
    arr = arr.reshape((h, w*p))
    return arr.tolist(), (w*p), h


def resize_image(image, height, width):
    """Resize the image (to work with less values)."""
    return image.resize((height, width))


def read_image(path):
    """Read an image from a file path"""
    try:
        image = Image.open(path)
        return image
    except Exception as e:
        print(e)


def get_filename(path):
    """Parse filepath, remove old file and retuen the filename."""
    name = os.path.basename(path).lower()
    name = name.replace(".jpeg", "").replace(".jpg", "").replace(".png", "")
    name += ".xlsx"
    if os.path.exists(name):
        os.remove(name)
    return name


if __name__ == "__main__":
    main(sys.argv[1])
