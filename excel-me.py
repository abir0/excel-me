import sys, os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill, Border
from openpyxl.styles.differential import DifferentialStyle


def main(arg):

    image = read_image(arg)
    # Make the image smaller
    image = resize_image(image, 240, 352)   # height, width for 240p
    list = image_to_list(image)
    #image.show()

    filename = get_filename(arg)
    wb = Workbook()
    sheet = wb.active

    for row in list:
        sheet.append(row)

    wb.save(filename=filename)


def image_to_list(image):
    """Convert image to tabular list."""
    arr = np.array(image)
    h, w, p = arr.shape[0], arr.shape[1], arr.shape[2]
    arr = arr.reshape((h, w*p))
    return arr.tolist()


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
