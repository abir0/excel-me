import sys, os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill, Border
from openpyxl.styles.differential import DifferentialStyle


def main(arg):
    image = read_image(arg)
    img_array = np.array(image)
    image.show()

    filename = get_filename(arg) + ".xlsx"
    wb = Workbook()

    wb.save(filename=filename)


def read_image(path):
    """Read an image from a file path"""
    try:
        image = Image.open(path)
        return image
    except Exception as e:
        print(e)


def get_filename(path):
    """Parse filepath and retuen the name part."""
    name = os.path.basename(path).lower()
    name = name.replace(".jpeg", "").replace(".jpg", "").replace(".png", "").replace(".png", "")
    return name


if __name__ == "__main__":
    main(sys.argv[1])
