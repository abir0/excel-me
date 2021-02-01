import sys
from PIL import Image
import numpy as np

def main(arg):
    image = read_image(arg)
    image.show()

def read_image(path):
    """Read an image from a file path"""
    try:
        image = Image.open(path)
        return image
    except Exception as e:
        print(e)


if __name__ == "__main__":
    main(sys.argv[1])
