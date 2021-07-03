Excel Yourself
==============

This repo contains a command line program [excel_me](./excel_me) that turns an image into an excel sheet.

### Getting Started

Download or clone this repo, go into the downloaded folder and run the following command to install the project dependencies:

```PowerShell
pip install -r requirements.txt
```

Then to execute the program, go to the folder and run the following command:

```PowerShell
python excel_me <image_filename>
```

That's it! Open the newly created excel sheet and zoom out to see the magic.

### Example

<p align='center'>
  <img src='Lenna.png' height='500' width='500'></br>
  <i align='center'>Original PNG image</i>
</p>

Run `python excel_me Lenna.png` and the program will generate:

<p align='center'>
  <img src='https://i.imgur.com/ebDk21T.png' height='500' width='500'></br>
  <i align='center'>Zoomed out view of Lenna in excel sheet</i>
</p>
