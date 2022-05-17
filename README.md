# image_to_xlsx
Converts a collection of images to spreadsheets, where each cell in the spreadsheet represents 1 pixel. All of the spreadsheets are saved in a single .xlsx file.

## Setup
To run image_to_xl.py, some dependencies must also be installed
```
python -m pip install -r requirements.txt
```

## Usage
```
python image_to_xl.py -i IMAGE_PATH... -n NUM_COLORS [-o OUTPUT_PATH]
```

- IMAGE_PATH... - a list of paths to the images that will be converted
- NUM_COLORS - a positive integer, less than or equal to 256, that specifies the number of discrete colors in the created spreadsheet
- OUTPUT_PATH - an optional field specifying the location the created .xlsx file will be saved to. Defaults to _.\output.xlsx_.
