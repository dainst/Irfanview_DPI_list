# Irfanview DPI List

The Irfanview DPI List is a program that analyzes a folder or a series of subfolders containing images to determine their resolutions and dimensions, which are critical for assessing usability in various print and digital formats. It generates an Excel report that details these metrics for each image, aiding in the quality assurance process. It is currently set up for the publications of the German Archaeological Institute.

## Goals & Usefulness

This tool streamlines the process of checking image resolution and dimensions across a collection, providing a first step before manual inspection. By leveraging IrfanView's capabilities to generate image metadata and organizing this data into an informative Excel report, users can quickly identify images that meet their specific requirements for quality and suitability.

## Excel Structure

The output of the program is an Excel workbook composed of three detailed sheets:

### Sheet 1. DAI-Zeitschriften

This sheet lists the essential metadata extracted from each image, such as file name, image type, resolution, dimensions in pixels, and print sizes in various units. It is designed for quick reference, and gives the DPI for each image according to the standard DAI journal measurements.

The following print image widths are given for journal publication:

|Spalten|Width in cm|
|----------|----------|
|2|4.03|
|3|6.28|
|4|8.52|
|5|10.76|
|6|13|
|7|--|
|8|17.5|
|Full page|25.17|

### Sheet 2. DAI-Reihen

Similar to DAI-Zeitschriften but tailored for the DAI book series, this sheet includes comprehensive metadata while focusing on specific print dimensions and resolutions for both A4 and Überformat sizes.

### Sheet 3. Max+Interactive

This interactive sheet provides dynamic calculations of potential print sizes at 2 different DPI settings, allowing users to see the corresponding print dimensions for each image. Also, it allows users to input a desired image width to determine the DPI of the image at that size.

## Program Structure

The Irfanview DPI List script operates by executing several key steps:

- It first checks the current version of the script against an online repository to ensure the user has the latest version.
- The script then identifies the path to the IrfanView executable, which is necessary to extract image metadata.
-Users can select a specific directory or a parent directory to analyze multiple image folders recursively.
-Using IrfanView's command-line interface, the script extracts detailed image metadata and compiles it into a text file.
-This metadata is then parsed and organized into an Excel workbook with three specialized sheets, offering various insights and data representations for the analyzed images.
-The program employs Python's openpyxl library to create and format the Excel report, ensuring that the data is clear, accurate, and useful for the end-user.

## Note

The program does not replace the user's examination of the files. These should still be examined for resolution problems, such as low resolution images that have been interpolated to higher resolutions, or for other issues that may not be apparent from the metadata. Also, the program only examines the DPI on the X axis, so users should determine if the images are adequate or too large in the Y dimension. However, the program does give a warning if the DPI along the Y axis does not match the X axis.
