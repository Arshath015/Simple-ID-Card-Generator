# ID Card Generator with QR Code

This project provides a simple ID card generator using Python. The generated ID card includes details entered by the user and a QR code that redirects to a specified website. Additionally, the user details are stored in an Excel file for easy management.

## Features

- **Generate ID Card**: Create an ID card with customizable details.
- **QR Code**: Include a QR code on the ID card that redirects to a user-specified website.
- **Excel Storage**: Save the entered details in an Excel file for future reference.

## Usage
- Clone or Download the Repository:
      Clone the repository using Git or download the ZIP file from GitHub.

- Open Jupyter Notebook:
      Launch Jupyter Notebook and open a new or existing notebook.

- Copy and Paste the Code:
      Copy the provided Python code into a new cell in Jupyter Notebook.

- Run the Code:
      Execute the cell. The code will prompt you to enter details such as Name, Student ID, Department, Phone Number, College, Blood Group, and Website URL.

- View the ID Card:
      The generated ID card will be displayed. It includes the entered details and a QR code.

- Check the Excel File:
      The details will be stored in an Excel file named id_card_details.xlsx. You can open this file with Excel or any spreadsheet software to view or manage the data.

## Requirements

- Python 3.x
- Jupyter Notebook
- Required Python Libraries:
  - `Pillow` for image processing
  - `qrcode` for generating QR codes
  - `xlsxwriter` for handling Excel files

You can install the required libraries using pip:

```bash
pip install pillow qrcode xlsxwriter

