# Barcode Generator

## Description

The Barcode Generator is a Windows Forms application that allows users to generate barcodes from data provided in an Excel file. Each column in the Excel file corresponds to a category, and a folder is created for each category containing the generated barcode images for each data entry. The application supports customization of barcode dimensions and provides a user-friendly interface for file selection and configuration.

## Features

- **Upload Excel File**: Select an Excel file (.xlsx) to load data for barcode generation.
- **Choose Save Directory**: Select a directory to save the generated barcode images.
- **Custom Barcode Dimensions**: Specify the width and height of the barcode images.
- **Generate Barcodes**: Create barcode images for each data entry in the Excel file, organized by category.
- **Processing Indicator**: Display a processing message while generating barcodes.
- **Open Output Folder**: Automatically open the main output folder after generation.

## Prerequisites

- .NET Framework (or .NET Core/5/6) installed on your machine.
- Visual Studio for building and running the application.

## Installation

1. **Clone the Repository**:
 

2. **Open the Project in Visual Studio**:
    - Open `BarcodeGenerator.sln` in Visual Studio.

3. **Restore NuGet Packages**:
    - Visual Studio will prompt you to restore NuGet packages. Click `Restore` to download the necessary dependencies.

4. **Build the Solution**:
    - Go to `Build > Build Solution` or press `Ctrl+Shift+B`.

5. **Run the Application**:
    - Press `F5` or go to `Debug > Start Debugging`.

## Usage

1. **Upload Excel File**:
    - Click the `Upload Excel File` button and select an Excel file (.xlsx) containing the data.

2. **Choose Save Directory**:
    - Click the `Choose Save Directory` button and select a folder where the barcode images will be saved.

3. **Specify Barcode Dimensions**:
    - Enter the desired width and height for the barcode images in the provided text boxes.

4. **Generate Barcodes**:
    - Click the `Generate Barcodes` button to start generating barcode images. A processing message will appear during the operation.
    - Once the barcodes are generated, a success message will be displayed, and the main output folder will open automatically.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## Contact

For questions or support, please open an issue on GitHub or contact the repository owner.

---

**Author**: Anas Charjane

**Version**: 1.0.0

