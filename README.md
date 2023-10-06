# PowerPoint Text Extractor

This utility allows you to extract text from PowerPoint presentations (.pptx files) and print it to the console. The program traverses each slide in the PowerPoint presentation and retrieves the text content, normalizing whitespace.

## Requirements

* .NET Core or .NET 5 or higher
* DocumentFormat.OpenXml library (used for interacting with .pptx files)

## Usage

1. Clone the repository:

```
git clone https://github.com/imksoo/pptx2text.git
```

2. Navigate to the cloned directory:
```
cd pptx2text
```

3. Run the program:
```
dotnet run [file1.pptx] [file2.pptx] ...
```

Replace [file1.pptx] [file2.pptx] ... with the paths to the .pptx files you want to extract text from.

## Functionality

* Check if the provided file path exists.
* Open the PowerPoint file and extract text from each slide.
* Normalize the extracted text by replacing multiple consecutive white spaces with a single space.
* Print the normalized text to the console.

## License

This project is licensed under the terms of the MIT License. Please see the LICENSE file for details.
