# Simple Doc-To-PPT&#x20;

**DocToPPT Converter** is a tool that converts scripts written in Microsoft Word documents into PowerPoint presentations, making it easy to project and present content seamlessly. Designed with a user-friendly interface, this program ensures that each paragraph in the Word document becomes a visually appealing slide.

---

## Features

- **Word to PowerPoint Conversion**: Transforms each paragraph in a Word document into individual PowerPoint slides.
- **Customizable Text Formatting**:
  - Bold and italicized text is preserved.
  - Text size and colors are automatically adjusted for readability.
- **Slide Design**:
  - Black background with white or yellow text for high contrast.
  - Automatic text alignment to the left for consistency.
- **Error Handling**: Ensures invalid files are flagged and guides users to provide the correct input and output formats.
- **Empty Slide Removal**: Automatically deletes any empty slides created during the conversion process.

---

## How to Use

1. **Launch the Program**:
   Run the `DocToPPT_UI.pyw` script:

   ```bash
   python DocToPPT_UI.pyw
   ```

2. **Select Input and Output Files**:

   - Use the "Select Input File" field to choose a `.docx` file.
   - Use the "Select Output File" field to specify the name and location for the PowerPoint `.pptx` file.

3. **Convert**:

   - Click the **Convert** button to start the process.
   - Once completed, a popup will confirm the successful creation of the PowerPoint file.

4. **View Slides**:

   - Open the generated `.pptx` file in PowerPoint for projection or further editing.

---

## Requirements

- Python 3.x
- Libraries:
  - `python-docx`
  - `python-pptx`
  - `PySimpleGUI`

To install the required libraries, run:

```bash
pip install python-docx python-pptx PySimpleGUI
```

---

## Recommendations

- **File Format**:
  Ensure that your input file is in `.docx` format and your output file is saved with a `.pptx` extension.

- **Slide Design**:
  Keep paragraphs concise to prevent overcrowding on slides.

---

## Installation

1. Install Python 3.x from [python.org](https://www.python.org/).
2. Install the required libraries using the command above.
3. Download this program and run `DocToPPT_UI.pyw`.

---

## Support

For any issues or suggestions, please reach out to [DonghyunWon2@gmail.com](mailto\:DonghyunWon2@gmail.com).

---

## Future Enhancements

- Support for additional text formatting and custom slide themes.
- Batch processing of multiple Word documents into PowerPoint files.
- Enhanced error feedback for unsupported file formats.

---

