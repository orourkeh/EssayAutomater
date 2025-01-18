# README: Word Document Batch Updater

## Overview
The **Word Document Batch Updater** is a Python-based GUI application designed to streamline the batch processing of Microsoft Word documents (`.docx`). It allows users to perform the following tasks:
- Rename files with custom prefixes, suffixes, and frequent keywords from the content.
- Highlight specified keywords within the documents using custom colors.
- Append boilerplate text to all documents.
- Preview new filenames before making changes.
- Toggle between light and dark mode for a better user experience.

## Features
1. **File Selection**: Select multiple `.docx` files for batch processing.
2. **File Renaming**:
   - Add custom prefixes and suffixes to filenames.
   - Use the most frequent keyword from the document content as part of the filename.
   - Flexible options to use the keyword as a prefix or suffix.
3. **Keyword Highlighting**:
   - Highlight specified keywords in the document content.
   - Choose up to three highlight colors.
   - Bold highlighted keywords for better visibility.
4. **Boilerplate Appending**: Add a custom text block to the end of all processed documents.
5. **Dark Mode**: Switch between light and dark themes for comfortable usage.
6. **Preview Changes**: View the updated filenames in real-time before processing files.
7. **Progress Bar**: Visualize the processing progress for multiple files.

## Installation
### Prerequisites
- **Python**: Ensure you have Python 3.6 or later installed.
- **Dependencies**: Install the required Python libraries using the following command:
  ```bash
  pip install python-docx
  ```

### Running the Application
1. Save the provided code into a file named `word_batch_updater.py`.
2. Run the script using Python:
   ```bash
   python word_batch_updater.py
   ```
3. The GUI window will appear, allowing you to use the application.

## Usage
### Main Workflow
1. **Select Files**: 
   - Click the **"Select Files"** button and choose the `.docx` files you want to process.
   - The number of selected files will be displayed.
2. **File Renaming**:
   - Enter a prefix and/or suffix for new filenames.
   - Optionally, specify keywords to extract the most frequent one from each file and use it in the filename.
   - Choose whether the keyword is added as a prefix or suffix using the provided radio buttons.
   - Preview the new filenames in the **Preview** section.
3. **Keyword Highlighting**:
   - Specify keywords to highlight in the documents, each in a separate line.
   - Assign a highlight color for each keyword group (up to three groups).
4. **Boilerplate Text**:
   - Enter any text you want appended to the end of each document in the provided text box.
5. **Process Files**:
   - Click the **"Process Files"** button to apply the changes.
   - A progress bar will show the processing status.
   - A warning will appear if any files will be overwritten.

### Additional Features
- **Dark Mode**: Click the **"Toggle Darkmode"** button to switch themes.
- **Preview**: Updates to filenames are reflected in real-time in the **Preview** section.

## File Overwrite Warning
If the processed files will overwrite existing ones, the application prompts for confirmation to prevent accidental data loss.

## Error Handling
The application gracefully handles errors, displaying messages in the console if any issues occur during file processing.

## Customization
The script is designed to be user-friendly and customizable:
- Modify the `color_mapping` dictionary to add or change the available highlight colors.
- Update the default GUI theme or font size by editing the `toggle_theme` function or `default_font` configuration.

## Dependencies
- **tkinter**: For creating the graphical user interface.
- **python-docx**: For reading and writing Word documents.
- **collections.Counter**: For analyzing keyword frequencies.
- **re**: For advanced string manipulation and pattern matching.

## Limitations
- The application currently supports only `.docx` files. Older `.doc` files are not supported.
- Keywords are case-insensitive but must be complete words or phrases.
- Large files or numerous documents may take time to process due to the complexity of operations.

## Future Enhancements
- Add support for `.pdf` or other document formats.
- Provide more advanced options for keyword frequency analysis.
- Enable user-defined font styles for highlighted text.

## Support
For questions or issues, feel free to contact the developer or submit bug reports to the repository (if hosted).

---

Enjoy using the **Word Document Batch Updater** to simplify your document processing tasks!
