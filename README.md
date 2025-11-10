# showFileMetadata

A diagnostic tool to extract and display file metadata information.

## Usage

**Drag and drop** a file onto `showFileMetadata.js`

The script will generate a text file (`<filename>.data.txt`) containing all available metadata.

## Output

The generated `.data.txt` file includes:

- **Shell.NameSpace metadata** - Windows file properties (all available headers and values)
- **PowerShell/System.Drawing metadata** - EXIF tags and PropertyItems (for images)

## Requirements

- **Operating System**: Windows (tested on Windows 10 & 11)
- **Runtime**: Windows Script Host (WScript)

## Use Cases

- Check EXIF data in photos
- Debug metadata issues
- Find available file properties
