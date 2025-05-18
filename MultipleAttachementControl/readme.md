# React Fluent UI Multiple Attachment (File Upload) PCF Control

## Introduction

This project is a Multiple Attachment (File Upload) control built using React and Fluent UI for Microsoft Power Apps Component Framework (PCF). It enables users to upload, preview, and manage multiple file attachments (as Notes) on a Dataverse record, supporting modern drag-and-drop and file selection experiences. The control is designed for Model-Driven Apps and leverages Dataverse's annotation (note) entity for file storage.

## Configuration

To configure the control, set the following properties in the PCF control configuration:

### Property 1: `Parent Entity Name`

- **Description**: Logical name of the parent entity (table) to which files will be attached as notes.
- **Example**: `opportunity`, `account`, `contact`

### Property 2: `Parent Record ID`

- **Description**: The unique identifier (GUID) of the parent record. This is typically mapped to the record's primary key field. Select primary entity id column from list
- **Example**: `Account (Text)`

### Property 3: `File Data Field`

- **Description**: The text field in the entity where the comma-separated list of file names will be stored. Used for displaying and managing existing files.
- **Example**: `fileData`

### Property 4: `Is Uploading`

- **Description**: Create two option field and configure for isUploading property. This is to display while creation and deletion of notes.
- **Example**: `isUploading`

## Installation

Download the latest solution zip from the repo and import it into your environment, or use the Power Platform CLI:

```powershell
pac solution import --path C:\Path\To\FileUploaderSolution.zip
```

## Usage

After importing the solution, follow these steps to use the Multiple Attachment control:

1. **Choose a Text Field**: Select or create a text field in your form to bind the control (for storing file names).
2. **Configure the Control**:
   - Open the form editor and select the text field.
   - In the field properties, go to the "Components" section.
   - Add the Multiple Attachment (File Upload) control from the list.
3. **Set Properties**:
   - Set the `Parent Entity Name` and `Parent Record ID` properties as per your scenario.
   - Optionally, configure file size limits and allowed file types in the manifest if needed.
4. **Save and Publish**:
   - Save the form and publish your customizations.

## Features

- **Modern UI**: Built with Fluent UI for a responsive, accessible experience.
- **Drag-and-Drop**: Easily drag files into the upload area or use the file picker.
- **File Preview**: Shows file icons, types, and sizes for selected and existing files.
- **Multiple File Upload**: Upload several files at once, with progress indication.
- **Existing File Management**: Displays files already attached to the record; allows removal (with confirmation).
- **Note Creation**: Each uploaded file is stored as a Note (annotation) and linked to the parent record.
- **Error Handling**: User-friendly messages for upload errors, file size limits, and Dataverse constraints.
- **Responsive Design**: Works well in various form factors and screen sizes.

## How it works

- The control parses the configured text field for a comma-separated list of file names to display existing attachments.
- New files are uploaded as Notes (annotation records) and associated with the parent record.
- File removal deletes the corresponding Note after user confirmation.
- Progress and status messages are shown during upload and deletion operations.

## Technical Considerations

- **File Size Limits**: Respects Dataverse attachment limits (default 10MB per file, configurable).
- **Performance**: Efficient file handling and batch processing for multiple uploads.
- **Error Handling**: Validation and recovery for upload failures.

## Video Demo

https://www.youtube.com/watch?v=4wPPHakaq8I&t=4s
