# Multiple File Upload PCF Control Project Requirements

## Project Tech Stack

- React + Typescript (Latest Version)
- Fluent UI (Latest Version)
- Power Apps Component Framework (PCF)
- Microsoft Dataverse

## Phase 1: Basic File Upload Interface

### Requirements

1. Create a box-styled upload control
   - Auto-adjusting height similar to multi-line text field
   - Drag and drop functionality
   - File selection through browser dialog
2. Implement file list management
   - Store selected files in state
   - Display file names
   - Allow file removal

## Phase 2: File Preview and Icons

### Requirements

1. File icon display system
   - Show appropriate icons based on file type
   - Use Fluent UI icon set
2. File preview capabilities
   - Basic file information (size, type)
   - Organize files in grid layout

## Phase 3: Field Value Management

### Requirements

1. Data format handling
   - Store file names with file format in comma-separated. for example: file1.pdf, file2.jpg, file3.txt
   - Ensure data consistency
2. Value synchronization
   - Parse comma-separated file names on control load
   - Display existing files in upload interface
   - Update field value after new uploads
   - Maintain field value on file removal

## Phase 4: Note Creation and Attachment

### Requirements

1. Submit button implementation
   - Position within control
   - Enable/disable based on file selection
2. Dataverse integration
   - Create Note records for each uploaded file. Title as file name and add as an attachement of each respective file.
   - Link/Associate Notes to parent record. Create a parameter called parentEntityName to set while configuring the control. For example opportunity
   - Associate notes record to parent entity/table.
3. Progress indication
   - Show upload status
   - Display success/error messages

## Phase 5: Note Deletion with Attachments

### Requirements

1. Remove File button
   - Delete notes record from the parent table
   - Before deletion ask user want to delete the file or not, after confirmation delete the file
2. Progress indication
   - Show deleted status
   - Display success/error messages

## Technical Considerations

1. File size limits
   - Respect Dataverse attachment limits
   - Implement file size validation
2. Performance optimization
   - Efficient file handling
   - Batch processing for multiple files
3. Error handling
   - Validation messages
   - Upload error recovery

## Acceptance Criteria

- Control works in model-driven apps
- Files successfully upload and attach
- Note records are properly created
- UI is responsive and user-friendly
- All file operations have proper error handling
