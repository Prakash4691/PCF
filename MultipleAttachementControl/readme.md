# Multiple File Attachment (Dataverse Notes) PCF Control

Modern, responsive multi-file uploader for Model‑Driven Apps. Built with React + Fluent UI. Stores each file as a Dataverse Note (annotation) linked to a parent record. Supports drag & drop, blocked extensions, per‑file progress, preview, download, and deletion with confirmation.

## Key Features

- Drag & drop or device file picker (`Device.pickFile`)
- Multi-file batch upload with per‑file progress & status
- Duplicate file name detection (prompts instead of overwriting)
- Blocked extensions (semicolon separated list) with consolidated dialog feedback
- Max file size enforcement (per file) before upload (handled at selection / pick)
- Inline delete (removes underlying annotation) with confirmation
- Lightweight existing file representation (1‑byte placeholder File objects until preview/download)
- Previews: image, PDF (iframe), and text / JSON / XML (rendered text). All types downloadable.
- Responsive layout via `trackContainerResize(true)` and width breakpoints (`is-xs|is-sm|is-md|is-lg`)
- Event `filesUploaded` fired after all selected new files complete successfully

## Installation / Build

1. Clone repository
2. Install dependencies inside `MultipleAttachementControl`:
   ```bash
   npm install
   ```
3. Build the control:
   ```bash
   npm run build
   ```
4. (Optional) Local harness / debug:
   ```bash
   npm start
   ```
5. Package & deploy (import the prebuilt solution zip under `FileUploaderSolution` or create your own solution and add the control).
6. Add the control to a Model‑Driven form and map properties (see below).

## Dataverse Configuration Properties

| Manifest Name              | Type / Usage              | Required | Purpose / Notes                                                                                                                 |
| -------------------------- | ------------------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `parentEntityName`         | Input (Single Line Text)  | Yes      | Logical name of table (e.g. `account`, `opportunity`). Used to build annotation relationship bind path.                         |
| `parentRecordId`           | Bound (Single Line Text)  | Yes      | The GUID of the current record (usually map to primary key column). Control will not upload if missing.                         |
| `fileData`                 | Bound (Multiple)          | Yes      | Stores a comma‑separated list of URI‑encoded file names. Used to render existing files. Modify only via control.                |
| `isUploading`              | Bound (Two Options)       | Yes      | Reflects current upload/delete activity (true while processing). Can be surfaced in form logic or ribbon rules.                 |
| `blockedFileExtension`     | Input (Multiple)          | Yes      | Semicolon separated list of disallowed extensions (e.g. `exe;bat;cmd`). Comparison is case‑insensitive, no leading dots needed. |
| `maxFileSizeForAttachment` | Input (Text - numeric KB) | Yes      | Max file size (KB). Converted to bytes internally. Files above limit prompt a size dialog.                                      |
| Event: `filesUploaded`     | Output event              | -        | Raised after all pending uploads finish successfully; use to trigger custom form logic (save, refresh, notifications).          |

> Note: A commented optional property `triggerSaveRefresh` exists in code but is not active in the manifest; auto‑save can instead be implemented by handling the `filesUploaded` event in a form script.

## Data Storage Format (`fileData`)

- Control writes: `encodeURIComponent(name1),encodeURIComponent(name2),...`
- Existing (legacy) unencoded values with commas inside names are heuristically reassembled.
- Always treat `fileData` as internal; do not manually append raw names containing commas without encoding.

## Typical Usage Flow

1. Place control on a form section targeting a record form (edit mode preferred).
2. Map properties (see table above).
3. Set allowed strategy: enter blocked extensions and maximum size (KB) in property panel.
4. Publish.
5. User drags files or clicks area to open picker; duplicates are prevented.
6. Click Upload – each file becomes an annotation; progress displayed per file.
7. Delete removes annotation (confirmation dialog) and updates `fileData`.
8. Preview or download existing files on demand (content only retrieved when needed).

## Preview & Download Support

| Capability    | Types                                                              | Behavior                                     |
| ------------- | ------------------------------------------------------------------ | -------------------------------------------- |
| Image preview | image/\*                                                           | Rendered via `<img>` with object URL caching |
| PDF preview   | application/pdf                                                    | Displayed in `<iframe>` (browser viewer)     |
| Text preview  | text/\*, application/json, application/xml, application/javascript | Contents streamed & shown as text            |
| Other types   | any                                                                | Not previewed; download still available      |

Flow: Placeholder file (size 1) -> On preview/download the control fetches annotation (`documentbody`), converts base64 to Blob, then renders or triggers download. Object URLs cached for session reuse.

## Progress & Status

- Status per file: pending → uploading (25/50/75% checkpoints) → completed or error
- Failed files are removed from list with consolidated error message referencing the last error captured
- Upload panel auto-clears progress state after a short delay (currently 3s)

## Event Handling

Handle `filesUploaded` in a form script (classic or modern) to perform follow‑up actions (save, refresh timeline, show notification). Pseudocode example:

```javascript
function onFilesUploaded(executionContext) {
  const formCtx = executionContext.getFormContext();
  formCtx.data.save(); // or custom logic
}
// Register on the PCF control's filesUploaded event in form designer
```

## Extensibility Points

- Add more MIME mappings in `utils/mimeTypes.ts`.
- Enhance annotation queries / security in `services/DataverseNotesOperations.ts`.
- Extend preview logic (e.g., Office docs -> online viewer) in `utils/filePreviewUtils.ts`.
- Re‑enable auto‑save by uncommenting the manifest property and setting a field or simply using the event.

## Troubleshooting

| Issue                                        | Cause                                      | Resolution                                                                     |
| -------------------------------------------- | ------------------------------------------ | ------------------------------------------------------------------------------ |
| No upload button action / error about parent | `parentRecordId` empty (form in Create)    | Require record save first; hide control on Create until record exists.         |
| Files silently skipped                       | Duplicate names detected                   | Rename locally or delete existing copy first.                                  |
| Blocked file dialog appears                  | Extension matches list                     | Remove extension from `blockedFileExtension` or rename file with allowed type. |
| Size limit dialog                            | Over configured `maxFileSizeForAttachment` | Increase limit (respect tenant max) or reduce file size.                       |
| Preview blank for large file                 | Non‑previewable type or unsupported viewer | Download instead; extend preview logic if needed.                              |
| Delete seems slow                            | NotesId lookup + delete call               | Network latency; monitor in browser dev tools.                                 |

### Known Limitations

- Relies on annotation entity (per‑record total storage & size governed by Dataverse limits).
- Does not batch upload across multiple records.
- No retry queue for failures (manual re‑upload required).
- Access/security relies on standard Dataverse privileges (no custom access checks).

## Responsive Behavior Summary

Height & width from host drive layout classes: <450 (`is-xs`), <700 (`is-sm`), <1024 (`is-md`), else `is-lg`. Height >0 -> container stretches; otherwise natural height. File cards flex/wrap adaptively.

## Tech Stack

- PCF (virtual control)
- React 16.14 + Fluent UI 8
- TypeScript

## Contributing

PRs welcome: lint (`npm run lint`), build (`npm run build`), keep manifest & README property tables in sync. Please describe functional changes clearly.

## License

This project is licensed under the MIT License - see the [LICENSE](../LICENSE) file for details.

## Demo

Video: https://youtu.be/IcgaxKI8nkU

---

Maintained component: feedback & enhancements encouraged.
