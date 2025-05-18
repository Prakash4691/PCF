import { IInputs } from "../generated/ManifestTypes";

export class DataverseNotesOperations {
  private context: ComponentFramework.Context<IInputs>;

  constructor(context: ComponentFramework.Context<IInputs>) {
    this.context = context;
  }

  public async getNotesId(
    fileName: string,
    parentRecordId: string
  ): Promise<string | null> {
    try {
      const response = await this.context.webAPI.retrieveMultipleRecords(
        "annotation",
        `?$select=annotationid&$filter=subject eq '${fileName}' and _objectid_value eq ${parentRecordId}`
      );

      if (!response.entities || response.entities.length === 0) {
        return null;
      }

      return response.entities[0]["annotationid"] as string;
    } catch (error) {
      console.error("Error retrieving notes ID:", error);
      throw new Error(`Failed to retrieve notes ID for file ${fileName}`);
    }
  }

  public async deleteNote(notesId: string): Promise<void> {
    await this.context.webAPI.deleteRecord("annotation", notesId);
  }

  public async createNoteWithAttachment(
    file: File,
    fileContent: string,
    parentId: string,
    parentEntityName: string,
    mimeType: string
  ): Promise<string> {
    const target = {
      filename: file.name,
      mimetype: mimeType,
      subject: file.name,
      [`objectid_${parentEntityName}@odata.bind`]: `/${parentEntityName}s(${parentId})`,
      isdocument: true,
      documentbody: fileContent,
    };

    const response = await this.context.webAPI.createRecord(
      "annotation",
      target
    );
    return response.id as string;
  }
}
