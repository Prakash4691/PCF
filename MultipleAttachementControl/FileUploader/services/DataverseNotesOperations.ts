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

  /** Convenience: resolves either existing id or fetches fresh */
  public async ensureNotesId(
    fileName: string,
    parentRecordId: string,
    existing?: string | Promise<string | null>
  ): Promise<string | null> {
    try {
      if (existing) {
        if (typeof existing === "string") return existing;
        const v = await existing;
        if (v) return v;
      }
    } catch (e) {
      // ignore and fallback to lookup
    }
    return this.getNotesId(fileName, parentRecordId);
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

  /**
   * Retrieves full note record including documentbody for a given annotation (note) id.
   * Returns base64 content, filename and mimetype.
   */
  public async retrieveNote(
    noteId: string
  ): Promise<{ base64: string; fileName: string; mimeType: string } | null> {
    try {
      interface AnnotationRetrieve {
        documentbody?: string;
        filename?: string;
        mimetype?: string;
      }
      const response = (await this.context.webAPI.retrieveRecord(
        "annotation",
        noteId,
        "?$select=documentbody,filename,mimetype"
      )) as unknown as AnnotationRetrieve;
      if (!response) return null;
      return {
        base64: response.documentbody || "",
        fileName: response.filename || "",
        mimeType: response.mimetype || "application/octet-stream",
      };
    } catch (error) {
      console.error("Error retrieving note content", error);
      return null;
    }
  }

  /**
   * Retrieves all notes with attachments for a given parent record.
   * Returns note metadata including subject, notetext, and creation info.
   */
  public async getAllNotesForRecord(
    parentRecordId: string
  ): Promise<{
    annotationid: string;
    subject: string;
    notetext: string;
    filename: string;
    mimetype: string;
    createdon: string;
    modifiedon: string;
    createdby: string;
  }[]> {
    try {
      const response = await this.context.webAPI.retrieveMultipleRecords(
        "annotation",
        `?$select=annotationid,subject,notetext,filename,mimetype,createdon,modifiedon,_createdby_value&$filter=_objectid_value eq ${parentRecordId} and isdocument eq true and filename ne null&$orderby=modifiedon desc`
      );

      if (!response.entities || response.entities.length === 0) {
        return [];
      }

      return response.entities.map((entity: any) => ({
        annotationid: entity.annotationid,
        subject: entity.subject || entity.filename || "Untitled",
        notetext: entity.notetext || "",
        filename: entity.filename || "",
        mimetype: entity.mimetype || "application/octet-stream",
        createdon: entity.createdon || "",
        modifiedon: entity.modifiedon || "",
        createdby: entity._createdby_value || "",
      }));
    } catch (error) {
      console.error("Error retrieving all notes for record:", error);
      throw new Error(`Failed to retrieve notes for record ${parentRecordId}`);
    }
  }

  /**
   * Retrieves note summary metadata for a given annotation ID.
   * Returns subject, notetext, and other metadata without the document body.
   */
  public async getNoteSummary(
    noteId: string
  ): Promise<{
    subject: string;
    notetext: string;
    filename: string;
    mimetype: string;
    createdon: string;
    modifiedon: string;
    createdby: string;
  } | null> {
    try {
      interface AnnotationSummary {
        subject?: string;
        notetext?: string;
        filename?: string;
        mimetype?: string;
        createdon?: string;
        modifiedon?: string;
        _createdby_value?: string;
      }
      const response = (await this.context.webAPI.retrieveRecord(
        "annotation",
        noteId,
        "?$select=subject,notetext,filename,mimetype,createdon,modifiedon,_createdby_value"
      )) as unknown as AnnotationSummary;
      
      if (!response) return null;
      
      return {
        subject: response.subject || response.filename || "Untitled",
        notetext: response.notetext || "",
        filename: response.filename || "",
        mimetype: response.mimetype || "application/octet-stream",
        createdon: response.createdon || "",
        modifiedon: response.modifiedon || "",
        createdby: response._createdby_value || "",
      };
    } catch (error) {
      console.error("Error retrieving note summary", error);
      return null;
    }
  }
}
