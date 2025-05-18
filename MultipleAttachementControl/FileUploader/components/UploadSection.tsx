import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar } from "@fluentui/react/lib/MessageBar";
import { PrimaryButton } from "@fluentui/react/lib/Button";

interface UploadSectionProps {
  isUploading: boolean;
  operationType?: string;
  uploadMessage?: { text: string; type: any } | null;
  onSubmit: () => void;
  disabled: boolean;
}

export const UploadSection: React.FC<UploadSectionProps> = ({
  isUploading,
  operationType,
  uploadMessage,
  onSubmit,
  disabled,
}) => {
  return (
    <div className="upload-section">
      {isUploading && (
        <div className="progress-container">
          <Spinner
            size={SpinnerSize.large}
            label={operationType}
            labelPosition="right"
            styles={{ label: { marginLeft: "8px" } }}
          />
        </div>
      )}

      {uploadMessage && (
        <MessageBar
          messageBarType={uploadMessage.type}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
        >
          {uploadMessage.text}
        </MessageBar>
      )}

      <div className="submit-button-container">
        <PrimaryButton
          text={isUploading ? "Uploading..." : "Upload Files"}
          onClick={onSubmit}
          disabled={disabled}
          iconProps={{
            iconName: isUploading ? "Spinner" : "CloudUpload",
          }}
        />
      </div>
    </div>
  );
};
