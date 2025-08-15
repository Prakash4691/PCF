import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
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
            size={SpinnerSize.medium}
            label={operationType || "Processing..."}
            labelPosition="right"
            styles={{
              label: {
                marginLeft: "8px",
                color: "var(--colorNeutralForeground1, #323130)",
                fontSize: "var(--fontSizeBase300, 14px)",
              },
              circle: {
                borderTopColor: "var(--colorBrandBackground, #0078d4)",
              },
            }}
          />
        </div>
      )}

      {uploadMessage && (
        <MessageBar
          messageBarType={uploadMessage.type}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
          styles={{
            root: {
              borderRadius: "var(--borderRadiusMedium, 4px)",
              fontSize: "var(--fontSizeBase300, 14px)",
            },
          }}
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
          styles={{
            root: {
              backgroundColor: "var(--colorBrandBackground, #0078d4)",
              borderColor: "var(--colorBrandBackground, #0078d4)",
              color: "var(--colorNeutralForegroundOnBrand, #ffffff)",
              fontWeight: "var(--fontWeightSemibold, 600)",
              fontSize: "var(--fontSizeBase300, 14px)",
              minHeight: "32px",
              borderRadius: "var(--borderRadiusMedium, 4px)",
            },
            rootHovered: {
              backgroundColor: "var(--colorBrandBackgroundHover, #106ebe)",
              borderColor: "var(--colorBrandBackgroundHover, #106ebe)",
            },
            rootPressed: {
              backgroundColor: "var(--colorBrandBackgroundPressed, #005a9e)",
              borderColor: "var(--colorBrandBackgroundPressed, #005a9e)",
            },
            rootDisabled: {
              backgroundColor: "var(--colorNeutralBackgroundDisabled, #f3f2f1)",
              borderColor: "var(--colorNeutralStrokeDisabled, #c8c6c4)",
              color: "var(--colorNeutralForegroundDisabled, #a19f9d)",
            },
          }}
        />
      </div>
    </div>
  );
};
