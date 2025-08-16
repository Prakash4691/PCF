import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { Text } from "@fluentui/react/lib/Text";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Separator } from "@fluentui/react/lib/Separator";

interface FileHeaderProps {
  fileCount: number;
  onAddFiles: () => void;
  /** Mode determines which label to show: 'selected' (new pending files) or 'uploaded' (all existing/completed) */
  mode?: "selected" | "uploaded";
}

export const FileHeader: React.FC<FileHeaderProps> = ({
  fileCount,
  onAddFiles,
  mode = "selected",
}) => {
  return (
    <>
      <div className="files-header">
        <Text
          variant="mediumPlus"
          styles={{
            root: {
              fontWeight: 600,
              color: "var(--colorNeutralForeground1, #323130)",
            },
          }}
        >
          {mode === "uploaded" ? "Uploaded Files" : "Selected Files"} (
          {fileCount})
        </Text>
        <DefaultButton
          className="add-files-button"
          onClick={onAddFiles}
          iconProps={{ iconName: "Add" }}
          styles={{
            root: {
              border: "1px solid var(--colorBrandStroke1, #0078d4)",
              backgroundColor: "transparent",
              color: "var(--colorBrandForeground1, #0078d4)",
            },
            rootHovered: {
              backgroundColor: "var(--colorBrandBackground, #0078d4)",
              color: "var(--colorNeutralForegroundOnBrand, #ffffff)",
              borderColor: "var(--colorBrandBackground, #0078d4)",
            },
            rootPressed: {
              backgroundColor: "var(--colorBrandBackgroundPressed, #005a9e)",
              color: "var(--colorNeutralForegroundOnBrand, #ffffff)",
              borderColor: "var(--colorBrandBackgroundPressed, #005a9e)",
            },
          }}
        >
          Add Files
        </DefaultButton>
      </div>
      <Separator
        styles={{
          root: {
            color: "var(--colorNeutralStroke2, #d1d1d1)",
            selectors: {
              "::before": {
                backgroundColor: "var(--colorNeutralStroke2, #d1d1d1)",
              },
            },
          },
        }}
      />
    </>
  );
};
