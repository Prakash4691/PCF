import * as React from "react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Text } from "@fluentui/react/lib/Text";
import { isPreviewable, isTextType } from "../../utils/filePreviewUtils";

export interface PreviewDialogData {
  fileName: string;
  mimeType: string;
  objectUrl?: string; // for binary preview like images/pdf
  textContent?: string; // for text preview
}

interface PreviewDialogProps {
  isOpen: boolean;
  data: PreviewDialogData | null;
  onDismiss: () => void;
  isLoading?: boolean;
}

export const PreviewDialog: React.FC<PreviewDialogProps> = ({
  isOpen,
  data,
  onDismiss,
  isLoading,
}) => {
  const isPreviewSupported = data && isPreviewable(data.mimeType);

  const renderPreview = () => {
    if (!data) return null;
    if (!isPreviewSupported) {
      return (
        <Text>
          Preview not available for this file type. Please download to view.
        </Text>
      );
    }

    if (isLoading) {
      return <Spinner size={SpinnerSize.large} label="Loading preview..." />;
    }

    if (isTextType(data.mimeType)) {
      return (
        <pre className="preview-text">{data.textContent || "No content"}</pre>
      );
    }

    if (data.mimeType === "application/pdf") {
      return (
        <iframe
          title={data.fileName}
          src={data.objectUrl}
          className="preview-pdf-frame"
        />
      );
    }

    if (data.mimeType.startsWith("image/")) {
      return (
        <img
          src={data.objectUrl}
          alt={data.fileName}
          className="preview-image"
        />
      );
    }

    return <Text>Preview not available.</Text>;
  };

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: data?.fileName || "File Preview",
      }}
      minWidth={600}
      maxWidth={800}
    >
      <div className="preview-content-wrapper">{renderPreview()}</div>
      <DialogFooter>
        <DefaultButton onClick={onDismiss} text="Close" />
      </DialogFooter>
    </Dialog>
  );
};
