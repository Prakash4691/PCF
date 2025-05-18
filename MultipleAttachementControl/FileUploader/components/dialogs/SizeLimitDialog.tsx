import * as React from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
} from "@fluentui/react";

interface SizeLimitDialogProps {
  isOpen: boolean;
  title: string;
  subText: string;
  onDismiss: () => void;
}

export const SizeLimitDialog: React.FC<SizeLimitDialogProps> = ({
  isOpen,
  title,
  subText,
  onDismiss,
}) => {
  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: title,
        closeButtonAriaLabel: "Close",
        subText: subText,
      }}
      modalProps={{
        isBlocking: true,
        styles: { main: { maxWidth: 450 } },
      }}
    >
      <DialogFooter>
        <DefaultButton onClick={onDismiss} text="OK" />
      </DialogFooter>
    </Dialog>
  );
};
