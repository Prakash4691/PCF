import * as React from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react";

interface DeleteConfirmationDialogProps {
  isOpen: boolean;
  fileName: string;
  onDismiss: () => void;
  onConfirm: () => void;
}

export const DeleteConfirmationDialog: React.FC<
  DeleteConfirmationDialogProps
> = ({ isOpen, fileName, onDismiss, onConfirm }) => {
  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: "Confirm File Deletion",
        closeButtonAriaLabel: "Cancel",
        subText: `Are you sure you want to delete the file "${fileName}"?`,
      }}
      modalProps={{
        isBlocking: true,
        styles: { main: { maxWidth: 450 } },
      }}
    >
      <DialogFooter>
        <PrimaryButton onClick={onConfirm} text="Delete" />
        <DefaultButton onClick={onDismiss} text="Cancel" />
      </DialogFooter>
    </Dialog>
  );
};
