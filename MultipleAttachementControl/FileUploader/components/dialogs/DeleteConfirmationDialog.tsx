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
        subText: `Are you sure you want to delete the file "${fileName}"? This action cannot be undone.`,
        styles: {
          title: {
            fontSize: "var(--fontSizeBase400, 16px)",
            fontWeight: "var(--fontWeightSemibold, 600)",
            color: "var(--colorNeutralForeground1, #323130)",
          },
          subText: {
            fontSize: "var(--fontSizeBase300, 14px)",
            color: "var(--colorNeutralForeground2, #616161)",
            marginTop: "var(--spacingVerticalS, 8px)",
          },
        },
      }}
      modalProps={{
        isBlocking: true,
        styles: {
          main: {
            maxWidth: 480,
            borderRadius: "var(--borderRadiusMedium, 4px)",
            boxShadow: "var(--shadow16, 0 8px 16px rgba(0, 0, 0, 0.14))",
          },
        },
      }}
    >
      <DialogFooter
        styles={{
          actions: {
            marginTop: "var(--spacingVerticalL, 24px)",
          },
        }}
      >
        <PrimaryButton
          onClick={onConfirm}
          text="Delete"
          styles={{
            root: {
              backgroundColor: "var(--colorStatusDangerBackground1, #d13438)",
              borderColor: "var(--colorStatusDangerBackground1, #d13438)",
              color: "var(--colorNeutralForegroundOnBrand, #ffffff)",
            },
            rootHovered: {
              backgroundColor: "var(--colorStatusDangerForeground1, #b22528)",
              borderColor: "var(--colorStatusDangerForeground1, #b22528)",
            },
          }}
        />
        <DefaultButton
          onClick={onDismiss}
          text="Cancel"
          styles={{
            root: {
              borderColor: "var(--colorNeutralStroke1, #c7c7c7)",
              color: "var(--colorNeutralForeground1, #323130)",
            },
          }}
        />
      </DialogFooter>
    </Dialog>
  );
};
