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
        <DefaultButton
          onClick={onDismiss}
          text="OK"
          styles={{
            root: {
              backgroundColor: "var(--colorBrandBackground, #0078d4)",
              borderColor: "var(--colorBrandBackground, #0078d4)",
              color: "var(--colorNeutralForegroundOnBrand, #ffffff)",
            },
            rootHovered: {
              backgroundColor: "var(--colorBrandBackgroundHover, #106ebe)",
              borderColor: "var(--colorBrandBackgroundHover, #106ebe)",
            },
          }}
        />
      </DialogFooter>
    </Dialog>
  );
};
