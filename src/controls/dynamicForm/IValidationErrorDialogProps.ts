import { DialogType, IDialogStyles } from "@fluentui/react";
import React from "react";

export interface IValidationErrorDialogProps {
  /**
   * Specifies if a dialog should be shown when validation fails. Default - false
   */
  showDialogOnValidationError?: boolean;
  /**
   * Specifies a custom title to be shown in the validation dialog. Default - empty
   */
  customTitle?: string;
  /**
   * Specifies a custom message to be shown in the validation dialog. Default - empty
   */
  customMessage?: string;

  /**
   * Specifies the dialog type. Default - DialogType.normal
   */
  dialogType?: DialogType;

  /**
   * Specifies custom styles for the validation error dialog
   */
  dialogStyles?: Partial<IDialogStyles>;
}
