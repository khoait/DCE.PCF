import { IComboBoxStyles, IStyle, TimePicker } from "@fluentui/react";
import React from "react";
import { TimePickerControlProps } from "../types/typings";

const BORDER_HIGHLIGHT: IStyle = {
  borderWidth: "1px !important",
  borderRadius: "0px !important",
  borderColor: "#666",
};

const BORDER_TRANSPARENT: IStyle = {
  borderWidth: "1px !important",
  borderRadius: "0px !important",
  borderColor: "transparent",
};

const DEFAULT_BORDER_STYLES: IStyle = {
  ...BORDER_TRANSPARENT,
  "&:after": BORDER_TRANSPARENT,
  "&:hover": BORDER_HIGHLIGHT,
  "&:focus": BORDER_HIGHLIGHT,
  "&:hover:after": BORDER_HIGHLIGHT,
  "&:focus:after": BORDER_HIGHLIGHT,
};

const timePickerStyes: Partial<IComboBoxStyles> = {
  root: {
    width: "100%",
    ...DEFAULT_BORDER_STYLES,
    "&:active>.ms-ComboBox-Input": {
      fontWeight: 400,
    },
    "&:focus>.ms-ComboBox-Input": {
      fontWeight: 400,
    },
    "&:hover>.ms-ComboBox-Input": {
      fontWeight: 400,
    },
  },
  rootPressed: {
    "&:after": {
      borderRadius: "0px !important",
    },
  },
  rootDisallowFreeForm: DEFAULT_BORDER_STYLES,
  rootDisabled: {
    backgroundColor: "transparent",
    "&:hover": {
      backgroundColor: "#e2e2e2",
    },
    "&:after": {
      borderColor: "transparent",
    },
    "&:hover:after": {
      borderColor: "transparent",
      backgroundColor: "transparent",
    },
    "&>.ms-Button": {
      display: "none",
    },
  },
  inputDisabled: {
    backgroundColor: "transparent",
    color: "black",
    fontWeight: 600,
    "&:hover": {
      backgroundColor: "#e2e2e2",
    },
  },
  input: {
    color: "black",
    fontWeight: 600,
    "&:active": {
      fontWeight: 400,
    },
    "&:focus": {
      fontWeight: 400,
    },
    "&:hover": {
      fontWeight: 400,
    },
  },
  optionsContainerWrapper: {
    maxHeight: "500px",
  },
};

export default function TimePickerControl({
  disabled,
}: TimePickerControlProps) {
  const now = new Date();
  return <TimePicker defaultValue={now} styles={timePickerStyes} />;
}
