import { FluentProvider, Theme } from "@fluentui/react-components";
import { TimePicker } from "@fluentui/react-timepicker-compat";
import React from "react";

export interface ITimePickerControlNewLook {
  theme: Theme;
}

export default function TimePickerControlNewLook({
  theme,
}: ITimePickerControlNewLook) {
  return (
    <FluentProvider theme={theme}>
      <TimePicker />
    </FluentProvider>
  );
}
