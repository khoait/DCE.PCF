import {
  FluentProvider,
  Theme,
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import {
  TimePicker,
  formatDateToTimeString,
} from "@fluentui/react-timepicker-compat";
import React from "react";

export interface ITimePickerControlNewLook {
  theme: Theme;
}

const useStyles = makeStyles({
  root: {
    width: "100%",
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderColor(tokens.colorTransparentStroke),
    ":hover": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
      "&>.fui-TimePicker__expandIcon": {
        visibility: "visible",
      },
    },
    ":active": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
      "&>.fui-TimePicker__expandIcon": {
        visibility: "visible",
      },
    },
    ":focus": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
      "&>.fui-TimePicker__expandIcon": {
        visibility: "visible",
      },
    },
    ":focus-within": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
      "&>.fui-TimePicker__expandIcon": {
        visibility: "visible",
      },
    },
    "&>.fui-TimePicker__expandIcon": {
      visibility: "hidden",
    },
  },
});

export default function TimePickerControlNewLook({
  theme,
}: ITimePickerControlNewLook) {
  const styles = useStyles();

  const now = new Date();
  theme.colorNeutralBackground3;
  return (
    <FluentProvider theme={theme}>
      <TimePicker
        className={styles.root}
        defaultSelectedTime={now}
        defaultValue={formatDateToTimeString(now)}
        placeholder=""
      />
    </FluentProvider>
  );
}
