import {
  FluentProvider,
  Theme,
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import { Clock16Regular } from "@fluentui/react-icons";
import {
  TimePicker,
  TimePickerProps,
  formatDateToTimeString,
} from "@fluentui/react-timepicker-compat";
import React, { useState } from "react";
import { TimePickerControlProps } from "../types/typings";

const useStyles = makeStyles({
  root: {
    width: "100%",
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderColor(tokens.colorTransparentStroke),
    ":hover": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":active": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":focus": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":focus-within": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
  },
});

export default function TimePickerControlNewLook({
  theme,
  inputValue,
  dateAnchor,
  disabled,
  placeholder,
  freeform,
  hourCycle12,
  increment,
  startHour,
  endHour,
  onTimeChange,
}: { theme: Theme } & TimePickerControlProps) {
  const styles = useStyles();

  const [selectedTime, setSelectedTime] = useState<Date | null>(
    inputValue ?? null
  );
  const [value, setValue] = useState<string>(
    selectedTime
      ? formatDateToTimeString(selectedTime, {
          hourCycle: hourCycle12 ? "h12" : "h23",
          showSeconds: false,
        })
      : ""
  );

  const handleTimeChange: TimePickerProps["onTimeChange"] = (_ev, data) => {
    setSelectedTime(data.selectedTime);
    setValue(
      data.selectedTime && data.selectedTimeText ? data.selectedTimeText : ""
    );
    if (onTimeChange) {
      onTimeChange(data.selectedTime);
    }
  };

  const handleOnInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
    setValue(ev.target.value);
  };

  let start: number | undefined;
  let end: number | undefined;
  if (startHour !== undefined) {
    start = startHour;
    if (startHour > 23) {
      start = 23;
    } else if (startHour < 0) {
      start = 0;
    }
  }

  if (endHour !== undefined) {
    end = endHour;
    if (endHour > 24) {
      end = 24;
    } else if (endHour < 1) {
      end = 1;
    }
  }

  return (
    <FluentProvider theme={theme}>
      <TimePicker
        className={styles.root}
        selectedTime={selectedTime}
        value={value}
        dateAnchor={dateAnchor ?? undefined}
        disabled={disabled}
        placeholder={placeholder}
        increment={increment}
        freeform={freeform ?? false}
        hourCycle={hourCycle12 ? "h12" : "h23"}
        startHour={start as any}
        endHour={end as any}
        onTimeChange={handleTimeChange}
        onInput={freeform ? handleOnInput : undefined}
        expandIcon={<Clock16Regular />}
      />
    </FluentProvider>
  );
}
