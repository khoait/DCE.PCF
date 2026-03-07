import { FluentProvider, Input, Theme, makeStyles, shorthands, tokens } from "@fluentui/react-components";
import { Clock16Regular } from "@fluentui/react-icons";
import { TimePicker, TimePickerProps, formatDateToTimeString } from "@fluentui/react-timepicker-compat";
import React, { useEffect, useState } from "react";
import { TimePickerControlProps } from "../types/typings";

const useStyles = makeStyles({
  root: {
    width: "100%",
  },
});

export default function TimePickerControlNewLook({
  inputValue,
  dateAnchor,
  disabled,
  placeholder,
  freeform,
  hourCycle12,
  increment,
  startHour,
  endHour,
  fluentDesign,
  onTimeChange,
}: TimePickerControlProps) {
  const styles = useStyles();

  const [selectedTime, setSelectedTime] = useState<Date | null>(inputValue ?? null);
  const [value, setValue] = useState<string>(
    selectedTime
      ? formatDateToTimeString(selectedTime, {
          hourCycle: hourCycle12 ? "h12" : "h23",
          showSeconds: false,
        })
      : ""
  );

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

  const handleTimeChange: TimePickerProps["onTimeChange"] = (_ev, data) => {
    setSelectedTime(data.selectedTime);
    setValue(data.selectedTime && data.selectedTimeText ? data.selectedTimeText : "");
    onTimeChange?.(data.selectedTime);
  };

  const handleOnInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
    setValue(ev.target.value);
  };

  // trigger onChange when dateAnchor changes
  useEffect(() => {
    if (selectedTime && dateAnchor) {
      const newSelectedTime = new Date(selectedTime);
      newSelectedTime.setFullYear(dateAnchor.getFullYear(), dateAnchor.getMonth(), dateAnchor.getDate());
      setSelectedTime(newSelectedTime);
      onTimeChange?.(newSelectedTime);
    } else if (dateAnchor === null) {
      setSelectedTime(null);
      onTimeChange?.(null);
    }
  }, [dateAnchor]);

  const theme = fluentDesign?.tokenTheme;
  const currentTheme = disabled
    ? {
        ...theme,
        colorCompoundBrandStroke: theme?.colorNeutralStroke1,
        colorCompoundBrandStrokeHover: theme?.colorNeutralStroke1Hover,
        colorCompoundBrandStrokePressed: theme?.colorNeutralStroke1Pressed,
        colorCompoundBrandStrokeSelected: theme?.colorNeutralStroke1Selected,
      }
    : theme;

  return (
    <FluentProvider theme={currentTheme}>
      {disabled ? (
        <Input className={styles.root} appearance="filled-darker" readOnly value={value} />
      ) : (
        <TimePicker
          className={styles.root}
          appearance="filled-darker"
          clearable
          selectedTime={selectedTime}
          value={value}
          dateAnchor={dateAnchor ?? undefined}
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
      )}
    </FluentProvider>
  );
}
