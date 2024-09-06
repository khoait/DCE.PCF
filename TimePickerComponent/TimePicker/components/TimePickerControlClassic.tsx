import { IComboBoxStyles, IconButton, IStyle, ITimeRange, Stack, TimePicker } from "@fluentui/react";
import React, { useCallback, useState } from "react";
import { isValidDate } from "../services/dateService";
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
  container: {
    flex: 1,
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

export default function TimePickerControlClassic({
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
}: TimePickerControlProps) {
  const [selectedTime, setSelectedTime] = useState<Date | null>(inputValue ?? null);

  const handleOnChange = useCallback((_: any, newTime: Date) => {
    let newValue: Date | null = null;
    if (isValidDate(newTime)) {
      newValue = newTime;
    }
    setSelectedTime(newValue);

    onTimeChange?.(newValue);
  }, []);

  let range: ITimeRange | undefined = undefined;
  let start = 0;
  let end = 24;
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

  if (startHour !== undefined || endHour !== undefined) {
    range = {
      start: start > end ? end : start,
      end: end < start ? start : end,
    };
  }

  return (
    <Stack horizontal styles={{ root: { width: "100%" } }}>
      <TimePicker
        value={selectedTime ?? undefined}
        dateAnchor={dateAnchor ?? undefined}
        disabled={disabled}
        placeholder={placeholder}
        increments={increment}
        allowFreeform={freeform ?? false}
        useHour12={hourCycle12 ?? false}
        timeRange={range}
        onChange={handleOnChange}
        styles={timePickerStyes}
        buttonIconProps={{
          iconName: "Clock",
          styles: { root: { height: "16px", fontSize: "16px" } },
        }}
      />
      {selectedTime ? <IconButton iconProps={{ iconName: "Cancel" }} onClick={() => onTimeChange?.(null)} /> : null}
    </Stack>
  );
}
