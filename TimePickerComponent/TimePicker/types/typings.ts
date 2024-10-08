export interface TimePickerControlProps {
  inputValue?: Date | null;
  dateAnchor?: Date | null;
  disabled?: boolean;
  placeholder?: string;
  freeform?: boolean;
  hourCycle12?: boolean;
  increment?: number;
  startHour?: number;
  endHour?: number;
  fluentDesign?: ComponentFramework.FluentDesignState;
  onTimeChange?: (selectedTime: Date | null) => void;
}
