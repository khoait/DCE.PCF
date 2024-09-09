import React from "react";
import { TimePickerControlProps } from "../types/typings";
import TimePickerControlNewLook from "./TimePickerControlNewLook";
import TimePickerControlClassic from "./TimePickerControlClassic";

export default function TimePickerControl(props: TimePickerControlProps) {
  const isNewLook = !!props.fluentDesign;
  return isNewLook ? <TimePickerControlNewLook {...props} /> : <TimePickerControlClassic {...props} />;
}
