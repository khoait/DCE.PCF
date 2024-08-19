import {
  Avatar,
  FluentProvider,
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerProps,
} from "@fluentui/react-components";
import React from "react";
import { PolyLookupProps } from "../types/typings";

const options = [
  "John Doe",
  "Jane Doe",
  "Max Mustermann",
  "Erika Mustermann",
  "Pierre Dupont",
  "Amelie Dupont",
  "Mario Rossi",
  "Maria Rossi",
];

export default function PolyLookupControlNewLook(props: PolyLookupProps) {
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const onOptionSelect: TagPickerProps["onOptionSelect"] = (e, data) => {
    setSelectedOptions(data.selectedOptions);
  };
  const tagPickerOptions = options.filter((option) => !selectedOptions.includes(option));

  return (
    <FluentProvider style={{ width: "100%" }} theme={props.fluentDesign?.tokenTheme}>
      <TagPicker onOptionSelect={onOptionSelect} selectedOptions={selectedOptions}>
        <TagPickerControl>
          <TagPickerGroup>
            {selectedOptions.map((option) => (
              <Tag
                key={option}
                shape="rounded"
                media={<Avatar aria-hidden name={option} color="colorful" />}
                value={option}
              >
                {option}
              </Tag>
            ))}
          </TagPickerGroup>
          <TagPickerInput aria-label="Select Employees" />
        </TagPickerControl>
        <TagPickerList>
          {tagPickerOptions.length > 0
            ? tagPickerOptions.map((option) => (
                <TagPickerOption
                  secondaryContent="Microsoft FTE"
                  media={<Avatar shape="square" aria-hidden name={option} color="colorful" />}
                  value={option}
                  key={option}
                >
                  {option}
                </TagPickerOption>
              ))
            : "No options available"}
        </TagPickerList>
      </TagPicker>
    </FluentProvider>
  );
}
