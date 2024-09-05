import React from "react";
import { IconButton, ILabelStyles, IStyle, Label, Stack, Text } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { EntityOption } from "../types/typings";

export interface ISuggestionInfoProps {
  data: EntityOption;
  columns: string[];
}

const commonStyle: IStyle = {
  display: "block",
  fontWeight: 400,
  padding: 0,
  textAlign: "left",
  overflow: "hidden",
  textOverflow: "ellipsis",
  whiteSpace: "nowrap",
  width: "100%",
};
const primaryStyle: Partial<ILabelStyles> = { root: commonStyle };
const secondaryStyle: Partial<ILabelStyles> = {
  root: { ...commonStyle, color: "#666" },
};

export const SuggestionInfo = ({ data, columns }: ISuggestionInfoProps) => {
  const [showMore, { toggle: toggleshowMore }] = useBoolean(false);

  const infoMap = new Map<string, string>();
  columns.forEach((column) => {
    let displayValue = data.entity[column + "@OData.Community.Display.V1.FormattedValue"];
    if (!displayValue) {
      displayValue = data.entity[column];
    }
    infoMap.set(column, displayValue ?? "");
  });

  return (
    <Stack horizontal grow styles={{ root: { flex: 1, minWidth: 0 } }}>
      <Stack.Item grow align="stretch" styles={{ root: { minWidth: "0", padding: 10 } }}>
        {Array.from(infoMap).map(([key, value], index) => {
          if (value === "") return null;
          if (!showMore && index > 0) return null;
          return (
            <div key={key}>
              <Text styles={index === 0 ? primaryStyle : secondaryStyle}>{value}</Text>
            </div>
          );
        })}
      </Stack.Item>
      {columns.length > 1 ? (
        <Stack.Item align="stretch">
          <IconButton
            iconProps={{ iconName: showMore ? "ChevronUp" : "ChevronDown" }}
            styles={{
              root: { height: "100%", color: "#000" },
              flexContainer: { alignItems: showMore ? "flex-start" : "center", paddingTop: showMore ? 11 : 0 },
            }}
            title="More details"
            onClick={(e) => {
              e.preventDefault();
              e.stopPropagation();
              toggleshowMore();
              return false;
            }}
          />
        </Stack.Item>
      ) : null}
    </Stack>
  );
};
