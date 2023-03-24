import React from "react";
import { IconButton, ILabelStyles, IStyle, Label, Stack } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";

export interface ISuggestionInfoProps {
  infoMap: Map<string, string>;
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

export const SuggestionInfo = ({ infoMap }: ISuggestionInfoProps) => {
  const [showMore, { toggle: toggleshowMore }] = useBoolean(false);

  let displayValueCount = 0;

  infoMap.forEach((value) => {
    if (value !== "") displayValueCount++;
  });

  return (
    <Stack horizontal grow styles={{ root: { width: "100%" } }}>
      <Stack.Item grow align="stretch" styles={{ root: { minWidth: "0", padding: 10 } }}>
        {Array.from(infoMap).map(([key, value], index) => {
          if (value === "") return null;
          if (!showMore && index > 0) return null;
          return (
            <div key={key}>
              <Label styles={index === 0 ? primaryStyle : secondaryStyle}>{value}</Label>
            </div>
          );
        })}
      </Stack.Item>
      {displayValueCount > 1 ? (
        <Stack.Item align="stretch">
          <IconButton
            iconProps={{ iconName: showMore ? "ChevronUp" : "ChevronDown" }}
            styles={{ root: { height: "100%", color: "#000" } }}
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
