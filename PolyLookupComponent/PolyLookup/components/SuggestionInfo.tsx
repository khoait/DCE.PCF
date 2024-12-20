import { IconButton, ILabelStyles, IStyle, Stack, Text, TooltipHost } from "@fluentui/react";
import { Tooltip } from "@fluentui/react-components";
import { useBoolean } from "@fluentui/react-hooks";
import { Info16Regular } from "@fluentui/react-icons";
import React from "react";
import { EntityOption, ShowOptionDetailsEnum } from "../types/typings";

export interface ISuggestionInfoProps {
  data: EntityOption;
  columns: string[];
  showOptionDetails: ShowOptionDetailsEnum;
  isModern?: boolean;
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

export const SuggestionInfo = ({ data, columns, showOptionDetails, isModern }: ISuggestionInfoProps) => {
  const [showMore, { toggle: toggleshowMore }] = useBoolean(showOptionDetails === ShowOptionDetailsEnum.Expanded);

  const infoMap = new Map<string, string>();
  columns.forEach((column) => {
    const formattedValue1 = data.entity[`_${column}_value@OData.Community.Display.V1.FormattedValue`];
    const formattedValue2 = data.entity[`${column}@OData.Community.Display.V1.FormattedValue`];

    const displayValue = formattedValue1 ?? formattedValue2 ?? data.entity[column] ?? "";

    infoMap.set(column, displayValue);
  });

  const tooltipContent = Array.from(infoMap)
    .filter(([key, value], index) => value !== "" && index > 0)
    .map(([key, value]) => <div key={`tooltip_${key}`}>{value}</div>);

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
          {showOptionDetails === ShowOptionDetailsEnum.Tooltip ? (
            <Stack verticalAlign="center" horizontalAlign="center" verticalFill>
              {isModern ? (
                <Tooltip
                  content={{
                    children: tooltipContent,
                  }}
                  positioning="above-start"
                  withArrow
                  relationship="description"
                >
                  <Info16Regular tabIndex={0} />
                </Tooltip>
              ) : (
                <TooltipHost content={tooltipContent}>
                  <Info16Regular tabIndex={0} style={{ padding: "0 4px" }} />
                </TooltipHost>
              )}
            </Stack>
          ) : (
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
          )}
        </Stack.Item>
      ) : null}
    </Stack>
  );
};
