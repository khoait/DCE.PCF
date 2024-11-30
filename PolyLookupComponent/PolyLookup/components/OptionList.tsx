import { TagPickerOption, Avatar, Divider, Text, tokens, makeStyles } from "@fluentui/react-components";
import React from "react";
import { LanguagePack } from "../types/languagePack";
import { EntityOption, ShowIconOptions, ShowOptionDetailsEnum } from "../types/typings";
import { SuggestionInfo } from "./SuggestionInfo";

const useStyles = makeStyles({
  optionListFooter: {
    padding: tokens.spacingVerticalS,
  },
  listBox: {
    maxHeight: "50vh",
    overflowX: "hidden",
    overflowY: "auto",
    padding: tokens.spacingVerticalXS,
  },
  transparentBackground: {
    backgroundColor: tokens.colorTransparentBackground,
  },
});

export interface OptionListProps {
  options: EntityOption[];
  columns: string[];
  languagePack: LanguagePack;
  isLoading?: boolean;
  hasMoreRecords?: boolean;
  showIcon?: ShowIconOptions;
  entityIconUrl?: string;
  showOptionDetails?: ShowOptionDetailsEnum;
}

export function OptionList({
  options,
  columns,
  languagePack,
  isLoading,
  hasMoreRecords,
  showIcon,
  entityIconUrl,
  showOptionDetails,
}: OptionListProps) {
  const { optionListFooter, listBox, transparentBackground } = useStyles();

  if (!options?.length) {
    return (
      <div className={optionListFooter}>
        <Text>{isLoading ? languagePack.LoadingMessage : languagePack.EmptyListDefaultMessage}</Text>
      </div>
    );
  }

  return (
    <div>
      <div className={listBox}>
        {options?.map((option) => (
          <TagPickerOption
            media={
              showIcon ? (
                <Avatar
                  className={transparentBackground}
                  size={showIcon === ShowIconOptions.EntityIcon ? 16 : 28}
                  shape="square"
                  name={showIcon === ShowIconOptions.EntityIcon ? "" : option.optionText}
                  image={{
                    className: transparentBackground,
                    src: showIcon === ShowIconOptions.EntityIcon ? entityIconUrl : option.iconSrc,
                  }}
                  color={showIcon === ShowIconOptions.EntityIcon ? "neutral" : "colorful"}
                  aria-hidden
                />
              ) : undefined
            }
            key={option.id}
            value={option.id}
            text={option.optionText}
          >
            <SuggestionInfo
              data={option}
              columns={columns}
              showOptionDetails={showOptionDetails ?? ShowOptionDetailsEnum.Collapsed}
            />
          </TagPickerOption>
        ))}
      </div>
      <Divider />
      <div className={optionListFooter}>
        <Text>{hasMoreRecords ? languagePack.SuggestionListFullMessage : languagePack.NoMoreRecordsMessage}</Text>
      </div>
    </div>
  );
}
