import {
  Dropdown,
  DropdownProps,
  FluentProvider,
  Image,
  Menu,
  MenuButton,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Option,
  OptionGroup,
  Tooltip,
  makeStyles,
  mergeClasses,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import { AddRegular, MoreVerticalRegular, OpenRegular, SearchRegular } from "@fluentui/react-icons";
import { useQueryClient } from "@tanstack/react-query";
import React, { MouseEventHandler, useCallback, useEffect, useState } from "react";
import { getCustomFilterString, getHandlebarsVariables } from "../services/TemplateService";
import { EntityOption, IconSizes, LookdownControlProps, OpenRecordMode, ShowIconOptions } from "../types/typings";
import { useAttributeOnChange } from "../hooks/useAttributeOnChange";
import { useEntityOptions } from "../hooks/queries/useEntityOptions";
import { useLanguagePack } from "../hooks/queries/useLanguagePack";
import { useMetadata } from "../hooks/queries/useMetadata";
import { getSelectedItemQueryOptions, useSelectedItem } from "../hooks/queries/useSelectedItem";

const useStyle = makeStyles({
  root: {
    display: "flex",
    flexDirection: "row",
    flexWrap: "nowrap",
    alignItems: "center",
    ...shorthands.gap(tokens.spacingHorizontalXS),
  },
  dropdown: {
    width: "100%",
    minWidth: "0px",
  },
  optionLayout: {
    display: "flex",
    flexDirection: "row",
    flexWrap: "nowrap",
    alignItems: "center",
    ...shorthands.gap(tokens.spacingHorizontalS),
  },
  flexFill: {
    flexGrow: 1,
    flexShrink: 1,
    flexBasis: 0,
    minWidth: 0,
  },
  breakWord: {
    wordBreak: "break-word",
  },
  disabled: {
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderColor(tokens.colorTransparentStroke),
    "&>button": {
      color: tokens.colorNeutralForeground1,
      cursor: "default",
    },
  },
  hidden: {
    display: "none",
  },
});

export default function LookdownControlNewLook({
  lookupViewId,
  lookupEntity,
  selectedId,
  selectedText,
  customFilter,
  groupBy,
  optionTemplate,
  selectedItemTemplate,
  showIcon,
  iconSize,
  openRecordMode,
  allowQuickCreate,
  allowLookupPanel,
  disabled,
  defaultLanguagePack,
  languagePackPath,
  fluentDesign,
  onChange,
}: LookdownControlProps) {
  const queryClient = useQueryClient();
  const styles = useStyle();

  const [selectedDisplayText, setSelectedDisplayText] = useState<string>(selectedText ?? "");
  const [selectedValues, setSelectedValues] = useState<string[]>(selectedId ? [selectedId] : []);

  const { data: loadedLanguagePack, isError: isErrorLanguagePack } = useLanguagePack(
    languagePackPath,
    defaultLanguagePack
  );

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const { data: metadata, isError: isErrorMetadata } = useMetadata(lookupEntity ?? "", lookupViewId ?? "");

  const { data: selectedOption, isError: isErrorSelectedOption } = useSelectedItem(
    metadata,
    selectedId,
    selectedItemTemplate,
    showIcon,
    iconSize
  );

  const { data: entityOptions, isError: isErrorEntityOptions } = useEntityOptions(
    metadata,
    customFilter,
    groupBy,
    optionTemplate,
    showIcon,
    iconSize
  );

  useAttributeOnChange(metadata, customFilter);

  const isError = isErrorLanguagePack || isErrorMetadata || isErrorSelectedOption || isErrorEntityOptions;

  const dropdownStyles = mergeClasses(styles.dropdown, styles.flexFill, (disabled || isError) && styles.disabled);

  const handleOptionSelect: DropdownProps["onOptionSelect"] = (ev, data) => {
    setSelectedValues(data.selectedOptions);
    setSelectedDisplayText(data.optionText ?? "");

    const selectedOptionId = data.selectedOptions.at(0);
    if (!selectedOptionId) {
      onChange?.(null);
      return;
    }

    if (metadata && entityOptions) {
      const selectionItem = entityOptions.find((item) => item.id === selectedOptionId);

      let selectedValue: ComponentFramework.LookupValue | null = null;

      if (selectionItem) {
        selectedValue = {
          entityType: metadata.lookupEntity.LogicalName,
          id: selectionItem.id,
          name: selectionItem.primaryName,
        } as ComponentFramework.LookupValue;

        setSelectedValues([selectionItem.id]);
        setSelectedDisplayText(selectionItem?.optionText);
      } else {
        setSelectedValues([]);
        setSelectedDisplayText("");
      }
      onChange?.(selectedValue);
    }
  };

  const handleOpenRecordMenuClick = () => {
    if (!openRecordMode) return;
    if (!lookupEntity) return;
    if (!selectedId) return;

    Xrm.Navigation.navigateTo(
      {
        pageType: "entityrecord",
        entityName: lookupEntity,
        entityId: selectedId,
      },
      { target: openRecordMode === OpenRecordMode.Dialog ? 2 : 1 }
    );
  };

  const handleQuickCreateMenuClick = async () => {
    if (!allowQuickCreate) return;
    if (!lookupEntity) return;
    if (!metadata) return;

    let createdRecord: ComponentFramework.LookupValue | null = null;
    if (metadata.lookupEntity.IsQuickCreateEnabled) {
      const result = await Xrm.Navigation.openForm({
        entityName: lookupEntity,
        useQuickCreateForm: true,
      });
      if (!result.savedEntityReference?.length) return;

      createdRecord = result.savedEntityReference[0];
    } else {
      const result = await Xrm.Navigation.navigateTo(
        { pageType: "entityrecord", entityName: lookupEntity },
        { target: 2 }
      );

      if (!result?.savedEntityReference?.length) return;

      createdRecord = result.savedEntityReference[0];
    }

    if (!createdRecord) return;

    onChange?.(createdRecord);

    queryClient.invalidateQueries({ queryKey: ["entityRecords"] });
  };

  const handleLookupPanelMenuClick = async () => {
    if (!allowLookupPanel) return;
    if (!lookupEntity) return;
    const result = await Xrm.Utility.lookupObjects({
      entityTypes: [lookupEntity],
      defaultEntityType: lookupEntity,
      allowMultiSelect: false,
      disableMru: true,
      filters: customFilter?.length
        ? [{ entityLogicalName: lookupEntity, filterXml: getCustomFilterString(customFilter) }]
        : [],
    });

    if (!result) return;

    const selectedItem = {
      entityType: lookupEntity,
      id: result[0].id,
      name: result[0].name,
    } as ComponentFramework.LookupValue;

    onChange?.(selectedItem);
  };

  const hasActions = openRecordMode || allowQuickCreate || allowLookupPanel;

  const renderOptions = () => {
    if (!entityOptions) return null;

    const entityIcon =
      metadata?.lookupEntity.IconVectorName ??
      (iconSize === IconSizes.Large
        ? metadata?.lookupEntity.IconMediumName ?? metadata?.lookupEntity.IconSmallName
        : metadata?.lookupEntity.IconSmallName);

    let iconTemplate = "";
    if (showIcon === ShowIconOptions.RecordImage) {
      iconTemplate = metadata?.lookupEntity.RecordImageUrlTemplate ?? "";
    } else if (showIcon === ShowIconOptions.EntityIcon) {
      iconTemplate = entityIcon ?? "";
    }

    if (!groupBy) {
      return entityOptions.map((item) => {
        return (
          <Option key={item.id} value={item.id} text={item.optionText}>
            <OptionDisplay
              optionText={item.optionText}
              iconSrc={item.iconSrc}
              iconSize={item.iconSize}
              onMouseEnter={() =>
                queryClient.prefetchQuery(
                  getSelectedItemQueryOptions(
                    metadata?.lookupEntity.EntitySetName ?? "",
                    item.id,
                    metadata?.lookupEntity.PrimaryIdAttribute ?? "",
                    metadata?.lookupEntity.PrimaryNameAttribute ?? "",
                    selectedItemTemplate,
                    showIcon,
                    iconTemplate,
                    iconSize
                  )
                )
              }
            />
          </Option>
        );
      });
    }

    // group results by groupBy field
    const grouped: Record<string, EntityOption[]> = {};
    entityOptions.forEach((item) => {
      const key = item.group ?? languagePack.BlankValueLabel ?? "";

      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(item);
    });

    return Object.keys(grouped).map((key) => {
      return (
        <OptionGroup key={key} label={key}>
          {grouped[key].map((item) => {
            return (
              <Option key={item.id} value={item.id} text={item.optionText}>
                <OptionDisplay
                  optionText={item.optionText}
                  iconSrc={item.iconSrc}
                  iconSize={item.iconSize}
                  onMouseEnter={() =>
                    queryClient.prefetchQuery(
                      getSelectedItemQueryOptions(
                        metadata?.lookupEntity.EntitySetName ?? "",
                        item.id,
                        metadata?.lookupEntity.PrimaryIdAttribute ?? "",
                        metadata?.lookupEntity.PrimaryNameAttribute ?? "",
                        selectedItemTemplate,
                        showIcon,
                        iconTemplate,
                        iconSize
                      )
                    )
                  }
                />
              </Option>
            );
          })}
        </OptionGroup>
      );
    });
  };

  return (
    <FluentProvider style={{ width: "100%" }} theme={fluentDesign?.tokenTheme}>
      <div className={styles.root}>
        <Dropdown
          appearance="filled-darker"
          className={dropdownStyles}
          clearable={!disabled && !isError}
          disabled={disabled || isError}
          placeholder={isError ? languagePack.LoadDataErrorMessage : "---"}
          value={isError ? "" : selectedDisplayText}
          selectedOptions={isError ? [] : selectedValues}
          button={
            <OptionDisplay
              optionText={selectedOption?.optionText ?? ""}
              iconSrc={selectedOption?.iconSrc}
              iconSize={selectedOption?.iconSize}
            />
          }
          expandIcon={disabled || isError ? { className: styles.hidden } : undefined}
          onOptionSelect={handleOptionSelect}
        >
          {renderOptions()}
        </Dropdown>
        {hasActions ? (
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Tooltip content="more actions" relationship="label">
                <MenuButton appearance="subtle" icon={<MoreVerticalRegular />} />
              </Tooltip>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                {openRecordMode ? (
                  <MenuItem icon={<OpenRegular />} onClick={handleOpenRecordMenuClick}>
                    {languagePack.OpenRecordLabel}
                  </MenuItem>
                ) : null}
                {allowQuickCreate && !disabled && !isError ? (
                  <MenuItem icon={<AddRegular />} onClick={handleQuickCreateMenuClick}>
                    {languagePack.QuickCreateLabel}
                  </MenuItem>
                ) : null}
                {allowLookupPanel && !disabled && !isError ? (
                  <MenuItem icon={<SearchRegular />} onClick={handleLookupPanelMenuClick}>
                    {languagePack.LookupPanelLabel}
                  </MenuItem>
                ) : null}
              </MenuList>
            </MenuPopover>
          </Menu>
        ) : null}
      </div>
    </FluentProvider>
  );
}

interface OptionDisplayProps {
  optionText: string;
  iconSrc?: string;
  iconSize?: number;
  onMouseEnter?: MouseEventHandler<HTMLDivElement> | undefined;
}

function OptionDisplay({ optionText, iconSrc, iconSize, onMouseEnter }: OptionDisplayProps) {
  const styles = useStyle();
  const optionTextStyles = mergeClasses(styles.flexFill, styles.breakWord);

  return (
    <div className={styles.optionLayout} onMouseEnter={onMouseEnter}>
      {iconSrc && iconSize ? (
        <div style={{ width: iconSize, height: iconSize }}>
          <Image fit="cover" src={iconSrc}></Image>
        </div>
      ) : null}

      <div className={optionTextStyles}>{optionText}</div>
    </div>
  );
}
