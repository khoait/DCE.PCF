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
import React, { useCallback, useEffect, useState } from "react";
import { useEntityOptions, useLanguagePack, useMetadata } from "../services/DataverseService";
import { getCustomFilterString, getHandlebarsVariables } from "../services/TemplateService";
import { EntityOption, LookdownControlProps, OpenRecordMode } from "../types/typings";

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
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderColor(tokens.colorTransparentStroke),
    ":hover": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":active": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":focus": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
    ":focus-within": {
      ...shorthands.borderColor(tokens.colorTransparentStroke),
    },
  },
  optionLayout: {
    display: "flex",
    flexDirection: "row",
    flexWrap: "nowrap",
    alignItems: "center",
    ...shorthands.gap(tokens.spacingHorizontalS),
  },
  flexFill: {
    flex: 1,
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

  const dropdownStyles = mergeClasses(styles.dropdown, styles.flexFill);

  const [selectedDisplayText, setSelectedDisplayText] = useState<string>(selectedText ?? "");
  const [selectedValues, setSelectedValues] = useState<string[]>(selectedId ? [selectedId] : []);

  const { data: loadedLanguagePack, isError: isErrorLanguagePack } = useLanguagePack(
    languagePackPath,
    defaultLanguagePack
  );

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const { data: metadata, isError: isErrorMetadata } = useMetadata(lookupEntity ?? "", lookupViewId ?? "");

  const { data: entityOptions, isError: isErrorEntityOptions } = useEntityOptions(
    metadata,
    customFilter,
    groupBy,
    optionTemplate,
    selectedItemTemplate,
    showIcon,
    iconSize
  );

  const isError = isErrorLanguagePack || isErrorMetadata || isErrorEntityOptions;

  useEffect(() => {
    if (selectedId && entityOptions) {
      const selectionItem = entityOptions.find((item) => item.id === selectedId);

      if (selectionItem) {
        setSelectedValues([selectedId]);
        setSelectedDisplayText(selectionItem?.selectedOptionText);
        return;
      }
    }

    setSelectedValues([]);
    setSelectedDisplayText(selectedText ?? "");
  }, [selectedId, selectedText, entityOptions]);

  const invalidateFetchDataFn = useCallback((context) => {
    queryClient.invalidateQueries({ queryKey: ["entityRecords"] });
  }, []);

  useEffect(() => {
    const templateVar = getHandlebarsVariables(metadata?.lookupView.fetchxml ?? "" + customFilter ?? "");
    if (templateVar.length) {
      templateVar.forEach((v) => {
        Xrm.Page.data.entity.attributes.get(v)?.addOnChange(invalidateFetchDataFn);
      });
    }

    // TODO: not sure why this is not working
    // return () => {
    //   console.log("cleanup");
    //   templateVar.forEach((v) => {
    //     Xrm.Page.data.entity.attributes.get(v)?.removeOnChange(invalidateFetchDataFn);
    //   });
    // };
  }, [metadata?.lookupView.fetchxml, customFilter]);

  const getSelectedOptionDisplay = () => {
    const selectedOption = entityOptions?.find((item) => item.id === selectedValues.at(0) ?? "");
    if (!selectedOption) return <span>---</span>;
    return (
      <OptionDisplay
        optionText={selectedOption.optionText}
        iconSrc={selectedOption.iconSrc}
        iconSize={selectedOption.iconSize}
      />
    );
  };

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
        setSelectedDisplayText(selectionItem?.selectedOptionText);
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

    if (!groupBy) {
      return entityOptions.map((item) => {
        return (
          <Option key={item.id} value={item.id} text={item.selectedOptionText}>
            <OptionDisplay optionText={item.optionText} iconSrc={item.iconSrc} iconSize={item.iconSize} />
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
              <Option key={item.id} value={item.id} text={item.selectedOptionText}>
                <OptionDisplay optionText={item.optionText} iconSrc={item.iconSrc} iconSize={item.iconSize} />
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
          className={styles.dropdown}
          clearable
          placeholder={isError ? languagePack.LoadDataErrorMessage : "---"}
          value={isError ? "" : selectedDisplayText}
          selectedOptions={isError ? [] : selectedValues}
          button={getSelectedOptionDisplay()}
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
                {allowQuickCreate ? (
                  <MenuItem icon={<AddRegular />} onClick={handleQuickCreateMenuClick}>
                    {languagePack.QuickCreateLabel}
                  </MenuItem>
                ) : null}
                {allowLookupPanel ? (
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
}

function OptionDisplay({ optionText, iconSrc, iconSize }: OptionDisplayProps) {
  const styles = useStyle();

  return (
    <div className={styles.optionLayout}>
      {iconSrc && iconSize ? (
        <div style={{ width: iconSize, height: iconSize }}>
          <Image fit="cover" src={iconSrc}></Image>
        </div>
      ) : null}

      <div className={styles.flexFill}>{optionText}</div>
    </div>
  );
}
