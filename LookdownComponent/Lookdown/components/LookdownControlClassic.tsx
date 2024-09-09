import {
  Dropdown,
  DropdownMenuItemType,
  IContextualMenuItemStyles,
  IContextualMenuProps,
  IDropdownOption,
  IIconProps,
  ILabelStyles,
  IStyle,
  IconButton,
  Image,
  ImageFit,
  Label,
  Stack,
} from "@fluentui/react";
import { useQueryClient } from "@tanstack/react-query";
import React, { useCallback, useEffect } from "react";
import { useEntityOptions, useLanguagePack, useMetadata } from "../services/DataverseService";
import { getClassicDropdownOptions } from "../services/DropdownHelper";
import { getCustomFilterString, getHandlebarsVariables } from "../services/TemplateService";
import { EntityOption, LookdownControlProps, OpenRecordMode } from "../types/typings";

const DEFAULT_BORDER_STYLES: IStyle = {
  borderColor: "#666",
  borderWidth: 1,
  borderRadius: 0,
};

const groupLabelStyles: ILabelStyles = {
  root: {
    fontWeight: 600,
    color: "#0078d4",
    textOverflow: "ellipsis",
    overflow: "hidden",
    whiteSpace: "nowrap",
  },
};

const optionLabelStyles: ILabelStyles = {
  root: {
    fontWeight: 400,
  },
};

const selectionOptionLabelStyles: ILabelStyles = {
  root: {
    flex: 1,
    padding: 0,
    fontWeight: 600,
    textOverflow: "ellipsis",
    overflow: "hidden",
    whiteSpace: "nowrap",
    "&:active": {
      fontWeight: 400,
    },
    "&:hover": {
      fontWeight: 400,
    },
  },
};

const commandIcon: IIconProps = { iconName: "MoreVertical" };

const hiddenCommandStyle: IContextualMenuItemStyles = {
  root: {
    display: "none",
  },
};

const visibleCommandStyle: IContextualMenuItemStyles = {
  root: {
    display: "block",
  },
};

export default function LookdownControlClassic({
  lookupViewId,
  lookupEntity,
  selectedId,
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
  onChange,
}: LookdownControlProps) {
  const queryClient = useQueryClient();

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

  const onRenderOption = (
    option?: IDropdownOption<EntityOption>,
    defaultRenderer?: (option?: IDropdownOption<EntityOption>) => JSX.Element | null
  ) => {
    if (!option || !showIcon) {
      return defaultRenderer ? defaultRenderer(option) : null;
    }

    const isGroup = option.itemType === DropdownMenuItemType.Header;

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        {option.data?.iconSrc && !isGroup ? (
          <Image
            imageFit={ImageFit.cover}
            src={option.data.iconSrc}
            width={option.data.iconSize}
            height={option.data.iconSize}
          ></Image>
        ) : null}
        <Stack.Item
          grow
          styles={{
            root: { flex: 1, textOverflow: "ellipsis", overflow: "hidden", whiteSpace: isGroup ? "nowrap" : "normal" },
          }}
        >
          <Label styles={isGroup ? groupLabelStyles : optionLabelStyles}>{option.text}</Label>
        </Stack.Item>
      </Stack>
    );
  };

  const onRenderSelectedOption = (
    options?: IDropdownOption<EntityOption>[],
    defaultRenderer?: (options?: IDropdownOption<EntityOption>[]) => JSX.Element | null
  ) => {
    const option = options?.at(0);

    if (!option || !showIcon) {
      return defaultRenderer ? defaultRenderer(options) : null;
    }

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        {option.data?.iconSrc ? (
          <Image
            imageFit={ImageFit.cover}
            src={option.data.iconSrc}
            width={option.data.iconSize}
            height={option.data.iconSize}
          ></Image>
        ) : null}
        <Label styles={selectionOptionLabelStyles}>{option.data?.selectedOptionText}</Label>
      </Stack>
    );
  };

  const hasCommand = openRecordMode || allowQuickCreate || allowLookupPanel;

  const onOpenRecordCommandClick = () => {
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

  const onQuickCreateCommandClick = async () => {
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

  const onLookupPanelCommandClick = async () => {
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

  const menuProps: IContextualMenuProps = {
    items: [
      {
        key: "open-record",
        text: languagePack.OpenRecordLabel,
        iconProps: { iconName: "OpenInNewWindow" },
        itemProps: { styles: !openRecordMode ? hiddenCommandStyle : visibleCommandStyle },
        disabled: !selectedId,
        onClick: onOpenRecordCommandClick,
      },
      {
        key: "quick-create",
        text: languagePack.QuickCreateLabel,
        itemProps: { styles: !allowQuickCreate ? hiddenCommandStyle : visibleCommandStyle },
        iconProps: { iconName: "Add" },
        disabled: !allowQuickCreate,
        onClick: onQuickCreateCommandClick,
      },
      {
        key: "lookup-panel",
        text: languagePack.LookupPanelLabel,
        itemProps: { styles: !allowLookupPanel ? hiddenCommandStyle : visibleCommandStyle },
        iconProps: { iconName: "LookupEntities" },
        disabled: !allowLookupPanel,
        onClick: onLookupPanelCommandClick,
      },
    ],
  };

  const invalidateFetchDataFn = useCallback((context) => {
    queryClient.invalidateQueries({ queryKey: ["entityRecords"] });
  }, []);

  useEffect(() => {
    const templateVar = getHandlebarsVariables((metadata?.lookupView.fetchxml ?? "") + (customFilter ?? ""));
    if (templateVar.length) {
      templateVar.forEach((v) => {
        Xrm.Page.data.entity.attributes.get(v)?.addOnChange(invalidateFetchDataFn);
      });
    }
  }, [metadata?.lookupView.fetchxml, customFilter]);

  return (
    <Stack horizontal styles={{ root: { width: "100%" } }}>
      <Stack.Item
        grow
        styles={{
          root: {
            overflow: "hidden",
            textOverflow: "ellipsis",
          },
        }}
      >
        <Dropdown
          selectedKey={isError ? null : selectedId ?? null}
          placeholder={isError ? languagePack.LoadDataErrorMessage : "---"}
          options={getClassicDropdownOptions(entityOptions ?? [], groupBy ?? null, languagePack)}
          disabled={isError ? true : disabled}
          onChange={(event, option) => {
            if (!onChange) return;

            if (!option) {
              onChange(null);
              return;
            }

            const data = option.data as EntityOption;
            const selectedItem = {
              entityType: lookupEntity,
              id: data.id,
              name: data.primaryName,
            } as ComponentFramework.LookupValue;
            onChange(selectedItem);
          }}
          onRenderOption={showIcon ? onRenderOption : undefined}
          onRenderTitle={showIcon ? onRenderSelectedOption : undefined}
          styles={(styleProps) => {
            return {
              root: {
                width: "100%",
              },
              title: {
                borderColor: styleProps.isOpen ? "#666" : "transparent",
                borderRadius: styleProps.isOpen ? 1 : 0,
                fontWeight: styleProps.isOpen ? 400 : 600,
                color: "#000",
                "&:active": {
                  fontWeight: 400,
                },
                "&:hover": {
                  fontWeight: 400,
                },
              },
              dropdown: {
                "&:active": DEFAULT_BORDER_STYLES,
                "&:focus": DEFAULT_BORDER_STYLES,
                "&:focus:after": DEFAULT_BORDER_STYLES,
                "&:hover .ms-Dropdown-caretDownWrapper": {
                  display: "block",
                },
                "&:focus .ms-Dropdown-caretDownWrapper": {
                  display: "block",
                },
                "&:active .ms-Dropdown-title": {
                  fontWeight: 400,
                },
                "&:focus .ms-Dropdown-title": {
                  fontWeight: 400,
                },
              },
              dropdownItem: {
                height: "auto",
              },
              dropdownItemSelected: {
                height: "auto",
              },
              dropdownOptionText: {
                whiteSpace: "normal",
              },
              caretDownWrapper: {
                display: styleProps.isOpen ? "block" : "none",
              },
              callout: {
                maxHeight: 500,
                "& .ms-Callout-main": {
                  maxHeight: "500px !important",
                },
              },
            };
          }}
        />
      </Stack.Item>
      <Stack.Item>
        {hasCommand ? <IconButton iconProps={commandIcon} menuProps={menuProps} onRenderMenuIcon={() => null} /> : null}
      </Stack.Item>
    </Stack>
  );
}
