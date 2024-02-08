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
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import Handlebars from "handlebars";
import React from "react";
import { getCurrentRecord, useFetchXmlData, useLanguagePack, useMetadata } from "../services/DataverseService";
import { LanguagePack } from "../types/languagePack";

const queryClient = new QueryClient();

export enum ShowIconOptions {
  None = 0,
  EntityIcon = 1,
  RecordImage = 2,
}

export enum IconSizes {
  Normal = 0,
  Large = 1,
}

export enum OpenRecordMode {
  None = 0,
  Inline = 1,
  Dialog = 2,
}

export interface LookdownProps {
  lookupViewId?: string | null;
  lookupEntity?: string | null;
  selectedId?: string | null;
  customFilter?: string | null;
  groupBy?: string | null;
  optionTemplate?: string | null; // example of optionTemplate: {{ fullname }} - {{ emailaddress1 }}
  selectedItemTemplate?: string | null; // example of selectedItemTemplate: {{ fullname }} - {{ emailaddress1 }}
  showIcon?: ShowIconOptions;
  iconSize?: IconSizes;
  openRecordMode?: OpenRecordMode;
  allowQuickCreate?: boolean;
  allowLookupPanel?: boolean;
  disabled?: boolean;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  onChange?: (selectedItem: ComponentFramework.LookupValue | null) => void;
}

const DEFAULT_BORDER_STYLES: IStyle = {
  borderColor: "#666",
  borderWidth: 1,
  borderRadius: 0,
};

const groupLabelStyles: ILabelStyles = {
  root: {
    fontWeight: 600,
    color: "#0078d4",
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

const Body = ({
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
}: LookdownProps) => {
  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(lookupEntity, lookupViewId, showIcon === ShowIconOptions.EntityIcon);

  const templateColumns: string[] = [];
  if (optionTemplate || selectedItemTemplate) {
    templateColumns.push(...getHandlebarsVariables(optionTemplate ?? "" + " " + selectedItemTemplate ?? ""));
  }

  const fetchXml = getFetchTemplateString(metadata?.lookupView.fetchxml ?? "", customFilter, templateColumns);

  const {
    data: fetchData,
    isLoading: isLoadingFetchData,
    isSuccess: isLoadingFetchDataSuccess,
  } = useFetchXmlData(metadata?.lookupEntity.EntitySetName, lookupViewId, fetchXml, customFilter, groupBy);

  const getDropdownOptions = (): IDropdownOption<ComponentFramework.WebApi.Entity>[] => {
    if (!metadata) return [];
    if (!fetchData) return [];

    let options: IDropdownOption<ComponentFramework.WebApi.Entity>[] = [];

    let optionTemplateCompiled: HandlebarsTemplateDelegate | null = null;

    if (optionTemplate) {
      optionTemplateCompiled = Handlebars.compile(optionTemplate);
    }

    // group results by groupBy field
    if (groupBy) {
      const grouped: Record<string, ComponentFramework.WebApi.Entity[]> = {};
      fetchData.forEach((item) => {
        const key = item[groupBy]?.toString() ?? languagePack.BlankValueLabel;

        if (!grouped[key]) grouped[key] = [];
        grouped[key].push(item);
      });

      Object.keys(grouped).forEach((key) => {
        options.push({
          key: key,
          text: key,
          itemType: DropdownMenuItemType.Header,
        });

        grouped[key].forEach((item) => {
          options.push({
            key: item[metadata.lookupEntity.PrimaryIdAttribute],
            text: optionTemplateCompiled
              ? optionTemplateCompiled(item)
              : item[metadata.lookupEntity.PrimaryNameAttribute],
            data: item,
          });
        });
      });
    } else {
      options = fetchData.map((item) => {
        return {
          key: item[metadata.lookupEntity.PrimaryIdAttribute],
          text: optionTemplateCompiled
            ? optionTemplateCompiled(item)
            : item[metadata.lookupEntity.PrimaryNameAttribute],
          data: item,
        };
      });
    }

    if (options.length === 0) {
      options.push({
        key: "no-records",
        text: languagePack.EmptyListMessage,
        disabled: true,
      });
    }

    return options;
  };

  const entityIcon =
    metadata?.lookupEntity.IconVectorName ??
    (iconSize === IconSizes.Large
      ? metadata?.lookupEntity.IconMediumName ?? metadata?.lookupEntity.IconSmallName
      : metadata?.lookupEntity.IconSmallName);

  const recordImageTemplate =
    showIcon === ShowIconOptions.RecordImage
      ? Handlebars.compile(metadata?.lookupEntity.RecordImageUrlTemplate ?? "")
      : null;

  const imgSize = iconSize === IconSizes.Large ? 32 : 16;

  const onRenderOption = (
    option?: IDropdownOption<ComponentFramework.WebApi.Entity>,
    defaultRenderer?: (option?: IDropdownOption<ComponentFramework.WebApi.Entity>) => JSX.Element | null
  ) => {
    if (!option || !showIcon) {
      return defaultRenderer ? defaultRenderer(option) : null;
    }

    let imgSrc;

    if (showIcon === ShowIconOptions.EntityIcon) {
      imgSrc = entityIcon;
    } else if (showIcon === ShowIconOptions.RecordImage) {
      if (metadata && metadata.lookupEntity.RecordImageUrlTemplate && option.data && recordImageTemplate) {
        imgSrc = recordImageTemplate(option.data);
      }
    }

    const isGroup = option.itemType === DropdownMenuItemType.Header;

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        {imgSrc && !isGroup ? (
          <Image imageFit={ImageFit.cover} src={imgSrc} width={imgSize} height={imgSize}></Image>
        ) : null}
        <Label styles={isGroup ? groupLabelStyles : optionLabelStyles}>{option.text}</Label>
      </Stack>
    );
  };

  const onRenderSelectedOption = (
    options?: IDropdownOption<ComponentFramework.WebApi.Entity>[],
    defaultRenderer?: (options?: IDropdownOption<ComponentFramework.WebApi.Entity>[]) => JSX.Element | null
  ) => {
    const option = options?.at(0);

    if (!option || !showIcon) {
      return defaultRenderer ? defaultRenderer(options) : null;
    }

    let imgSrc;

    if (showIcon === ShowIconOptions.EntityIcon) {
      imgSrc = entityIcon;
    } else if (showIcon === ShowIconOptions.RecordImage) {
      if (metadata && metadata.lookupEntity.RecordImageUrlTemplate && option.data && recordImageTemplate) {
        imgSrc = recordImageTemplate(option.data);
      }
    }

    let selectedItemTemplateCompiled: HandlebarsTemplateDelegate | null = null;

    if (selectedItemTemplate) {
      selectedItemTemplateCompiled = Handlebars.compile(selectedItemTemplate);
    }

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        {imgSrc ? <Image imageFit={ImageFit.cover} src={imgSrc} width={imgSize} height={imgSize}></Image> : null}
        <Label styles={selectionOptionLabelStyles}>
          {selectedItemTemplateCompiled ? selectedItemTemplateCompiled(option.data) : option.text}
        </Label>
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
    if (!metadata) return;

    let createdRecord: ComponentFramework.LookupValue | null = null;
    if (metadata.lookupEntity.IsQuickCreateEnabled) {
      const result = await Xrm.Navigation.openForm({
        entityName: metadata.lookupEntity.LogicalName,
        useQuickCreateForm: true,
      });
      if (!result.savedEntityReference?.length) return;

      createdRecord = result.savedEntityReference[0];
    } else {
      const result = await Xrm.Navigation.navigateTo(
        { pageType: "entityrecord", entityName: metadata.lookupEntity.LogicalName },
        { target: 2 }
      );

      if (!result?.savedEntityReference?.length) return;

      createdRecord = result.savedEntityReference[0];
    }

    if (!createdRecord) return;

    onChange?.(createdRecord);

    queryClient.invalidateQueries({ queryKey: [`${metadata.lookupEntity.EntitySetName}_${lookupViewId}_fetchdata`] });
  };

  const onLookupPanelCommandClick = async () => {
    if (!allowLookupPanel) return;
    if (!metadata) return;
    const result = await Xrm.Utility.lookupObjects({
      entityTypes: [metadata.lookupEntity.LogicalName],
      defaultEntityType: metadata.lookupEntity.LogicalName,
      allowMultiSelect: false,
      disableMru: true,
      filters: customFilter?.length
        ? [{ entityLogicalName: metadata.lookupEntity.LogicalName, filterXml: getCustomFilterString(customFilter) }]
        : [],
    });

    if (!result) return;

    const selectedItem = {
      entityType: metadata.lookupEntity.LogicalName,
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

  return (
    <Stack horizontal styles={{ root: { width: "100%" } }}>
      <Stack.Item grow>
        <Dropdown
          selectedKey={selectedId ?? undefined}
          placeholder="---"
          options={getDropdownOptions()}
          disabled={disabled}
          onChange={(event, option) => {
            if (!onChange) return;
            if (!metadata) return;

            if (!option) {
              onChange(null);
              return;
            }

            const data = option.data as ComponentFramework.WebApi.Entity;
            const selectedItem = {
              entityType: metadata.lookupEntity.LogicalName,
              id: data[metadata.lookupEntity.PrimaryIdAttribute],
              name: data[metadata.lookupEntity.PrimaryNameAttribute],
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
};

export default function LookdownControl(props: LookdownProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}

function getFetchTemplateString(fetchXml: string, customFilter?: string | null, additionalColumns?: string[]) {
  if (fetchXml === "") return fetchXml;

  if (!customFilter && additionalColumns?.length === 0) return fetchXml;

  let templateString = fetchXml;

  const parser = new DOMParser();
  const doc = parser.parseFromString(fetchXml, "application/xml");
  const entity = doc.querySelector("entity");

  if (customFilter) {
    // create customFilterElement from customFilter
    const customFilterEl = parser.parseFromString(customFilter, "text/xml");
    entity?.appendChild(customFilterEl.documentElement);
  }

  if (additionalColumns?.length) {
    // get unique list of additionalColumns
    additionalColumns = [...new Set(additionalColumns)];

    // get all attribute names from fetchXml
    const attributes = doc.querySelectorAll("attribute");
    const attributeNames = Array.from(attributes).map((a) => a.getAttribute("name"));

    // get list of strings from additionalColumns not in attributeNames
    const addColumns = additionalColumns.filter((c) => !attributeNames.includes(c));

    addColumns.forEach((c) => {
      const attributeEl = parser.parseFromString(`<attribute name='${c}' />`, "text/xml");
      entity?.appendChild(attributeEl.documentElement);
    });
  }

  templateString = new XMLSerializer().serializeToString(doc);

  const fetchXmlTemplate = Handlebars.compile(templateString);
  return fetchXmlTemplate(getCurrentRecord());
}

function getCustomFilterString(customFilter: string) {
  if (customFilter === "") return customFilter;

  const parser = new DOMParser();
  const doc = parser.parseFromString(customFilter, "application/xml");
  const filter = doc.querySelector("filter");

  if (!filter) return customFilter;

  const filterTemplate = Handlebars.compile(customFilter);
  return filterTemplate(getCurrentRecord());
}

/**
 * Getting the variables from the Handlebars template.
 * Supports helpers too.
 * @param input
 */
function getHandlebarsVariables(input: string): string[] {
  const ast = Handlebars.parseWithoutProcessing(input);

  return ast.body
    .filter(({ type }: hbs.AST.Statement) => type === "MustacheStatement")
    .map((statement: hbs.AST.Statement) => {
      const moustacheStatement: hbs.AST.MustacheStatement = statement as hbs.AST.MustacheStatement;
      const paramsExpressionList = moustacheStatement.params as hbs.AST.PathExpression[];
      const pathExpression = moustacheStatement.path as hbs.AST.PathExpression;

      return paramsExpressionList[0]?.original || pathExpression.original;
    });
}
