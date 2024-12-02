import {
  Avatar,
  Button,
  Divider,
  FluentProvider,
  InteractionTag,
  InteractionTagPrimary,
  InteractionTagSecondary,
  TagGroup,
  TagGroupProps,
  TagPicker,
  TagPickerControl,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerProps,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import { useQueryClient } from "@tanstack/react-query";
import React, { useEffect, useState } from "react";
import { sprintf } from "sprintf-js";
import { useAttributeOnChange } from "../hooks/useAttributeOnChange";
import {
  EntityOption,
  EntityReference,
  PolyLookupProps,
  RelationshipTypeEnum,
  ShowIconOptions,
  ShowOptionDetailsEnum,
  TagAction,
} from "../types/typings";
import { SuggestionInfo } from "./SuggestionInfo";
import { useLanguagePack } from "../hooks/queries/useLanguagePack";
import { useEntityOptions } from "../hooks/queries/useEntityOptions";
import { useLookupViewConfig } from "../hooks/queries/useLookupViewConfig";
import { useMetadata } from "../hooks/queries/useMetadata";
import { useSelectedItems } from "../hooks/queries/useSelectedItems";
import { useAssociate } from "../hooks/mutations/useAssociate";
import { useDisassociate } from "../hooks/mutations/useDisassociate";
import { OptionList } from "./OptionList";
import { getPlaceholder } from "../utils";

const useStyle = makeStyles({
  tagGroup: {
    flexWrap: "wrap",
    gap: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
  },
  marginLeft: {
    marginLeft: tokens.spacingHorizontalS,
  },
  underline: {
    textDecoration: "underline",
  },
  borderTransparent: {
    borderLeftColor: "transparent",
  },
  tagFontSize: {
    fontSize: tokens.fontSizeBase300,
  },
  iconFontSize: {
    fontSize: tokens.fontSizeBase200,
  },
  transparentBackground: {
    backgroundColor: tokens.colorTransparentBackground,
  },
});

export default function PolyLookupControlNewLook({
  currentTable,
  currentRecordId,
  relationshipName,
  relationship2Name,
  relationshipType,
  clientUrl,
  lookupView,
  itemLimit,
  pageSize,
  disabled,
  outputSelectedItems,
  showIcon,
  tagAction,
  showOptionDetails,
  selectedItemTemplate,
  defaultLanguagePack,
  languagePackPath,
  isAuthoringMode,
  formType,
  fluentDesign,
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const { tagGroup, marginLeft, underline, borderTransparent, tagFontSize, iconFontSize, transparentBackground } =
    useStyle();

  const tagStyle = mergeClasses(!!tagAction && underline);
  const iconStyle = mergeClasses(borderTransparent, iconFontSize);
  const imageStyle = mergeClasses(
    showIcon === ShowIconOptions.EntityIcon && marginLeft,
    showIcon === ShowIconOptions.EntityIcon && transparentBackground
  );

  const [searchText, setSearchText] = useState<string>("");

  const [selectedEntitiesCreate, setSelectedEntitiesCreate] = useState<EntityOption[]>([]);

  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
    isError: isErrorMetadata,
    error: errorMetadata,
  } = useMetadata(
    currentTable,
    relationshipName,
    relationshipType === RelationshipTypeEnum.Custom || relationshipType === RelationshipTypeEnum.Connection
      ? relationship2Name
      : undefined
  );

  // get LookupView configuration
  const {
    data: lookupViewConfig,
    isPending: isLoadingLookupView,
    isSuccess: isSuccessLookupViewConfig,
    isError: isErrorLookupView,
    error: errorLookupView,
  } = useLookupViewConfig(
    metadata?.associatedEntity.LogicalName,
    metadata?.associatedEntity.PrimaryIdAttribute,
    metadata?.associatedEntity.PrimaryNameAttribute,
    lookupView
  );

  // get selected items
  const {
    data: selectedItems,
    isPending: isLoadingSelectedItems,
    isFetching: isFetchingSelectedItems,
    isSuccess: isLoadingSelectedItemsSuccess,
    isError: isErrorSelectedItems,
    error: errorSelectedItems,
  } = useSelectedItems(metadata, currentRecordId, formType, lookupViewConfig?.firstAttribute, selectedItemTemplate);

  const {
    data: entityOptions,
    isPending: isLoadingEntityOptions,
    isFetching: isFetchingEntityOptions,
    isSuccess: isSuccessEntityOptions,
    isError: isErrorEntityOptions,
    error: errorEntityOptions,
  } = useEntityOptions(metadata, lookupViewConfig, searchText, pageSize);

  const selectedOptions = formType === XrmEnum.FormType.Create ? selectedEntitiesCreate : selectedItems;

  const optionList =
    entityOptions?.records.filter((option) => !selectedOptions?.some((i) => i.associatedId === option.associatedId)) ??
    [];

  const { mutate: associateQuery } = useAssociate(metadata, currentRecordId, relationshipType, clientUrl, languagePack);

  const { mutate: disassociateQuery } = useDisassociate(metadata, currentRecordId, relationshipType, languagePack);

  useAttributeOnChange(lookupViewConfig?.fetchXml ?? "");

  useEffect(() => {
    if (isFetchingSelectedItems || !isLoadingSelectedItemsSuccess || !onChange) return;

    onChange(
      selectedOptions?.map((i) => {
        return {
          id: i.associatedId,
          name: i.selectedOptionText,
          etn: metadata?.associatedEntity.LogicalName ?? "",
        } as EntityReference;
      })
    );
  }, [isFetchingSelectedItems, isLoadingSelectedItemsSuccess, onChange]);

  const isDataLoading = isLoadingMetadata || isLoadingLookupView || isLoadingSelectedItems;
  const isError = isErrorMetadata || isErrorLookupView || isErrorSelectedItems;
  const isSupported =
    formType === XrmEnum.FormType.Update ||
    XrmEnum.FormType.ReadOnly ||
    XrmEnum.FormType.Disabled ||
    (formType === XrmEnum.FormType.Create && outputSelectedItems);

  const handleOnOptionSelect: TagPickerProps["onOptionSelect"] = (_e, { value, selectedOptions }) => {
    if (itemLimit && (selectedOptions?.length ?? 0) > itemLimit) {
      return;
    }

    if (formType === XrmEnum.FormType.Create) {
      const selectedEntity = entityOptions?.records.find((i) => i.id === value);
      if (!selectedEntity) {
        return;
      }
      const selectedEntities = [...selectedEntitiesCreate, selectedEntity];
      setSelectedEntitiesCreate(selectedEntities);
      onChange?.(
        selectedEntities?.map((i) => {
          return {
            id: i.associatedId,
            name: i.associatedName,
            etn: metadata?.associatedEntity.LogicalName ?? "",
          } as EntityReference;
        })
      );
    } else if (formType === XrmEnum.FormType.Update) {
      associateQuery(value);
    }

    setSearchText("");
  };

  const handleOnOptionDismiss: TagGroupProps["onDismiss"] = (_e, { value }) => {
    if (formType === XrmEnum.FormType.Create) {
      const selectedEntities = selectedEntitiesCreate.filter((e) => e.id !== value);
      setSelectedEntitiesCreate(selectedEntities);
      onChange?.(
        selectedEntities?.map((i) => {
          return {
            id: i.associatedId,
            name: i.associatedName,
            etn: metadata?.associatedEntity.LogicalName ?? "",
          } as EntityReference;
        })
      );
    } else if (formType === XrmEnum.FormType.Update) {
      disassociateQuery(value);
    }
  };

  const handleOnItemClick = async (item: EntityOption) => {
    if (!tagAction) return;
    if (!metadata?.associatedEntity) return;

    let targetEntity = metadata?.associatedEntity.LogicalName;
    let targetId = item.associatedId;

    if (
      (relationshipType === RelationshipTypeEnum.Custom || relationshipType === RelationshipTypeEnum.Connection) &&
      (tagAction === TagAction.OpenDialogIntersect || tagAction === TagAction.OpenInlineIntersect)
    ) {
      targetEntity = metadata.intersectEntity.LogicalName;
      targetId = item.intersectId;
    }

    await Xrm.Navigation.navigateTo(
      {
        pageType: "entityrecord",
        entityName: targetEntity,
        entityId: targetId,
      },
      { target: tagAction === TagAction.OpenDialog || tagAction === TagAction.OpenDialogIntersect ? 2 : 1 }
    );

    if (tagAction === TagAction.OpenDialog || tagAction === TagAction.OpenDialogIntersect) {
      queryClient.invalidateQueries({
        queryKey: ["selectedItems"],
      });
    }
  };

  const handleQuickCreate = async () => {
    if (!metadata?.associatedEntity || !onQuickCreate) return;

    onQuickCreate(
      metadata?.associatedEntity.LogicalName,
      metadata?.associatedEntity.PrimaryNameAttribute,
      searchText,
      metadata?.associatedEntity.IsQuickCreateEnabled
    )
      .then((result) => {
        if (result) {
          associateQuery(result, {
            onSuccess: () => {
              setSearchText("");
              queryClient.invalidateQueries({
                queryKey: ["selectedItems"],
              });
            },
          });
        }
      })
      .catch((err) => {
        console.log(err);
      });
  };

  const placeholder = getPlaceholder(
    isAuthoringMode ?? false,
    formType ?? XrmEnum.FormType.Undefined,
    outputSelectedItems ?? false,
    isDataLoading,
    isError,
    errorLookupView,
    selectedOptions?.length ?? 0,
    disabled ?? false,
    metadata?.associatedEntity.DisplayCollectionNameLocalized,
    languagePack
  );

  return (
    <FluentProvider style={{ width: "100%" }} theme={fluentDesign?.tokenTheme}>
      <TagPicker
        appearance="filled-darker"
        size="medium"
        selectedOptions={selectedOptions?.map((i) => i.id) ?? []}
        onOptionSelect={handleOnOptionSelect}
        disabled={!isSupported}
        noPopover={!isSupported || disabled || (itemLimit !== undefined && (selectedOptions?.length ?? 0) >= itemLimit)}
      >
        <TagPickerControl
          secondaryAction={
            onQuickCreate && !entityOptions?.records.length && !isFetchingEntityOptions && !isDataLoading ? (
              <Button appearance="transparent" size="small" shape="rounded" onClick={handleQuickCreate}>
                {languagePack.AddNewLabel}
              </Button>
            ) : undefined
          }
        >
          <TagGroup className={tagGroup} onDismiss={handleOnOptionDismiss}>
            {selectedOptions?.map((i) => (
              <InteractionTag
                key={i.id}
                shape="rounded"
                size="small"
                appearance={tagAction ? "brand" : "outline"}
                value={i.id}
              >
                <InteractionTagPrimary
                  className={tagStyle}
                  hasSecondaryAction={!disabled}
                  media={
                    showIcon ? (
                      <Avatar
                        className={imageStyle}
                        size={showIcon === ShowIconOptions.EntityIcon ? 16 : 20}
                        name={showIcon === ShowIconOptions.EntityIcon ? "" : i.optionText}
                        image={{
                          className: transparentBackground,
                          src:
                            showIcon === ShowIconOptions.EntityIcon
                              ? metadata?.associatedEntity.EntityIconUrl
                              : i.iconSrc,
                        }}
                        color={showIcon === ShowIconOptions.EntityIcon ? "neutral" : "colorful"}
                        aria-hidden
                      />
                    ) : undefined
                  }
                  onClick={() => handleOnItemClick(i)}
                >
                  <span className={tagFontSize}>{i.selectedOptionText}</span>
                </InteractionTagPrimary>
                {disabled ? null : <InteractionTagSecondary className={iconStyle} aria-label="remove" />}
              </InteractionTag>
            ))}
          </TagGroup>

          <TagPickerInput
            placeholder={placeholder}
            value={searchText}
            disabled={
              disabled ||
              isDataLoading ||
              isError ||
              !isSupported ||
              (itemLimit !== undefined && (selectedOptions?.length ?? 0) >= itemLimit)
            }
            onChange={(e) => setSearchText(e.target.value)}
          />
        </TagPickerControl>
        {disabled || !isSupported || (itemLimit !== undefined && (selectedOptions?.length ?? 0) >= itemLimit) ? (
          <></>
        ) : (
          <TagPickerList>
            <OptionList
              options={optionList}
              columns={lookupViewConfig?.columns ?? []}
              languagePack={languagePack}
              isLoading={isLoadingEntityOptions}
              hasMoreRecords={entityOptions?.moreRecords}
              showIcon={showIcon}
              entityIconUrl={metadata?.associatedEntity.EntityIconUrl}
              showOptionDetails={showOptionDetails}
            />
          </TagPickerList>
        )}
      </TagPicker>
    </FluentProvider>
  );
}
