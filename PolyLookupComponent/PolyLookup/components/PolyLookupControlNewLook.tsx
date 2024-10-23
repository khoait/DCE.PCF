import {
  Avatar,
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
  makeStyles,
  tokens,
  Text,
  Button,
  mergeClasses,
} from "@fluentui/react-components";
import { useQueryClient } from "@tanstack/react-query";
import React, { useEffect, useState } from "react";
import {
  useAssociateQuery,
  useDisassociateQuery,
  useEntityOptions,
  useLanguagePack,
  useLookupViewConfig,
  useMetadata,
  useSelectedItems,
} from "../services/DataverseService";
import {
  EntityOption,
  EntityReference,
  PolyLookupProps,
  RelationshipTypeEnum,
  ShowIconOptions,
  TagAction,
} from "../types/typings";
import { sprintf } from "sprintf-js";
import { SuggestionInfo } from "./SuggestionInfo";
import { useAttributeOnChange } from "../hooks/useAttributeOnChange";

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
  listBox: {
    maxHeight: "50vh",
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
  formType,
  outputSelectedItems,
  showIcon,
  tagAction,
  defaultLanguagePack,
  languagePackPath,
  fluentDesign,
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const {
    tagGroup,
    marginLeft,
    underline,
    borderTransparent,
    listBox,
    tagFontSize,
    iconFontSize,
    transparentBackground,
  } = useStyle();

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
  } = useSelectedItems(metadata, currentRecordId, formType);

  const {
    data: entityOptions,
    isPending: isLoadingEntityOptions,
    isFetching: isFetchingEntityOptions,
    isSuccess: isSuccessEntityOptions,
    isError: isErrorEntityOptions,
    error: errorEntityOptions,
  } = useEntityOptions(metadata, lookupViewConfig, searchText, pageSize);

  const selectedOptions = formType === XrmEnum.FormType.Create ? selectedEntitiesCreate : selectedItems;

  const optionList = entityOptions?.filter(
    (option) => !selectedOptions?.some((i) => i.associatedId === option.associatedId)
  );

  const { mutate: associateQuery } = useAssociateQuery(
    metadata,
    currentRecordId,
    relationshipType,
    clientUrl,
    languagePack
  );

  const { mutate: disassociateQuery } = useDisassociateQuery(metadata, currentRecordId, relationshipType, languagePack);

  useAttributeOnChange(lookupViewConfig?.fetchXml ?? "");

  useEffect(() => {
    if (isFetchingSelectedItems || !isLoadingSelectedItemsSuccess || !onChange) return;

    onChange(
      selectedOptions?.map((i) => {
        return {
          id: i.associatedId,
          name: i.associatedName,
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

  const getPlaceholder = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return languagePack.CreateFormNotSupportedMessage;
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return languagePack.ControlIsNotAvailableMessage;
    }

    if (isDataLoading) {
      return languagePack.LoadingMessage;
    }

    if (isError) {
      if (errorLookupView instanceof Error) {
        return languagePack.InvalidLookupViewMessage;
      }

      return languagePack.GenericErrorMessage;
    }

    if (selectedOptions?.length || disabled) {
      return "";
    }

    return metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.Placeholder, metadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.PlaceholderDefault;
  };

  const handleOnOptionSelect: TagPickerProps["onOptionSelect"] = (_e, { value, selectedOptions }) => {
    if (itemLimit && (selectedOptions?.length ?? 0) > itemLimit) {
      return;
    }

    if (formType === XrmEnum.FormType.Create) {
      const selectedEntity = entityOptions?.find((i) => i.id === value);
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
              queryClient.invalidateQueries({
                queryKey: ["selectedItems"],
              });
              setSearchText("");
            },
          });
        }
      })
      .catch((err) => {
        console.log(err);
      });
  };

  const placeholder = getPlaceholder();

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
            onQuickCreate && !entityOptions?.length && !isFetchingEntityOptions && !isDataLoading ? (
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
                        name={showIcon === ShowIconOptions.EntityIcon ? "" : i.selectedOptionText}
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
          <TagPickerList className={listBox}>
            {optionList?.length && isSuccessEntityOptions
              ? optionList.map((option) => (
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
                            src:
                              showIcon === ShowIconOptions.EntityIcon
                                ? metadata?.associatedEntity.EntityIconUrl
                                : option.iconSrc,
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
                    <SuggestionInfo data={option} columns={lookupViewConfig?.columns ?? []} />
                  </TagPickerOption>
                ))
              : "No options available"}
          </TagPickerList>
        )}
      </TagPicker>
    </FluentProvider>
  );
}
