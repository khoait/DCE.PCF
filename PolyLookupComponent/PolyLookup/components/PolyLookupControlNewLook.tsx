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
import { EntityOption, EntityReference, PolyLookupProps, RelationshipTypeEnum, TagAction } from "../types/typings";
import { sprintf } from "sprintf-js";

const useStyle = makeStyles({
  tagGroup: {
    flexWrap: "wrap",
    gap: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
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
  tagAction,
  defaultLanguagePack,
  languagePackPath,
  fluentDesign,
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const { tagGroup } = useStyle();

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
    isSuccess: isSuccessEntityOptions,
    isError: isErrorEntityOptions,
    error: errorEntityOptions,
  } = useEntityOptions(metadata, lookupViewConfig, searchText, pageSize);

  const { mutate: associateQuery } = useAssociateQuery(
    metadata,
    currentRecordId,
    relationshipType,
    clientUrl,
    languagePack
  );

  const { mutate: disassociateQuery } = useDisassociateQuery(metadata, currentRecordId, relationshipType, languagePack);

  useEffect(() => {
    if (isFetchingSelectedItems || !isLoadingSelectedItemsSuccess || !onChange) return;

    onChange(
      selectedItems?.map((i) => {
        return {
          id: i.associatedId,
          name: i.associatedName,
          etn: metadata?.associatedEntity.LogicalName ?? "",
        } as EntityReference;
      })
    );
  }, [isFetchingSelectedItems, isLoadingSelectedItemsSuccess, onChange]);

  const shouldDisable = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return true;
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return true;
    }
    return isError;
  };

  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems || isLoadingLookupView) && !shouldDisable();
  const isError = isErrorMetadata || isErrorLookupView || isErrorSelectedItems;

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

    if (selectedItems?.length || disabled) {
      return "";
    }

    return metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.Placeholder, metadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.PlaceholderDefault;
  };

  const handleOnOptionSelect: TagPickerProps["onOptionSelect"] = (_e, { value, selectedOptions }) => {
    if (itemLimit && (selectedItems?.length ?? 0) >= itemLimit) {
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
      associateQuery(value, {
        onSuccess: () => {
          queryClient.invalidateQueries({
            queryKey: ["selectedItems"],
          });
        },
      });
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
      disassociateQuery(value, {
        onSuccess: () => {
          queryClient.invalidateQueries({
            queryKey: ["selectedItems"],
          });
        },
      });
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
        size="large"
        selectedOptions={selectedItems?.map((i) => i.id) ?? []}
        onOptionSelect={handleOnOptionSelect}
        disabled={shouldDisable()}
        noPopover={shouldDisable() || disabled}
      >
        <TagPickerControl
          secondaryAction={
            onQuickCreate && !entityOptions?.length ? (
              <Button appearance="transparent" size="small" shape="rounded" onClick={handleQuickCreate}>
                {languagePack.AddNewLabel}
              </Button>
            ) : undefined
          }
        >
          {placeholder ? (
            <Text>{placeholder}</Text>
          ) : (
            <>
              <TagGroup className={tagGroup} onDismiss={handleOnOptionDismiss}>
                {(selectedItems ?? selectedEntitiesCreate).map((i) => (
                  <InteractionTag key={i.id} shape="rounded" appearance={tagAction ? "brand" : "outline"} value={i.id}>
                    <InteractionTagPrimary
                      hasSecondaryAction={!disabled}
                      media={
                        <Avatar
                          name={i.associatedName}
                          image={{ src: i.iconSrc }}
                          color={i.iconSrc?.startsWith("/WebResource") ? "neutral" : "colorful"}
                          aria-hidden
                        />
                      }
                      onClick={() => handleOnItemClick(i)}
                    >
                      {i.selectedOptionText}
                    </InteractionTagPrimary>
                    {disabled ? null : <InteractionTagSecondary aria-label="remove" />}
                  </InteractionTag>
                ))}
              </TagGroup>
              {disabled ? null : <TagPickerInput value={searchText} onChange={(e) => setSearchText(e.target.value)} />}
            </>
          )}
        </TagPickerControl>
        <TagPickerList>
          {entityOptions?.length && isSuccessEntityOptions
            ? entityOptions.map((option) => (
                <TagPickerOption
                  media={
                    <Avatar
                      shape="square"
                      name={option.associatedName}
                      image={{ src: option.iconSrc }}
                      color={option.iconSrc?.startsWith("/WebResource") ? "neutral" : "colorful"}
                      aria-hidden
                    />
                  }
                  value={option.id}
                  key={option.id}
                >
                  {option.optionText}
                </TagPickerOption>
              ))
            : "No options available"}
        </TagPickerList>
      </TagPicker>
    </FluentProvider>
  );
}
