import {
  ActionButton,
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
  ImageFit,
  ITag,
  ITagItemProps,
  Persona,
  PersonaSize,
  TagItem,
  TagPicker,
  TagPickerBase,
  ValidationState,
} from "@fluentui/react";
import { useQueryClient } from "@tanstack/react-query";
import Handlebars from "handlebars";
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { sprintf } from "sprintf-js";
import {
  useAssociateQuery,
  useDisassociateQuery,
  useFilterQuery,
  useLanguagePack,
  useLookupViewConfig,
  useMetadata,
  useSelectedItems,
} from "../services/DataverseService";
import { IMetadata } from "../types/metadata";
import {
  EntityOption,
  EntityReference,
  PolyLookupProps,
  RelationshipTypeEnum,
  ShowIconOptions,
  TagAction,
} from "../types/typings";
import { SuggestionInfo } from "./SuggestionInfo";
import { Avatar } from "@fluentui/react-components";

interface ITagWithData extends ITag {
  data: EntityOption;
}

export default function PolyLookupControlClassic({
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
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const [selectedItemsCreate, setSelectedItemsCreate] = useState<EntityOption[]>([]);

  const pickerRef = useRef<TagPickerBase>(null);

  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: languagePack.SuggestionListHeaderDefaultLabel,
    noResultsFoundText: languagePack.EmptyListDefaultMessage,
    forceResolveText: languagePack.AddNewLabel,
    showForceResolve: () => onQuickCreate !== undefined,
    resultsFooter: () => <div>{languagePack.NoMoreRecordsMessage}</div>,
    resultsFooterFull: () => <div>{languagePack.SuggestionListFullMessage}</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: languagePack.LoadMoreLabel,
  };

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
    isSuccess: isSuccessLookupViewConfig,
    isError: isErrorLookupView,
    error: errorLookupView,
  } = useLookupViewConfig(
    metadata?.associatedEntity.LogicalName,
    metadata?.associatedEntity.PrimaryIdAttribute,
    metadata?.associatedEntity.PrimaryNameAttribute,
    lookupView
  );

  const fetchXmlTemplateFn = useMemo(
    () => (lookupViewConfig?.fetchXml ? Handlebars.compile(lookupViewConfig?.fetchXml) : null),
    [lookupViewConfig?.fetchXml]
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

  const { mutateAsync: filterQueryAsync } = useFilterQuery(metadata, lookupViewConfig);

  const { mutate: associateQuery } = useAssociateQuery(
    metadata,
    currentRecordId,
    relationshipType,
    clientUrl,
    languagePack
  );

  const { mutate: disassociateQuery } = useDisassociateQuery(metadata, currentRecordId, relationshipType, languagePack);

  const filterSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      try {
        const results = await filterQueryAsync({ searchText: filterText, pageSizeParam: pageSize });
        return getSuggestionTags(results, metadata);
      } catch {
        // ignore error and return empty array
      }
      return [];
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showMoreSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      try {
        const results = await filterQueryAsync({
          searchText: filterText,
          pageSizeParam: (pageSize ?? 50) * 2 + 1,
        });
        return getSuggestionTags(results, metadata);
      } catch {
        // ignore error and return empty array
      }
      return [];
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showAllSuggestions = useCallback(
    async (selectedTags?: ITag[]): Promise<ITag[]> => {
      try {
        const results = await filterQueryAsync({ searchText: "", pageSizeParam: pageSize });
        return getSuggestionTags(results, metadata);
      } catch {
        // ignore error and return empty array
      }
      return [];
    },
    [metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onPickerChange = (selectedTags?: ITag[]): void => {
    if (formType === XrmEnum.FormType.Create) {
      const removed = selectedItemsCreate?.filter(
        (i) =>
          !selectedTags?.some((t) => {
            const data = (t as ITagWithData).data;
            return data.associatedId === i.associatedId;
          })
      );

      const added = selectedTags
        ?.filter((t) => {
          const data = (t as ITagWithData).data;
          return !selectedItemsCreate?.some((i) => i.associatedId === data.associatedId);
        })
        .map((t) => (t as ITagWithData).data);

      const oldRemoved = selectedItemsCreate?.filter((o) => !removed?.some((r) => r.associatedId === o.associatedId));

      const newSelectedItems = [...oldRemoved, ...(added ?? [])];
      setSelectedItemsCreate(newSelectedItems);

      onChange?.(
        newSelectedItems?.map((i) => {
          return {
            id: i.associatedId,
            name: i.associatedName,
            etn: metadata?.associatedEntity.LogicalName ?? "",
          } as EntityReference;
        })
      );
    } else if (formType === XrmEnum.FormType.Update) {
      const removed = selectedItems
        ?.filter(
          (i) =>
            !selectedTags?.some((t) => {
              const data = (t as ITagWithData).data;
              return data.associatedId === i.associatedId;
            })
        )
        .map((i) => (relationshipType === RelationshipTypeEnum.ManyToMany ? i.associatedId : i.id));

      const added = selectedTags
        ?.filter((t) => {
          const data = (t as ITagWithData).data;
          return !selectedItems?.some((i) => i.associatedId === data.associatedId);
        })
        .map((t) => t.key);

      removed?.forEach((id) => disassociateQuery(id as string));
      added?.forEach((id) => associateQuery(id as string));
    }
  };

  const onItemSelected = (item?: ITag): ITag | null => {
    if (!item) return null;

    const primaryIdAttribute = metadata?.associatedEntity.PrimaryIdAttribute ?? "";
    const isSelected = (selectedItems: ComponentFramework.WebApi.Entity[] | undefined) =>
      selectedItems?.some((i) => i[primaryIdAttribute] === item.key) ?? false;

    if (
      (formType === XrmEnum.FormType.Create && !isSelected(selectedItemsCreate)) ||
      (formType === XrmEnum.FormType.Update && !isSelected(selectedItems))
    ) {
      return item;
    }

    return null;
  };

  const onCreateNew = (input: string): ValidationState => {
    if (onQuickCreate) {
      onQuickCreate(
        metadata?.associatedEntity.LogicalName,
        metadata?.associatedEntity.PrimaryNameAttribute,
        input,
        metadata?.associatedEntity.IsQuickCreateEnabled
      )
        .then((result) => {
          if (result) {
            associateQuery(result);
            // TODO: fix this hack
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            pickerRef.current.input.current?._updateValue("");
          }
        })
        .catch((err) => {
          console.log(err);
        });
    }
    return ValidationState.invalid;
  };

  const onTagClick = async (item: ITagWithData) => {
    if (!tagAction) return;
    if (!metadata?.associatedEntity) return;

    let targetEntity = metadata?.associatedEntity.LogicalName;
    let targetId = item.key as string;

    if (
      (relationshipType === RelationshipTypeEnum.Custom || relationshipType === RelationshipTypeEnum.Connection) &&
      (tagAction === TagAction.OpenDialogIntersect || tagAction === TagAction.OpenInlineIntersect)
    ) {
      targetEntity = metadata.intersectEntity.LogicalName;
      targetId = item.data.intersectId;
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

    if (selectedItems?.length || selectedItemsCreate.length || disabled) {
      return "";
    }

    return metadata?.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.Placeholder, metadata?.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.PlaceholderDefault;
  };

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

  const isDataLoading = (isLoadingMetadata || isLoadingSelectedItems) && !shouldDisable();
  const isError = isErrorMetadata || isErrorLookupView || isErrorSelectedItems;

  if (metadata && isLoadingMetadataSuccess) {
    pickerSuggestionsProps.suggestionsHeaderText = metadata.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.SuggestionListHeaderLabel, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.SuggestionListHeaderDefaultLabel;

    pickerSuggestionsProps.noResultsFoundText = metadata.associatedEntity.DisplayCollectionNameLocalized
      ? sprintf(languagePack.EmptyListMessage, metadata.associatedEntity.DisplayCollectionNameLocalized)
      : languagePack.EmptyListDefaultMessage;
  }

  return (
    <TagPicker
      ref={pickerRef}
      selectedItems={
        isError
          ? []
          : getSuggestionTags(formType === XrmEnum.FormType.Create ? selectedItemsCreate : selectedItems, metadata)
      }
      onResolveSuggestions={filterSuggestions}
      onEmptyResolveSuggestions={showAllSuggestions}
      onGetMoreResults={showMoreSuggestions}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      styles={({ isFocused }) => {
        // eslint-disable-next-line react/prop-types
        const pickerStyles: Partial<IBasePickerStyles> = {
          root: { backgroundColor: "#fff", width: "100%" },
          input: { minWidth: "0", display: disabled ? "none" : "inline-block" },
          text: {
            minWidth: "0",
            borderColor: "transparent",
            borderWidth: disabled ? 0 : 1,
            borderRadius: 0,
            "&:after": {
              backgroundColor: "transparent",
              borderColor: isFocused && !disabled ? "#666" : "transparent",
              borderWidth: disabled ? 0 : 1,
              borderRadius: 0,
            },
            "&:hover:after": { backgroundColor: "transparent" },
          },
        };
        return pickerStyles;
      }}
      pickerSuggestionsProps={pickerSuggestionsProps}
      disabled={shouldDisable()}
      onRenderItem={(props: ITagItemProps) => {
        props.styles = {
          close: [disabled && { display: "none" }],
          text: [
            {
              flex: 1,
              margin: 0,
            },
          ],
          root: [
            {
              gap: "4px",
            },
            !!tagAction && {
              // "&:focus-within button": {
              //   color: "#fff",
              // },
              // "&:focus-within:hover button": {
              //   color: "#fff",
              // },
              "&:focus-within .ms-TagItem-close:hover": {
                color: "#fff",
                backgroundColor: "#005a9e",
              },
            },
          ],
        };

        const item = props.item as ITagWithData;
        const iconSize = showIcon === ShowIconOptions.RecordImage ? 24 : 16;

        return (
          <TagItem {...props}>
            {tagAction ? (
              <ActionButton
                styles={{
                  root: { display: "block", width: "100%", height: "100%", padding: 0 },
                  flexContainer: { gap: 4 },
                  icon: {
                    marginLeft: "4px",
                  },
                }}
                iconProps={
                  showIcon
                    ? {
                        styles: {
                          root: {
                            width: iconSize,
                            height: iconSize,
                            //margin: 0,
                            lineHeight: "normal",
                          },
                        },
                        imageProps: {
                          src:
                            showIcon === ShowIconOptions.EntityIcon
                              ? metadata?.associatedEntity.EntityIconUrl
                              : item.data.iconSrc,
                          imageFit: ImageFit.cover,
                        },
                        imageErrorAs: (imageProps) => (
                          <Persona
                            styles={{
                              details: {
                                display: "none",
                              },
                            }}
                            coinProps={{
                              styles: {
                                initials: {
                                  borderRadius: "2px",
                                },
                              },
                            }}
                            size={showIcon === ShowIconOptions.RecordImage ? PersonaSize.size24 : PersonaSize.size16}
                            text={item.name}
                            //imageInitials={item.name.split(" ").join("").substring(0, 1).toUpperCase()}
                          />
                        ),
                      }
                    : undefined
                }
                onClick={() => onTagClick(item)}
              >
                {props.item.name}
              </ActionButton>
            ) : (
              props.item.name
            )}
          </TagItem>
        );
      }}
      onRenderSuggestionsItem={(tag: ITag) => {
        const data = (tag as ITagWithData).data;
        const infoMap = new Map<string, string>();
        lookupViewConfig?.columns.forEach((column) => {
          let displayValue = data.entity[column + "@OData.Community.Display.V1.FormattedValue"];
          if (!displayValue) {
            displayValue = data.entity[column];
          }
          infoMap.set(column, displayValue ?? "");
        });
        return <SuggestionInfo infoMap={infoMap}></SuggestionInfo>;
      }}
      resolveDelay={100}
      inputProps={{
        placeholder: getPlaceholder(),
      }}
      pickerCalloutProps={{
        calloutMaxWidth: 500,
      }}
      itemLimit={itemLimit}
      onValidateInput={onCreateNew}
    />
  );
}

function getSuggestionTags(suggestions: EntityOption[] | undefined, metadata: IMetadata | undefined) {
  return (
    suggestions?.map(
      (i) =>
        ({
          key: i.associatedId ?? "",
          name: i.associatedName ?? "",
          data: i,
        }) as ITagWithData
    ) ?? []
  );
}
