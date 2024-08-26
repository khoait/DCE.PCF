import {
  ActionButton,
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
  ITag,
  ITagItemProps,
  TagItem,
  TagPicker,
  TagPickerBase,
  ValidationState,
} from "@fluentui/react";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import Handlebars from "handlebars";
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { sprintf } from "sprintf-js";
import {
  associateRecord,
  createRecord,
  deleteRecord,
  disassociateRecord,
  getCurrentRecord,
  retrieveMultipleFetch,
  useLanguagePack,
  useLookupViewConfig,
  useMetadata,
  useSelectedItems,
} from "../services/DataverseService";
import { IMetadata } from "../types/metadata";
import { EntityReference, PolyLookupProps, RelationshipTypeEnum, TagAction } from "../types/typings";
import { SuggestionInfo } from "./SuggestionInfo";

const parser = new DOMParser();
const serializer = new XMLSerializer();

interface ITagWithData extends ITag {
  data: ComponentFramework.WebApi.Entity;
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
  tagAction,
  defaultLanguagePack,
  languagePackPath,
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const [selectedItemsCreate, setSelectedItemsCreate] = useState<ComponentFramework.WebApi.Entity[]>([]);

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

  // filter query
  const filterQuery = useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText: string; pageSizeParam: number | undefined }) => {
      let fetchXml = lookupViewConfig?.fetchXml ?? "";

      let shouldDefaultSearch = false;
      if (!lookupView && lookupViewConfig?.isSystemLookupView) {
        shouldDefaultSearch = true;
      } else {
        if (
          !fetchXml.includes("{{PolyLookupSearch}}") &&
          !fetchXml.includes("{{ PolyLookupSearch}}") &&
          !fetchXml.includes("{{PolyLookupSearch }}") &&
          !fetchXml.includes("{{ PolyLookupSearch }}")
        ) {
          shouldDefaultSearch = true;
        }

        const currentRecord = getCurrentRecord();
        fetchXml =
          fetchXmlTemplateFn?.({
            ...currentRecord,
            PolyLookupSearch: searchText,
          }) ?? fetchXml;
      }

      if (shouldDefaultSearch) {
        // if lookup view is not specified and using default lookup fiew,
        // add filter condition to fetchxml to support search
        const doc = parser.parseFromString(fetchXml, "application/xml");
        const entities = doc.documentElement.getElementsByTagName("entity");
        for (let i = 0; i < entities.length; i++) {
          const entity = entities[i];
          if (entity.getAttribute("name") === metadata?.associatedEntity.LogicalName) {
            const filter = doc.createElement("filter");
            const condition = doc.createElement("condition");
            condition.setAttribute("attribute", metadata?.associatedEntity.PrimaryNameAttribute ?? "");
            if (searchText.includes("*")) {
              const beginsWithWildCard = searchText.startsWith("*") ? true : false;
              const endsWithWildCard = searchText.endsWith("*") ? true : false;
              if (beginsWithWildCard || endsWithWildCard) {
                if (beginsWithWildCard) {
                  searchText = searchText.replace("*", "");
                }
                if (endsWithWildCard) {
                  searchText = searchText.substring(0, searchText.length - 1);
                }
                searchText = "%" + searchText + "%";
              } else {
                searchText = searchText + "%";
              }
              searchText = searchText.split("*").join("%");
              condition.setAttribute("operator", "like");
              condition.setAttribute("value", `${searchText}`);
            } else {
              condition.setAttribute("operator", "begins-with");
              condition.setAttribute("value", `${searchText}`);
            }
            filter.appendChild(condition);
            entity.appendChild(filter);
          }
        }
        fetchXml = serializer.serializeToString(doc);
      }

      return retrieveMultipleFetch(metadata?.associatedEntity.EntitySetName, fetchXml, 1, pageSizeParam);
    },
  });

  // associate query
  const associateQuery = useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return associateRecord(
          metadata?.currentEntity.EntitySetName,
          currentRecordId,
          metadata?.associatedEntity?.EntitySetName,
          id,
          relationshipName,
          clientUrl
        );
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return createRecord(metadata?.intersectEntity.EntitySetName, {
          [`${metadata?.currentEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.currentEntity.EntitySetName}(${currentRecordId})`,
          [`${metadata?.associatedEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.associatedEntity.EntitySetName}(${id})`,
        });
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: (data, variables, context) => {
      queryClient.invalidateQueries({
        queryKey: ["selectedItems"],
      });
    },
  });

  // disassociate query
  const disassociateQuery = useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return disassociateRecord(metadata?.currentEntity?.EntitySetName, currentRecordId, relationshipName, id);
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return deleteRecord(metadata?.intersectEntity.EntitySetName, id);
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: (data, variables, context) => {
      queryClient.invalidateQueries({
        queryKey: ["selectedItems"],
      });
    },
  });

  const filterSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      try {
        const results = await filterQuery.mutateAsync({ searchText: filterText, pageSizeParam: pageSize });
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
        const results = await filterQuery.mutateAsync({
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
        const results = await filterQuery.mutateAsync({ searchText: "", pageSizeParam: pageSize });
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
            return (
              data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
              i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            );
          })
      );

      const added = selectedTags
        ?.filter((t) => {
          const data = (t as ITagWithData).data;
          return !selectedItemsCreate?.some(
            (i) =>
              i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
              data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
          );
        })
        .map((t) => (t as ITagWithData).data);

      const oldRemoved = selectedItemsCreate?.filter(
        (o) =>
          !removed?.some(
            (r) =>
              r[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
              o[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
          )
      );

      const newSelectedItems = [...oldRemoved, ...(added ?? [])];
      setSelectedItemsCreate(newSelectedItems);

      onChange?.(
        newSelectedItems?.map((i) => {
          return {
            id: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
            name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
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
              return data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] === i.associatedId;
            })
        )
        .map((i) => (relationshipType === RelationshipTypeEnum.ManyToMany ? i.associatedId : i.id));

      const added = selectedTags
        ?.filter((t) => {
          const data = (t as ITagWithData).data;
          return !selectedItems?.some(
            (i) => i.associatedId === data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
          );
        })
        .map((t) => t.key);

      removed?.forEach((id) => disassociateQuery.mutate(id as string));
      added?.forEach((id) => associateQuery.mutate(id as string));
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
            associateQuery.mutate(result);
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
      targetId = item.data[metadata?.intersectEntity.PrimaryIdAttribute];
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
          root: [
            !!tagAction && {
              "&:focus-within button": {
                color: "#fff",
              },
              "&:focus-within:hover button": {
                color: "#fff",
              },
              "&:focus-within .ms-TagItem-close:hover": {
                color: "#fff",
                backgroundColor: "#005a9e",
              },
            },
          ],
        };

        return (
          <TagItem {...props}>
            {tagAction ? (
              <ActionButton
                styles={{ root: { height: "100%" } }}
                onClick={() => onTagClick(props.item as ITagWithData)}
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
          let displayValue = data[column + "@OData.Community.Display.V1.FormattedValue"];
          if (!displayValue) {
            displayValue = data[column];
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

function getSuggestionTags(
  suggestions: ComponentFramework.WebApi.Entity[] | undefined,
  metadata: IMetadata | undefined
) {
  return (
    suggestions?.map(
      (i) =>
        ({
          key: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ?? "",
          name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""] ?? "",
          data: i,
        }) as ITagWithData
    ) ?? []
  );
}
