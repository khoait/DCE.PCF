import React, { useCallback } from "react";
import {
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
  ITag,
  ITagItemProps,
  TagPicker,
  TagPickerBase,
  ValidationState,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation } from "@tanstack/react-query";
import {
  associateRecord,
  createRecord,
  deleteRecord,
  disassociateRecord,
  getCurrentRecord,
  retrieveMultipleFetch,
  useMetadata,
  useSelectedItems,
  useSuggestions,
} from "../services/DataverseService";
import { SuggestionInfo } from "./SuggestionInfo";
import { IMetadata } from "../types/metadata";

// TODO: fix this import in handlebars next version
import Handlebars from "handlebars/lib/handlebars";

const queryClient = new QueryClient();
const parser = new DOMParser();
const serializer = new XMLSerializer();

export enum RelationshipTypeEnum {
  ManyToMany,
  Custom,
  Connection,
}

export interface PolyLookupProps {
  currentTable: string;
  currentRecordId: string;
  relationshipName: string;
  relationship2Name: string | undefined;
  relationshipType: RelationshipTypeEnum;
  clientUrl: string;
  lookupView?: string;
  itemLimit?: number;
  pageSize?: number;
  disabled?: boolean;
  formType?: XrmEnum.FormType;
  outputSelectedItems?: boolean;
  onChange?: (selectedItems: ComponentFramework.EntityReference[] | undefined) => void;
  onQuickCreate?: (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => Promise<string | undefined>;
}

interface ITagWithData extends ITag {
  data: ComponentFramework.WebApi.Entity;
}

const Body = ({
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
  onChange,
  onQuickCreate,
}: PolyLookupProps) => {
  const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    noResultsFoundText: "No records found",
    forceResolveText: "Quick Create",
    showForceResolve: () => onQuickCreate !== undefined,
    resultsFooter: () => <div>No more records</div>,
    resultsFooterFull: () => <div>Refine search term for more</div>,
    resultsMaximumNumber: (pageSize ?? 50) * 2,
    searchForMoreText: "Load more",
  };
  const [selectedItemsCreate, setSelectedItemsCreate] = React.useState<ComponentFramework.WebApi.Entity[]>([]);

  const pickerRef = React.useRef<TagPickerBase>(null);

  const getPlaceholder = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return "Please save the record first";
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return "The control is not available in this form";
    }

    if (isDataLoading) {
      return "Loading";
    }

    if (selectedItems?.length || selectedItemsCreate.length || disabled) {
      return "";
    }

    return `Select ${metadata?.associatedEntity.DisplayCollectionNameLocalized ?? "an item"}`;
  };

  const shouldDisable = () => {
    if (formType === XrmEnum.FormType.Create) {
      if (!outputSelectedItems) {
        return true;
      }
    } else if (formType !== XrmEnum.FormType.Update) {
      return true;
    }
    return false;
  };

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(
    currentTable,
    relationshipName,
    relationshipType === RelationshipTypeEnum.Custom ? relationship2Name ?? undefined : undefined,
    lookupView
  );

  if (metadata && isLoadingMetadataSuccess) {
    pickerSuggestionsProps.suggestionsHeaderText = `Suggested ${metadata.associatedEntity.DisplayCollectionNameLocalized}`;
    pickerSuggestionsProps.noResultsFoundText = `No ${metadata.associatedEntity.DisplayCollectionNameLocalized} found`;
  }

  const associatedTableSetName = metadata?.associatedEntity.EntitySetName ?? "";
  const associatedFetchXml = metadata?.associatedView.fetchxml;

  const fetchXmlTemplate = Handlebars.compile(associatedFetchXml ?? "");

  // get top 50 suggestions from associated table
  const { data: suggestions, isLoading: isLoadingSuggestions } = useSuggestions(
    associatedTableSetName,
    fetchXmlTemplate,
    pageSize
  );

  // get selected items
  const {
    data: selectedItems,
    isInitialLoading: isLoadingSelectedItems,
    isSuccess: isLoadingSelectedItemsSuccess,
    refetch: selectedItemsRefetch,
  } = useSelectedItems(currentTable, currentRecordId, metadata, formType);

  if (isLoadingSelectedItemsSuccess && onChange) {
    onChange(
      selectedItems?.map((i) => {
        return {
          id: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
          name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
          etn: metadata?.associatedEntity.LogicalName ?? "",
        } as ComponentFramework.EntityReference;
      })
    );
  }

  // filter query
  const filterQuery = useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText: string; pageSizeParam: number | undefined }) => {
      let fetchXml = metadata?.associatedView.fetchxml ?? "";
      if (!lookupView && metadata?.associatedView.querytype === 64) {
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
            condition.setAttribute("operator", "like");
            condition.setAttribute("value", `%${searchText}%`);
            filter.appendChild(condition);
            entity.appendChild(filter);
          }
        }
        fetchXml = serializer.serializeToString(doc);
      } else {
        const currentRecord = getCurrentRecord();
        fetchXml = fetchXmlTemplate({
          ...currentRecord,
          PolyLookupSearch: searchText,
        });
      }
      return retrieveMultipleFetch(associatedTableSetName, fetchXml, 1, pageSizeParam);
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
      } else if (relationshipType === RelationshipTypeEnum.Custom) {
        return createRecord(metadata?.intersectEntity.EntitySetName, {
          [`${metadata?.currentIntesectAttribute}@odata.bind`]: `/${metadata?.currentEntity.EntitySetName}(${currentRecordId})`,
          [`${metadata?.associatedIntesectAttribute}@odata.bind`]: `/${metadata?.associatedEntity.EntitySetName}(${id})`,
        });
      }
      return Promise.reject("Relationship type not supported");
    },
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  // disassociate query
  const disassociateQuery = useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return disassociateRecord(metadata?.currentEntity?.EntitySetName, currentRecordId, relationshipName, id);
      } else if (relationshipType === RelationshipTypeEnum.Custom) {
        return deleteRecord(metadata?.intersectEntity.EntitySetName, id);
      }
      return Promise.reject("Relationship type not supported");
    },
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  const filterSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      const results = await filterQuery.mutateAsync({ searchText: filterText, pageSizeParam: pageSize });
      return getSuggestionTags(results, metadata);
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showMoreSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      const results = await filterQuery.mutateAsync({
        searchText: filterText,
        pageSizeParam: (pageSize ?? 50) * 2 + 1,
      });
      return getSuggestionTags(results, metadata);
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showAllSuggestions = useCallback(
    async (selectedTags?: ITag[]): Promise<ITag[]> => {
      return getSuggestionTags(suggestions, metadata);
    },
    [suggestions, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onPickerChange = useCallback(
    (selectedTags?: ITag[]): void => {
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

        if (onChange) {
          onChange(
            newSelectedItems?.map((i) => {
              return {
                id: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
                name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
                etn: metadata?.associatedEntity.LogicalName ?? "",
              } as ComponentFramework.EntityReference;
            })
          );
        }
      } else if (formType === XrmEnum.FormType.Update) {
        const removed = selectedItems
          ?.filter(
            (i) =>
              !selectedTags?.some((t) => {
                const data = (t as ITagWithData).data;
                return (
                  data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
                  i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
                );
              })
          )
          .map((i) =>
            relationshipType === RelationshipTypeEnum.ManyToMany
              ? i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
              : i[metadata?.intersectEntity.PrimaryIdAttribute ?? ""]
          );

        const added = selectedTags
          ?.filter((t) => {
            const data = (t as ITagWithData).data;
            return !selectedItems?.some(
              (i) =>
                i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] ===
                data[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            );
          })
          .map((t) => t.key);

        added?.forEach((id) => associateQuery.mutate(id as string));
        removed?.forEach((id) => disassociateQuery.mutate(id as string));
      }
    },
    [selectedItems, selectedItemsCreate, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onItemSelected = useCallback(
    (item?: ITag): ITag | null => {
      if (!item) return null;

      if (
        formType === XrmEnum.FormType.Create &&
        !selectedItemsCreate?.some((i) => i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] === item.key)
      ) {
        return item;
      } else if (
        formType === XrmEnum.FormType.Update &&
        !selectedItems?.some((i) => i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] === item.key)
      ) {
        return item;
      }
      return null;
    },
    [selectedItems, metadata?.associatedEntity.PrimaryIdAttribute]
  );

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

  const isDataLoading = (isLoadingMetadata || isLoadingSuggestions || isLoadingSelectedItems) && !shouldDisable();

  return (
    <TagPicker
      ref={pickerRef}
      selectedItems={getSuggestionTags(
        formType === XrmEnum.FormType.Create ? selectedItemsCreate : selectedItems,
        metadata
      )}
      onResolveSuggestions={filterSuggestions}
      onEmptyResolveSuggestions={showAllSuggestions}
      onGetMoreResults={showMoreSuggestions}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      styles={(props) => {
        // eslint-disable-next-line react/prop-types
        const isFocused = props.isFocused;
        const pickerStyles: Partial<IBasePickerStyles> = {
          root: { backgroundColor: "#fff", width: "100%" },
          input: { minWidth: "0", display: disabled ? "none" : "inline-block" },
          text: {
            minWidth: "0",
            borderColor: "transparent",
            borderWidth: 1,
            borderRadius: 1,
            "&:after": {
              backgroundColor: "transparent",
              borderColor: isFocused ? "#666" : "transparent",
              borderWidth: 1,
              borderRadius: 1,
            },
            "&:hover:after": { backgroundColor: disabled ? "rgba(50, 50, 50, 0.1)" : "transparent" },
          },
        };
        return pickerStyles;
      }}
      pickerSuggestionsProps={pickerSuggestionsProps}
      disabled={disabled || shouldDisable()}
      onRenderItem={(props: ITagItemProps) => {
        if (disabled) {
          props.styles = { close: { display: "none" } };
        }
        return TagPickerBase.defaultProps.onRenderItem(props);
      }}
      onRenderSuggestionsItem={(tag: ITag) => {
        const data = (tag as ITagWithData).data;
        const infoMap = new Map<string, string>();
        metadata?.associatedView?.layoutjson?.Rows?.at(0)?.Cells.forEach((cell) => {
          let displayValue = data[cell.Name + "@OData.Community.Display.V1.FormattedValue"];
          if (!displayValue) {
            displayValue = data[cell.Name];
          }
          infoMap.set(cell.Name, displayValue ?? "");
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
};

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
        } as ITagWithData)
    ) ?? []
  );
}

export default function PolyLookupControl(props: PolyLookupProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}
