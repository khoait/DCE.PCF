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
  disassociateRecord,
  getCurrentRecord,
  retrieveMultipleFetch,
  useMetadata,
  useSelectedItems,
  useSuggestions,
} from "../services/DataverseService";
import { SuggestionInfo } from "./SuggestionInfo";

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
  relationshipType: RelationshipTypeEnum;
  clientUrl: string;
  lookupView?: string;
  itemLimit?: number;
  pageSize?: number;
  disabled?: boolean;
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
  relationshipType,
  clientUrl,
  lookupView,
  itemLimit,
  pageSize,
  disabled,
  onChange,
  onQuickCreate,
}: PolyLookupProps) => {
  const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    noResultsFoundText: "No records found",
    forceResolveText: "Quick Create",
    showForceResolve: () => onQuickCreate !== undefined,
  };

  const pickerRef = React.useRef<TagPickerBase>(null);

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(currentTable, relationshipName, lookupView);

  if (metadata && isLoadingMetadataSuccess) {
    pickerSuggestionsProps.suggestionsHeaderText = `Suggested ${metadata.associatedEntity.DisplayCollectionName.UserLocalizedLabel.Label}`;
    pickerSuggestionsProps.noResultsFoundText = `No ${metadata.associatedEntity.DisplayCollectionName.UserLocalizedLabel.Label} found`;
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
    isLoading: isLoadingSelectedItems,
    isSuccess: isLoadingSelectedItemsSuccess,
    refetch: selectedItemsRefetch,
  } = useSelectedItems(currentTable, currentRecordId, relationshipName, metadata);

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
    mutationFn: (searchText: string) => {
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
      return retrieveMultipleFetch(associatedTableSetName, fetchXml, 1, pageSize);
    },
  });

  // associate query
  const associateQuery = useMutation({
    mutationFn: (id: string) =>
      associateRecord(
        metadata?.currentEntity.EntitySetName,
        currentRecordId,
        metadata?.associatedEntity?.EntitySetName,
        id,
        relationshipName,
        clientUrl
      ),
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  // disassociate query
  const disassociateQuery = useMutation({
    mutationFn: (id: string) =>
      disassociateRecord(metadata?.currentEntity?.EntitySetName, currentRecordId, relationshipName, id),
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  const filterSuggestions = useCallback(
    async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
      if (!filterText) return [];

      const results = await filterQuery.mutateAsync(filterText);
      return (
        results.map(
          (i) =>
            ({
              key: i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""],
              name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
              data: i,
            } as ITagWithData)
        ) ?? []
      );
    },
    [metadata?.associatedEntity.EntitySetName]
  );

  const showAllSuggestions = useCallback(
    async (selectedTags?: ITag[]): Promise<ITag[]> => {
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
    },
    [suggestions, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onPickerChange = useCallback(
    (selectedTags?: ITag[]): void => {
      const removed = selectedItems
        ?.filter(
          (i) =>
            !selectedTags?.some(
              (t) => (t.key as string)?.split(":")?.at(1) === i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            )
        )
        .map((i) => i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]);

      const added = selectedTags?.filter((t) => (t.key as string).indexOf(":") === -1).map((t) => t.key);

      added?.forEach((id) => associateQuery.mutate(id as string));
      removed?.forEach((id) => disassociateQuery.mutate(id as string));
    },
    [selectedItems, metadata?.associatedEntity.PrimaryIdAttribute]
  );

  const onItemSelected = useCallback(
    (item?: ITag): ITag | null => {
      if (item && !selectedItems?.some((i) => i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""] === item.key)) {
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

  const isDataLoading = isLoadingMetadata || isLoadingSuggestions || isLoadingSelectedItems;

  return (
    <TagPicker
      ref={pickerRef}
      selectedItems={selectedItems?.map(
        (i) =>
          ({
            key: `${i[metadata?.intersectEntity.PrimaryIdAttribute ?? ""]}:${
              i[metadata?.associatedEntity.PrimaryIdAttribute ?? ""]
            }`,
            name: i[metadata?.associatedEntity.PrimaryNameAttribute ?? ""],
          } as ITag)
      )}
      onResolveSuggestions={filterSuggestions}
      onEmptyResolveSuggestions={showAllSuggestions}
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
      disabled={disabled}
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
        placeholder: isDataLoading
          ? "Loading"
          : selectedItems?.length || disabled
          ? ""
          : `Select ${metadata?.associatedEntity.DisplayCollectionName.UserLocalizedLabel.Label ?? "an item"}`,
      }}
      pickerCalloutProps={{
        calloutMaxWidth: 500,
      }}
      itemLimit={itemLimit}
      onValidateInput={onCreateNew}
    />
  );
};

export default function PolyLookupControl(props: PolyLookupProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}
