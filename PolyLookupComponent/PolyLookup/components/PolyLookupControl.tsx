import React, { useCallback, useEffect, useState } from "react";
import {
  IBasePickerStyles,
  IBasePickerSuggestionsProps,
  ITag,
  ITagItemProps,
  TagPicker,
  TagPickerBase,
} from "@fluentui/react";
import { QueryClient, QueryClientProvider, useMutation, useQuery } from "@tanstack/react-query";
import {
  associateRecord,
  disassociateRecord,
  getEntityDefinition,
  getManytoManyRelationShipDefinition,
  retrieveAssociatedRecords,
  retrieveMultiple,
} from "../services/DataverseService";

const queryClient = new QueryClient();

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
  sortBy?: string;
  pageSize?: number;
  disabled?: boolean;
  onChange?: (value: string | undefined) => void;
}

const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
  noResultsFoundText: "No records found",
};

export const Body = ({
  currentTable,
  currentRecordId,
  relationshipName,
  relationshipType,
  clientUrl,
  sortBy,
  pageSize,
  disabled,
  onChange,
}: PolyLookupProps) => {
  const pickerStyles: Partial<IBasePickerStyles> = {
    root: { backgroundColor: "#fff", width: "100%" },
    input: { minWidth: "0", display: disabled ? "none" : "inline-block" },
    text: {
      minWidth: "0",
      borderColor: "transparent",
      "&:hover": { borderColor: disabled ? "transparent" : "#000" },
      "&:after": { backgroundColor: "transparent" },
      "&:hover:after": { backgroundColor: disabled ? "rgba(50, 50, 50, 0.1)" : "transparent" },
    },
  };

  const { data: relationshipDefinition } = useQuery({
    queryKey: ["relationshipDefinition", { currentTable, relationshipName }],
    queryFn: () => getManytoManyRelationShipDefinition(currentTable, relationshipName),
  });

  const associatedTable =
    relationshipDefinition?.Entity1LogicalName === currentTable
      ? relationshipDefinition?.Entity2LogicalName
      : relationshipDefinition?.Entity1LogicalName;

  const associatedIntesectAttribute =
    relationshipDefinition?.Entity1LogicalName === currentTable
      ? relationshipDefinition?.Entity2IntersectAttribute
      : relationshipDefinition?.Entity1IntersectAttribute;

  const currentIntesectAttribute =
    relationshipDefinition?.Entity1LogicalName === currentTable
      ? relationshipDefinition?.Entity1IntersectAttribute
      : relationshipDefinition?.Entity2IntersectAttribute;

  // get current table definition
  const { data: currentTableDefinition } = useQuery({
    queryKey: ["currentTableDefinition", currentTable],
    queryFn: () => getEntityDefinition(currentTable),
    enabled: !!relationshipDefinition,
  });

  // get associated table definition
  const { data: associatedTableDefinition, isSuccess: isAssociatedTableDefinitionSuccess } = useQuery({
    queryKey: ["associatedTableDefinition", associatedTable],
    queryFn: () => getEntityDefinition(associatedTable),
    enabled: !!relationshipDefinition,
  });

  const associatedTableSetName = associatedTableDefinition?.EntitySetName ?? "";

  if (isAssociatedTableDefinitionSuccess && associatedTableDefinition) {
    pickerSuggestionsProps.suggestionsHeaderText = `Suggested ${associatedTableDefinition?.DisplayCollectionName.UserLocalizedLabel.Label}`;
    pickerSuggestionsProps.noResultsFoundText = `No ${associatedTableDefinition?.DisplayCollectionName.UserLocalizedLabel.Label} found`;
  }

  // get top 20 suggestions from associated table
  const { data: suggestionItems } = useQuery({
    queryKey: ["suggestionItems", associatedTable],
    queryFn: () =>
      retrieveMultiple(
        associatedTableSetName,
        [associatedTableDefinition?.PrimaryIdAttribute, associatedTableDefinition?.PrimaryNameAttribute],
        undefined,
        pageSize ?? 20,
        sortBy
      ),
    enabled: !!associatedTableDefinition,
  });

  // get intersect table definition
  const { data: intersectTableDefinition } = useQuery({
    queryKey: ["intersectTableDefinition", relationshipDefinition?.IntersectEntityName],
    queryFn: () => getEntityDefinition(relationshipDefinition?.IntersectEntityName),
    enabled: !!relationshipDefinition?.IntersectEntityName,
  });

  // get selected items
  const {
    data: selectedItems,
    refetch: selectedItemsRefetch,
    isSuccess,
  } = useQuery({
    queryKey: ["selectedItems", { currentTable, relationshipName }],
    queryFn: () =>
      retrieveAssociatedRecords(
        currentRecordId,
        intersectTableDefinition?.LogicalName,
        intersectTableDefinition?.EntitySetName,
        intersectTableDefinition?.PrimaryIdAttribute,
        currentIntesectAttribute,
        associatedIntesectAttribute,
        associatedTable,
        associatedTableDefinition?.PrimaryIdAttribute,
        associatedTableDefinition?.PrimaryNameAttribute
      ),
    enabled: !!intersectTableDefinition,
  });

  if (isSuccess && onChange) {
    onChange(selectedItems?.map((i) => i[associatedTableDefinition?.PrimaryNameAttribute ?? ""] as string).join(", "));
  }

  // filter query
  const filterQuery = useMutation({
    mutationFn: (searchText: string) => {
      return retrieveMultiple(
        associatedTableSetName,
        [associatedTableDefinition?.PrimaryIdAttribute, associatedTableDefinition?.PrimaryNameAttribute],
        `startswith(${associatedTableDefinition?.PrimaryNameAttribute}, '${searchText}')`
      );
    },
  });

  // associate query
  const associateQuery = useMutation({
    mutationFn: (id: string) =>
      associateRecord(
        currentTableDefinition?.EntitySetName,
        currentRecordId,
        associatedTableDefinition?.EntitySetName,
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
      disassociateRecord(currentTableDefinition?.EntitySetName, currentRecordId, relationshipName, id),
    onSuccess: (data, variables, context) => {
      selectedItemsRefetch();
    },
  });

  const filterSuggestions = async (filterText: string, selectedTag?: ITag[]): Promise<ITag[]> => {
    if (!filterText) return [];

    const results = await filterQuery.mutateAsync(filterText);

    return (
      results.map(
        (i) =>
          ({
            key: i[associatedTableDefinition?.PrimaryIdAttribute ?? ""],
            name: i[associatedTableDefinition?.PrimaryNameAttribute ?? ""],
          } as ITag)
      ) ?? []
    );
  };

  const showAllSuggestions = async (selectedTags?: ITag[]): Promise<ITag[]> => {
    return (
      suggestionItems?.map(
        (i) =>
          ({
            key: i[associatedTableDefinition?.PrimaryIdAttribute ?? ""] ?? "",
            name: i[associatedTableDefinition?.PrimaryNameAttribute ?? ""] ?? "",
          } as ITag)
      ) ?? []
    );
  };

  const onPickerChange = useCallback(
    (selectedTags?: ITag[]): void => {
      const removed = selectedItems
        ?.filter(
          (i) =>
            !selectedTags?.some(
              (t) => (t.key as string)?.split(":")?.at(1) === i[associatedTableDefinition?.PrimaryIdAttribute ?? ""]
            )
        )
        .map((i) => i[associatedTableDefinition?.PrimaryIdAttribute ?? ""]);

      const added = selectedTags?.filter((t) => (t.key as string).indexOf(":") === -1).map((t) => t.key);

      added?.forEach((id) => associateQuery.mutate(id as string));
      removed?.forEach((id) => disassociateQuery.mutate(id as string));

      const output = selectedTags?.map((t) => t.name).join(", ");
      if (onChange) {
        onChange(output);
      }
    },
    [selectedItems]
  );

  const onItemSelected = (item?: ITag): ITag | null => {
    if (item && !selectedItems?.some((i) => i[associatedTableDefinition?.PrimaryIdAttribute ?? ""] === item.key)) {
      return item;
    }
    return null;
  };

  return (
    <TagPicker
      selectedItems={selectedItems?.map(
        (i) =>
          ({
            key: `${i[intersectTableDefinition?.PrimaryIdAttribute ?? ""]}:${
              i[associatedTableDefinition?.PrimaryIdAttribute ?? ""]
            }`,
            name: i[associatedTableDefinition?.PrimaryNameAttribute ?? ""],
          } as ITag)
      )}
      onResolveSuggestions={filterSuggestions}
      onEmptyResolveSuggestions={showAllSuggestions}
      onChange={onPickerChange}
      onItemSelected={onItemSelected}
      styles={pickerStyles}
      pickerSuggestionsProps={pickerSuggestionsProps}
      disabled={disabled}
      onRenderItem={(props: ITagItemProps) => {
        if (disabled) {
          props.styles = { close: { display: "none" } };
        }
        return TagPickerBase.defaultProps.onRenderItem(props);
      }}
      resolveDelay={100}
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
