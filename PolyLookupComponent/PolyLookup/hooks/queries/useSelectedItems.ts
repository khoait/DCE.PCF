import { useQuery } from "@tanstack/react-query";
import { retrieveAssociatedRecords } from "../../services/DataverseService";
import { IMetadata } from "../../types/metadata";

export function useSelectedItems(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  formType: XrmEnum.FormType | undefined,
  firstViewColumn?: string,
  selectedItemTemplate?: string | null,
  hasSelectedItemsView?: boolean,
  selectedItemsViewFetchXml?: string
) {
  const filterFetchXml = hasSelectedItemsView ? selectedItemsViewFetchXml : undefined;

  return useQuery({
    queryKey: ["selectedItems", metadata, currentRecordId, formType, firstViewColumn, selectedItemTemplate, filterFetchXml],
    queryFn: () => {
      if (!metadata || !currentRecordId || formType === XrmEnum.FormType.Create) {
        return [];
      }

      return retrieveAssociatedRecords(metadata, currentRecordId, firstViewColumn, selectedItemTemplate, filterFetchXml);
    },
    enabled:
      !!metadata?.intersectEntity.EntitySetName &&
      !!metadata?.associatedEntity.EntitySetName &&
      (!hasSelectedItemsView || !!selectedItemsViewFetchXml),
  });
}
