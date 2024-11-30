import { useQuery } from "@tanstack/react-query";
import { retrieveAssociatedRecords } from "../../services/DataverseService";
import { IMetadata } from "../../types/metadata";

export function useSelectedItems(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  formType: XrmEnum.FormType | undefined,
  firstViewColumn?: string,
  selectedItemTemplate?: string | null
) {
  return useQuery({
    queryKey: ["selectedItems", metadata, currentRecordId, formType, firstViewColumn, selectedItemTemplate],
    queryFn: () => {
      if (!metadata || !currentRecordId || formType === XrmEnum.FormType.Create) {
        return [];
      }

      return retrieveAssociatedRecords(metadata, currentRecordId, firstViewColumn, selectedItemTemplate);
    },
    enabled: !!metadata?.intersectEntity.EntitySetName && !!metadata?.associatedEntity.EntitySetName,
  });
}
