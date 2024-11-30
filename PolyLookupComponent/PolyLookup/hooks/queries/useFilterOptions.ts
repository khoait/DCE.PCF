import { useMutation } from "@tanstack/react-query";
import { getEntityOptions } from "../../services/DataverseService";
import { IMetadata, LookupView } from "../../types/metadata";

export function useFilterOptions(metadata: IMetadata | undefined, lookupViewConfig: LookupView | undefined) {
  return useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText?: string; pageSizeParam?: number }) =>
      getEntityOptions(metadata, lookupViewConfig, searchText, pageSizeParam),
  });
}
