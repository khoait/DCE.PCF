import { useQuery } from "@tanstack/react-query";
import { getEntityOptions } from "../../services/DataverseService";
import { IMetadata, LookupView } from "../../types/metadata";

export function useEntityOptions(
  metadata: IMetadata | undefined,
  lookupViewConfig: LookupView | undefined,
  searchText?: string,
  pageSize?: number
) {
  return useQuery({
    queryKey: ["entityOptions", lookupViewConfig, searchText],
    queryFn: () => getEntityOptions(metadata, lookupViewConfig, searchText, pageSize),
    enabled: !!lookupViewConfig?.fetchXml,
  });
}
