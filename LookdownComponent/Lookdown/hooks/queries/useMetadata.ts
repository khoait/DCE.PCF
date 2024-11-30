import { useQuery } from "@tanstack/react-query";
import { getMetadata } from "../../services/DataverseService";

export function useMetadata(lookupTable: string, lookupViewId: string) {
  return useQuery({
    queryKey: ["metadata", lookupTable, lookupViewId],
    queryFn: () => getMetadata(lookupTable, lookupViewId),
    enabled: !!lookupTable && !!lookupViewId,
  });
}
