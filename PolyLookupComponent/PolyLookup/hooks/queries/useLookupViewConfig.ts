import { useQuery } from "@tanstack/react-query";
import { getLookupViewConfig } from "../../services/DataverseService";

export function useLookupViewConfig(
  associatedEntityName: string | undefined,
  primaryIdAttribute: string | undefined,
  primaryNameAttribute: string | undefined,
  lookupViewValue: string | undefined
) {
  return useQuery({
    queryKey: ["lookupViewConfig", associatedEntityName, lookupViewValue],
    queryFn: () => getLookupViewConfig(associatedEntityName, primaryIdAttribute, primaryNameAttribute, lookupViewValue),
    enabled: !!associatedEntityName && !!primaryIdAttribute && !!primaryNameAttribute,
  });
}
