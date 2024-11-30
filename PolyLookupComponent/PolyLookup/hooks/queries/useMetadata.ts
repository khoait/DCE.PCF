import { useQuery } from "@tanstack/react-query";
import { getMetadata } from "../../services/DataverseService";

export function useMetadata(currentTable: string, relationshipName: string, relationship2Name: string | undefined) {
  return useQuery({
    queryKey: ["metadata", currentTable, relationshipName, relationship2Name],
    queryFn: () => getMetadata(currentTable, relationshipName, relationship2Name),
    enabled: !!currentTable && !!relationshipName,
  });
}
