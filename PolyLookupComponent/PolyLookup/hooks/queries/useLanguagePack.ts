import { useQuery } from "@tanstack/react-query";
import { LanguagePack } from "../../types/languagePack";
import { getLanguagePack } from "../../services/DataverseService";

export function useLanguagePack(webResourcePath: string | undefined, defaultLanguagePack: LanguagePack) {
  return useQuery({
    queryKey: ["languagePack", webResourcePath],
    queryFn: () => getLanguagePack(webResourcePath, defaultLanguagePack),
    enabled: !!webResourcePath,
  });
}
