import { useQuery } from "@tanstack/react-query";
import { getLanguagePack } from "../../services/DataverseService";
import { LanguagePack } from "../../types/languagePack";

export function useLanguagePack(webResourcePath: string | undefined, defaultLanguagePack: LanguagePack) {
  return useQuery({
    queryKey: ["languagePack", webResourcePath],
    queryFn: () => getLanguagePack(webResourcePath, defaultLanguagePack),
    enabled: !!webResourcePath,
  });
}
