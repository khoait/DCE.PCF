import { useQueryClient } from "@tanstack/react-query";
import { useEffect } from "react";
import { getHandlebarsVariables } from "../services/TemplateService";
import { IMetadata } from "../types/metadata";

export function useAttributeOnChange(metadata: IMetadata | undefined, customFilter?: string | null) {
  const queryClient = useQueryClient();

  useEffect(() => {
    const invalidateFetchDataFn = () => {
      queryClient.invalidateQueries({ queryKey: ["entityRecords"] });
    };

    const templateVar = getHandlebarsVariables((metadata?.lookupView.fetchxml ?? "") + (customFilter ?? ""));
    if (templateVar.length) {
      templateVar.forEach((v) => {
        Xrm.Page.getAttribute(v)?.addOnChange(invalidateFetchDataFn);
      });
    }

    return () => {
      templateVar.forEach((v) => {
        Xrm.Page.getAttribute(v)?.removeOnChange(invalidateFetchDataFn);
      });
    };
  }, [metadata?.lookupView.fetchxml, customFilter]);
}
