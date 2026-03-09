import { useQueryClient } from "@tanstack/react-query";
import { useEffect } from "react";
import { getHandlebarsVariables } from "../services/TemplateService";

export function useAttributeOnChange(fetchXml: string) {
  const queryClient = useQueryClient();

  useEffect(() => {
    const invalidateFetchDataFn = () => {
      queryClient.invalidateQueries({ queryKey: ["entityOptions"] });
    };

    const templateVar = getHandlebarsVariables(fetchXml);
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
  }, [fetchXml]);
}
