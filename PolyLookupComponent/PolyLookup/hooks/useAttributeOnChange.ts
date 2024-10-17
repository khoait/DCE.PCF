import { useQueryClient } from "@tanstack/react-query";
import { useCallback, useEffect } from "react";
import { IMetadata } from "../types/metadata";
import { getHandlebarsVariables } from "../services/TemplateService";

export function useAttributeOnChange(fetchXml: string) {
  const queryClient = useQueryClient();
  const invalidateFetchDataFn = useCallback(() => {
    queryClient.invalidateQueries({ queryKey: ["entityOptions"] });
  }, []);

  useEffect(() => {
    if (fetchXml) {
      const templateVar = getHandlebarsVariables(fetchXml);
      if (templateVar.length) {
        templateVar.forEach((v) => {
          Xrm.Page.getAttribute(v)?.addOnChange(invalidateFetchDataFn);
        });
      }
    }

    // TODO: not sure why this is not working
    // return () => {
    //   console.log("cleanup");
    //   templateVar.forEach((v) => {
    //     Xrm.Page.data.entity.attributes.get(v)?.removeOnChange(invalidateFetchDataFn);
    //   });
    // };
  }, [fetchXml]);
}
