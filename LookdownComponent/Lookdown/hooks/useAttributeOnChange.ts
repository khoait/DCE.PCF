import { useQueryClient } from "@tanstack/react-query";
import { useCallback, useEffect } from "react";
import { getHandlebarsVariables } from "../services/TemplateService";
import { IMetadata } from "../types/metadata";

export function useAttributeOnChange(metadata: IMetadata | undefined, customFilter?: string | null) {
  const queryClient = useQueryClient();
  const invalidateFetchDataFn = useCallback((context) => {
    queryClient.invalidateQueries({ queryKey: ["entityRecords"] });
  }, []);

  useEffect(() => {
    const templateVar = getHandlebarsVariables((metadata?.lookupView.fetchxml ?? "") + (customFilter ?? ""));
    if (templateVar.length) {
      templateVar.forEach((v) => {
        Xrm.Page.data.entity.attributes.get(v)?.addOnChange(invalidateFetchDataFn);
      });
    }

    // TODO: not sure why this is not working
    // return () => {
    //   console.log("cleanup");
    //   templateVar.forEach((v) => {
    //     Xrm.Page.data.entity.attributes.get(v)?.removeOnChange(invalidateFetchDataFn);
    //   });
    // };
  }, [metadata?.lookupView.fetchxml, customFilter]);
}
