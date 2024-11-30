import { useQuery } from "@tanstack/react-query";
import { getHandlebarsVariables, getFetchTemplateString } from "../../services/TemplateService";
import { IMetadata } from "../../types/metadata";
import { ShowIconOptions, IconSizes } from "../../types/typings";
import { getEntityRecords } from "../../services/DataverseService";

export function useEntityOptions(
  metadata: IMetadata | undefined,
  customFilter?: string | null,
  groupBy?: string | null,
  optionTemplate?: string | null,
  selectedItemTemplate?: string | null,
  iconOptions?: ShowIconOptions,
  iconSize?: IconSizes
) {
  const entitySetName = metadata?.lookupEntity.EntitySetName ?? "";
  const lookupViewFetchXml = metadata?.lookupView.fetchxml ?? "";

  const entityIcon =
    metadata?.lookupEntity.IconVectorName ??
    (iconSize === IconSizes.Large
      ? metadata?.lookupEntity.IconMediumName ?? metadata?.lookupEntity.IconSmallName
      : metadata?.lookupEntity.IconSmallName);

  const recordImageUrlTemplate = metadata?.lookupEntity.RecordImageUrlTemplate;

  let iconTemplate = "";
  if (iconOptions === ShowIconOptions.RecordImage) {
    iconTemplate = recordImageUrlTemplate ?? "";
  } else if (iconOptions === ShowIconOptions.EntityIcon) {
    iconTemplate = entityIcon ?? "";
  }

  return useQuery({
    queryKey: [
      "entityRecords",
      entitySetName,
      lookupViewFetchXml,
      customFilter,
      groupBy,
      optionTemplate,
      selectedItemTemplate,
    ],
    queryFn: () => {
      const templateColumns: string[] = [];
      if (optionTemplate || selectedItemTemplate) {
        templateColumns.push(...getHandlebarsVariables(optionTemplate ?? "" + " " + selectedItemTemplate ?? ""));
      }
      const populatedFetchXml = getFetchTemplateString(lookupViewFetchXml ?? "", customFilter, templateColumns);
      return getEntityRecords(
        entitySetName,
        metadata?.lookupEntity.PrimaryIdAttribute ?? "",
        metadata?.lookupEntity.PrimaryNameAttribute ?? "",
        populatedFetchXml,
        groupBy,
        optionTemplate,
        selectedItemTemplate,
        iconOptions,
        iconTemplate,
        iconSize
      );
    },
    enabled: !!entitySetName && !!lookupViewFetchXml,
  });
}
