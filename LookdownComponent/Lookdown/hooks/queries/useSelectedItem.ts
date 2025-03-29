import { useQuery } from "@tanstack/react-query";
import { IconSizes, ShowIconOptions } from "../../types/typings";
import { getHandlebarsVariables } from "../../services/TemplateService";
import { useMetadata } from "./useMetadata";
import { IMetadata } from "../../types/metadata";
import { getSelectedEntityOption } from "../../services/DataverseService";

export function getSelectedItemQueryOptions(
  entitySetName: string,
  selectedId: string | null | undefined,
  primaryIdAttribute: string,
  primaryNameAttribute: string,
  selectedItemTemplate?: string | null,
  iconOptions?: ShowIconOptions,
  iconTemplate?: string,
  iconSize?: IconSizes
) {
  return {
    queryKey: ["selectedItem", selectedId, selectedItemTemplate],
    queryFn: () =>
      getSelectedEntityOption(
        entitySetName,
        selectedId,
        primaryIdAttribute,
        primaryNameAttribute,
        selectedItemTemplate,
        iconOptions,
        iconTemplate,
        iconSize
      ),
    enabled: !!entitySetName,
    staleTime: 5000,
  };
}

export function useSelectedItem(
  metadata: IMetadata | undefined,
  selectedId: string | null | undefined,
  selectedItemTemplate?: string | null,
  iconOptions?: ShowIconOptions,
  iconSize?: IconSizes
) {
  const entitySetName = metadata?.lookupEntity.EntitySetName ?? "";
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

  return useQuery(
    getSelectedItemQueryOptions(
      entitySetName,
      selectedId,
      metadata?.lookupEntity.PrimaryIdAttribute ?? "",
      metadata?.lookupEntity.PrimaryNameAttribute ?? "",
      selectedItemTemplate,
      iconOptions,
      iconTemplate,
      iconSize
    )
  );
}
