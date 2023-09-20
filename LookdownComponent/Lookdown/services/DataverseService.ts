import axios from "axios";
import { useQuery } from "@tanstack/react-query";
import { IEntityDefinition, IMetadata, IViewDefinition, IViewLayout } from "../types/metadata";

const tableDefinitionColumns = [
  "LogicalName",
  "DisplayName",
  "PrimaryIdAttribute",
  "PrimaryNameAttribute",
  "EntitySetName",
  "DisplayCollectionName",
  "IsQuickCreateEnabled",
];

const viewDefinitionColumns = ["savedqueryid", "name", "fetchxml", "layoutjson", "querytype"];

const apiVersion = "9.2";

export function useMetadata(lookupTable: string | undefined, lookupViewId: string | undefined) {
  return useQuery({
    queryKey: [`${lookupTable}_${lookupViewId}_metadata`, { lookupTable, lookupViewId }],
    queryFn: () => getMetadata(lookupTable ?? "", lookupViewId ?? ""),
    enabled: !!lookupTable && !!lookupViewId,
  });
}

export function useFetchXmlData(
  entitySetName: string | undefined,
  lookupViewId: string | undefined,
  fetchXml: string | undefined,
  groupBy?: string | null
) {
  return useQuery({
    queryKey: [`${entitySetName}_${lookupViewId}_fetchdata`, { entitySetName, lookupViewId }],
    queryFn: () => {
      return retrieveMultipleFetch(entitySetName ?? "", fetchXml ?? "", groupBy);
    },
    enabled: !!entitySetName && !!lookupViewId && !!fetchXml,
  });
}

async function getMetadata(lookupTable: string, lookupViewId: string): Promise<IMetadata> {
  const [lookupEntity, lookupView] = await Promise.all([
    getEntityDefinition(lookupTable),
    getViewDefinition(lookupViewId),
  ]);

  return {
    lookupEntity,
    lookupView,
  };
}

async function getEntityDefinition(entityName: string): Promise<IEntityDefinition> {
  return !entityName
    ? Promise.reject(new Error("Invalid entity name"))
    : axios
        .get<IEntityDefinition>(`/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${entityName}')`, {
          params: {
            $select: tableDefinitionColumns.join(","),
          },
        })
        .then((res) => {
          return {
            ...res.data,
            DisplayNameLocalized: res.data.DisplayName.UserLocalizedLabel?.Label ?? "",
            DisplayCollectionNameLocalized: res.data.DisplayCollectionName.UserLocalizedLabel?.Label ?? "",
          } as IEntityDefinition;
        });
}

async function getViewDefinition(viewId: string) {
  if (!viewId) return Promise.reject(new Error("Invalid view id"));

  const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `savedqueryid eq '${viewId}'`,
      $select: viewDefinitionColumns.join(","),
    },
  });

  const view = result.data.value?.at(0);

  if (!view) return Promise.reject(new Error("View not found"));

  const layoutjson = view.layoutjson as unknown as string;
  const layout = JSON.parse(layoutjson) as IViewLayout;
  view.layoutjson = layout;
  return view;
}

async function retrieveMultipleFetch(entitySetName: string, fetchXml: string, groupBy?: string | null) {
  if (!entitySetName || !fetchXml) return Promise.reject(new Error("Invalid entity set name or fetchXml"));

  const doc = new DOMParser().parseFromString(fetchXml, "application/xml");
  doc.documentElement.setAttribute("count", "5000");

  if (groupBy && !doc.querySelector(`attribute[name="${groupBy}"]`)) {
    const entity = doc.querySelector("entity");
    const attribute = doc.createElement("attribute");
    attribute.setAttribute("name", groupBy);
    entity?.appendChild(attribute);
  }

  const newFetchXml = new XMLSerializer().serializeToString(doc);

  return axios
    .get<{ value: ComponentFramework.WebApi.Entity[] }>(`/api/data/v${apiVersion}/${entitySetName}`, {
      headers: {
        Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
      },
      params: {
        fetchXml: encodeURIComponent(newFetchXml),
      },
    })
    .then((res) => {
      return res.data.value;
    });
}
