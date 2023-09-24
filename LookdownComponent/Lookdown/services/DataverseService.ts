import axios from "axios";
import { useQuery } from "@tanstack/react-query";
import { IEntityDefinition, IMetadata, IViewDefinition, IViewLayout } from "../types/metadata";
import Handlebars from "handlebars";

const tableDefinitionColumns = [
  "LogicalName",
  "DisplayName",
  "PrimaryIdAttribute",
  "PrimaryNameAttribute",
  "EntitySetName",
  "DisplayCollectionName",
  "IsQuickCreateEnabled",
  "IconVectorName",
  "IconSmallName",
  "IconMediumName",
  "PrimaryImageAttribute",
];

const viewDefinitionColumns = ["savedqueryid", "name", "fetchxml", "layoutjson", "querytype"];

const apiVersion = "9.2";

export function useMetadata(
  lookupTable: string | null | undefined,
  lookupViewId: string | null | undefined,
  entityIcon = false
) {
  return useQuery({
    queryKey: [`${lookupTable}_${lookupViewId}_metadata`, { lookupTable, lookupViewId }],
    queryFn: () => getMetadata(lookupTable ?? "", lookupViewId ?? ""),
    enabled: !!lookupTable && !!lookupViewId,
  });
}

export function useFetchXmlData(
  entitySetName: string | null | undefined,
  lookupViewId: string | null | undefined,
  fetchXml: string | undefined,
  customFilter?: string | null,
  groupBy?: string | null
) {
  return useQuery({
    queryKey: [`${entitySetName}_${lookupViewId}_fetchdata`, { entitySetName, lookupViewId, fetchXml }],
    queryFn: () => {
      return retrieveMultipleFetch(entitySetName ?? "", fetchXml ?? "", groupBy);
    },
    enabled: !!entitySetName && !!lookupViewId && !!fetchXml,
  });
}

async function getMetadata(lookupTable: string, lookupViewId: string): Promise<IMetadata> {
  const [lookupEntity, lookupView] = await Promise.all([
    getEntityDefinition(lookupTable),
    getViewDefinition(lookupTable, lookupViewId),
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
          const {
            EntitySetName,
            PrimaryIdAttribute,
            DisplayName,
            DisplayCollectionName,
            IconVectorName,
            IconSmallName,
            IconMediumName,
            PrimaryImageAttribute,
          } = res.data;
          return {
            ...res.data,
            DisplayNameLocalized: DisplayName.UserLocalizedLabel?.Label ?? "",
            DisplayCollectionNameLocalized: DisplayCollectionName.UserLocalizedLabel?.Label ?? "",
            IconVectorName: IconVectorName ? `/WebResources/${IconVectorName}` : undefined,
            IconSmallName: IconSmallName ? `/WebResources/${IconSmallName}` : undefined,
            IconMediumName: IconMediumName ? `/WebResources/${IconMediumName}` : undefined,
            RecordImageUrlTemplate: PrimaryImageAttribute
              ? `/api/data/v${apiVersion}/${EntitySetName}({{${PrimaryIdAttribute}}})/${PrimaryImageAttribute}/$value`
              : undefined,
          } as IEntityDefinition;
        });
}

async function getViewDefinition(entityName: string, viewId?: string | null): Promise<IViewDefinition> {
  let view;
  if (viewId) {
    const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
      params: {
        $filter: `savedqueryid eq '${viewId}'`,
        $select: viewDefinitionColumns.join(","),
      },
    });
    view = result.data.value?.at(0);
    if (view) {
      const layoutjson = view.layoutjson as unknown as string;
      const layout = JSON.parse(layoutjson) as IViewLayout;
      view.layoutjson = layout;
    }
  }

  if (!view) {
    view = await getLookupViewDefinition(entityName);
  }

  if (!view) return Promise.reject(new Error("View not found"));

  return view;
}

export async function getLookupViewDefinition(entityName: string) {
  if (!entityName) return Promise.reject(new Error("Invalid entity name"));

  const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `returnedtypecode eq '${entityName}' and querytype eq 64`,
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

export function getCurrentRecord(): ComponentFramework.WebApi.Entity {
  return Object.fromEntries(
    /* global Xrm */
    Xrm.Page.data.entity.attributes.get().map((attribute) => {
      const attributeType = attribute.getAttributeType();
      if (attributeType === "lookup") {
        const lookupVal = (attribute as Xrm.Attributes.LookupAttribute).getValue()?.at(0);
        if (lookupVal) {
          lookupVal.id = lookupVal.id.replace("{", "").replace("}", "");
        }

        return [attribute.getName(), lookupVal];
      } else if (attributeType === "datetime") {
        return [attribute.getName(), (attribute as Xrm.Attributes.DateAttribute).getValue()?.toISOString()];
      }

      return [attribute.getName(), attribute.getValue()];
    })
  );
}
