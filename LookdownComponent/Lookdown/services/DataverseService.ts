import { useQuery } from "@tanstack/react-query";
import axios from "axios";
import Handlebars from "handlebars";
import { LanguagePack } from "../types/languagePack";
import { IEntityDefinition, IMetadata, IViewDefinition, IViewLayout } from "../types/metadata";
import { EntityOption, IconSizes, ShowIconOptions } from "../types/typings";
import { getFetchTemplateString, getHandlebarsVariables } from "./TemplateService";

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

const DEFAULT_HEADERS = {
  "OData-MaxVersion": "4.0",
  "OData-Version": "4.0",
  Accept: "application/json",
  "Content-Type": "application/json; charset=utf-8",
  Prefer:
    'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.associatednavigationproperty,Microsoft.Dynamics.CRM.lookuplogicalname"',
};

export function getLanguagePack(
  webResourceUrl: string | undefined,
  defaultLanguagePack: LanguagePack
): Promise<LanguagePack> {
  const languagePack: LanguagePack = { ...defaultLanguagePack };

  if (webResourceUrl === undefined) return Promise.resolve(languagePack);

  return axios
    .get(webResourceUrl, {
      headers: DEFAULT_HEADERS,
    })
    .then((res) => {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(res.data, "text/xml");

      const nodes = xmlDoc.getElementsByTagName("data");
      for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        const key = node.getAttribute("name");
        const value = node.getElementsByTagName("value")[0].childNodes[0].nodeValue ?? "";
        if (key && value) {
          languagePack[key as keyof LanguagePack] = value;
        }
      }

      return languagePack;
    })
    .catch(() => languagePack);
}

export async function getMetadata(lookupTable: string, lookupViewId: string): Promise<IMetadata> {
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
          headers: DEFAULT_HEADERS,
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
      headers: DEFAULT_HEADERS,
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
    headers: DEFAULT_HEADERS,
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

export async function getEntityRecords(
  entitySetName: string,
  primaryIdAttribute: string,
  primaryNameAttribute: string,
  fetchXml: string,
  customFilter?: string | null,
  groupBy?: string | null,
  optionTemplate?: string | null,
  iconOptions?: ShowIconOptions,
  iconTemplate?: string | null,
  iconSize?: IconSizes
) {
  const templateColumns: string[] = [];
  if (primaryNameAttribute) {
    templateColumns.push(primaryNameAttribute);
  }
  if (optionTemplate) {
    templateColumns.push(...getHandlebarsVariables(optionTemplate ?? ""));
  }
  const populatedFetchXml = getFetchTemplateString(fetchXml ?? "", customFilter, templateColumns);

  const results = await retrieveMultipleFetch(entitySetName, populatedFetchXml, groupBy);

  const optionTemplateFn = optionTemplate ? Handlebars.compile(optionTemplate) : null;

  const iconTemplateFn =
    iconTemplate && iconOptions === ShowIconOptions.RecordImage ? Handlebars.compile(iconTemplate) : null;

  return results.map((record) => {
    let iconSrc = "";
    if (iconOptions === ShowIconOptions.EntityIcon) {
      iconSrc = iconTemplate ?? "";
    } else if (iconOptions === ShowIconOptions.RecordImage) {
      iconSrc = iconTemplateFn?.(record) ?? "";
    }
    return {
      id: record[primaryIdAttribute],
      primaryName: record[primaryNameAttribute],
      optionText: optionTemplateFn?.(record) ?? record[primaryNameAttribute],
      iconSrc,
      iconSize: iconSize === IconSizes.Large ? 32 : 16,
      group: record[groupBy ?? ""],
    } as EntityOption;
  });
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
      headers: DEFAULT_HEADERS,
      params: {
        fetchXml: encodeURIComponent(newFetchXml),
      },
    })
    .then((res) => {
      // overwrite with formatted values
      res.data.value.forEach((record) => {
        Object.keys(record).forEach((key) => {
          if (!key.endsWith("@OData.Community.Display.V1.FormattedValue")) {
            record[key] = getAttributeFormattedValue(record, key);
          }
        });
      });

      return res.data.value;
    });
}

export async function getSelectedEntityOption(
  entitySetName: string,
  entityId: string | null | undefined,
  primaryIdAttribute: string,
  primaryNameAttribute: string,
  optionTemplate?: string | null,
  iconOptions?: ShowIconOptions,
  iconTemplate?: string | null,
  iconSize?: IconSizes
) {
  if (!entityId) return Promise.resolve(null);

  const templateColumns: string[] = [];
  if (primaryNameAttribute) {
    templateColumns.push(primaryNameAttribute);
  }
  if (optionTemplate) {
    templateColumns.push(...getHandlebarsVariables(optionTemplate));
  }

  const record = await retrieveRecord(entitySetName, entityId, templateColumns);

  if (!record) return null;

  const optionTemplateFn = optionTemplate ? Handlebars.compile(optionTemplate) : null;

  const iconTemplateFn =
    iconTemplate && iconOptions === ShowIconOptions.RecordImage ? Handlebars.compile(iconTemplate) : null;

  let iconSrc = "";
  if (iconOptions === ShowIconOptions.EntityIcon) {
    iconSrc = iconTemplate ?? "";
  } else if (iconOptions === ShowIconOptions.RecordImage) {
    iconSrc = iconTemplateFn?.(record) ?? "";
  }
  return {
    id: record[primaryIdAttribute],
    primaryName: record[primaryNameAttribute],
    optionText: optionTemplateFn?.(record) ?? record[primaryNameAttribute],
    iconSrc,
    iconSize: iconSize === IconSizes.Large ? 32 : 16,
    group: "",
  } as EntityOption;
}

export async function retrieveRecord(entitySetName: string, entityId: string, columns: string[]) {
  if (!entitySetName || !entityId) return Promise.reject(new Error("Invalid entity set name or entityId"));

  return axios
    .get<ComponentFramework.WebApi.Entity>(`/api/data/v${apiVersion}/${entitySetName}(${entityId})`, {
      headers: DEFAULT_HEADERS,
      params: {
        $select: columns.join(","),
      },
    })
    .then((res) => {
      // overwrite with formatted values
      Object.keys(res.data).forEach((key) => {
        if (!key.endsWith("@OData.Community.Display.V1.FormattedValue")) {
          res.data[key] = getAttributeFormattedValue(res.data, key);
        }
      });

      return res.data;
    });
}

export function getCurrentRecord(): ComponentFramework.WebApi.Entity | null {
  if (!Xrm.Page?.data?.entity) return null;

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

export function getAttributeFormattedValue(entity: ComponentFramework.WebApi.Entity, attributeName: string): string {
  if (!entity) return "";

  const formattedValue1 = entity[`_${attributeName}_value@OData.Community.Display.V1.FormattedValue`];
  const formattedValue2 = entity[`${attributeName}@OData.Community.Display.V1.FormattedValue`];

  const displayValue = formattedValue1 ?? formattedValue2 ?? entity[attributeName] ?? "";

  return displayValue;
}
