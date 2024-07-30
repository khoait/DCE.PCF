import { useQuery } from "@tanstack/react-query";
import axios from "axios";
import { LanguagePack } from "../types/languagePack";
import {
  IEntityDefinition,
  IManyToManyRelationship,
  IMetadata,
  IOneToManyRelationship,
  isManyToMany,
  IViewDefinition,
  IViewLayout,
  LookupView,
} from "../types/metadata";

const nToNColumns = [
  "SchemaName",
  "Entity1LogicalName",
  "Entity2LogicalName",
  "Entity1IntersectAttribute",
  "Entity2IntersectAttribute",
  "RelationshipType",
  "IntersectEntityName",
];

const oneToNColumns = [
  "SchemaName",
  "ReferencedEntity",
  "ReferencingEntity",
  "ReferencingEntityNavigationPropertyName",
  "ReferencedAttribute",
  "ReferencingAttribute",
  "RelationshipType",
];

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

export function useMetadata(currentTable: string, relationshipName: string, relationship2Name: string | undefined) {
  return useQuery({
    queryKey: ["metadata", currentTable, relationshipName, relationship2Name],
    queryFn: () => getMetadata(currentTable, relationshipName, relationship2Name),
    enabled: !!currentTable && !!relationshipName,
  });
}

export function useLookupViewConfig(
  associatedEntityName: string | undefined,
  primaryIdAttribute: string | undefined,
  primaryNameAttribute: string | undefined,
  lookupViewValue: string | undefined
) {
  return useQuery({
    queryKey: ["lookupViewConfig", associatedEntityName, lookupViewValue],
    queryFn: () => getLookupViewConfig(associatedEntityName, primaryIdAttribute, primaryNameAttribute, lookupViewValue),
    enabled: !!associatedEntityName && !!primaryIdAttribute && !!primaryNameAttribute,
  });
}

export function useSelectedItems(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  formType: XrmEnum.FormType | undefined
) {
  return useQuery({
    queryKey: ["selectedItems", metadata, currentRecordId, formType],
    queryFn: () =>
      retrieveAssociatedRecords(
        currentRecordId,
        metadata?.intersectEntity.LogicalName,
        metadata?.intersectEntity.EntitySetName,
        metadata?.intersectEntity.PrimaryIdAttribute,
        metadata?.currentIntesectAttribute,
        metadata?.associatedIntesectAttribute,
        metadata?.associatedEntity.LogicalName,
        metadata?.associatedEntity.PrimaryIdAttribute,
        metadata?.associatedEntity.PrimaryNameAttribute
      ),
    enabled:
      !!metadata?.intersectEntity.EntitySetName &&
      !!metadata?.associatedEntity.EntitySetName &&
      (formType === XrmEnum.FormType.Update ||
        formType === XrmEnum.FormType.ReadOnly ||
        formType === XrmEnum.FormType.Disabled),
  });
}

export function useLanguagePack(webResourcePath: string | undefined, defaultLanguagePack: LanguagePack) {
  return useQuery({
    queryKey: ["languagePack", webResourcePath],
    queryFn: () => getLanguagePack(webResourcePath, defaultLanguagePack),
    enabled: !!webResourcePath,
  });
}

export function getLanguagePack(
  webResourceUrl: string | undefined,
  defaultLanguagePack: LanguagePack
): Promise<LanguagePack> {
  const languagePack: LanguagePack = { ...defaultLanguagePack };

  if (webResourceUrl === undefined) return Promise.resolve(languagePack);

  return axios
    .get(webResourceUrl)
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

export function getManytoManyRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<IManyToManyRelationship>(
          `/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${currentTable}')/ManyToManyRelationships(SchemaName='${relationshipName}')`,
          {
            params: {
              $select: nToNColumns.join(","),
            },
          }
        )
        .then((res) => res.data);
}

export function getOnetoManyRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<IOneToManyRelationship>(
          `/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${currentTable}')/OneToManyRelationships(SchemaName='${relationshipName}')`,
          {
            params: {
              $select: oneToNColumns.join(","),
            },
          }
        )
        .then((res) => res.data);
}

export function getManytoOneRelationShipDefinition(
  currentTable: string | undefined,
  relationshipName: string | undefined
) {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid table or relationship name"))
    : axios
        .get<{ value: IOneToManyRelationship[] }>(
          `/api/data/v${apiVersion}/RelationshipDefinitions/Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata`,
          {
            params: {
              $select: oneToNColumns.join(","),
              $filter: `SchemaName eq '${relationshipName}' and ReferencingEntity eq '${currentTable}'`,
            },
          }
        )
        .then((res) => res.data.value.at(0));
}

export function getEntityDefinition(entityName: string | undefined) {
  return typeof entityName === "undefined"
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

export async function getViewDefinition(
  entityName: string | undefined,
  viewName: string | undefined,
  queryType?: number | undefined
) {
  if (typeof entityName === "undefined") return Promise.reject(new Error("Invalid arguments"));

  const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `returnedtypecode eq '${entityName}' ${viewName ? `and name eq '${viewName}'` : ""} ${
        queryType ? `and querytype eq ${queryType}` : ""
      } `,
      $select: viewDefinitionColumns.join(","),
    },
  });

  if (!result.data?.value.length) {
    return null;
  }

  const layoutjson = result.data.value[0].layoutjson as unknown as string;
  const layout = JSON.parse(layoutjson) as IViewLayout;
  result.data.value[0].layoutjson = layout;
  return result.data.value[0];
}

export async function getDefaultView(entityName: string | undefined, viewName: string | undefined) {
  if (viewName) {
    const viewByName = await getViewDefinition(entityName, viewName);
    if (viewByName) return viewByName;
  }

  const defaultView = await getViewDefinition(entityName, undefined, 64);
  return defaultView;
}

async function getMetadata(
  currentTable: string | undefined,
  relationshipName: string | undefined,
  relationship2Name: string | undefined
): Promise<IMetadata> {
  if (typeof currentTable === "undefined" || typeof relationshipName === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  let relationship1: IManyToManyRelationship | IOneToManyRelationship;
  let relationship2: IOneToManyRelationship | undefined;
  let intersectEntityName = "";
  let associatedEntityName: string | undefined;

  if (typeof relationship2Name === "undefined") {
    relationship1 = await getManytoManyRelationShipDefinition(currentTable, relationshipName);
    intersectEntityName = relationship1.IntersectEntityName;
    associatedEntityName =
      relationship1.Entity1LogicalName === currentTable
        ? relationship1.Entity2LogicalName
        : relationship1.Entity1LogicalName;
  } else {
    relationship1 = await getOnetoManyRelationShipDefinition(currentTable, relationshipName);
    intersectEntityName = relationship1.ReferencingEntity;
    relationship2 = await getManytoOneRelationShipDefinition(intersectEntityName, relationship2Name);
    associatedEntityName = relationship2?.ReferencedEntity;
  }

  try {
    const [currentEntity, intersectEntity, associatedEntity] = await Promise.all([
      getEntityDefinition(currentTable),
      getEntityDefinition(intersectEntityName),
      getEntityDefinition(associatedEntityName),
    ]);

    let currentIntesectAttribute: string;
    let associatedIntesectAttribute: string;
    let currentEntityNavigationPropertyName: string | undefined;
    let associatedEntityNavigationPropertyName: string | undefined;
    if (isManyToMany(relationship1)) {
      currentIntesectAttribute =
        relationship1.Entity1LogicalName === currentTable
          ? relationship1.Entity1IntersectAttribute
          : relationship1.Entity2IntersectAttribute;

      associatedIntesectAttribute =
        relationship1.Entity1LogicalName === currentTable
          ? relationship1.Entity2IntersectAttribute
          : relationship1.Entity1IntersectAttribute;
    } else {
      currentIntesectAttribute = relationship1.ReferencingAttribute;
      associatedIntesectAttribute = relationship2?.ReferencingAttribute ?? "";
      currentEntityNavigationPropertyName = relationship1.ReferencingEntityNavigationPropertyName;
      associatedEntityNavigationPropertyName = relationship2?.ReferencingEntityNavigationPropertyName;
    }

    return {
      relationship1,
      relationship2,
      currentEntity,
      intersectEntity,
      associatedEntity,
      currentIntesectAttribute,
      associatedIntesectAttribute,
      currentEntityNavigationPropertyName,
      associatedEntityNavigationPropertyName,
    };
  } catch (error) {
    console.log(error);
    throw error;
  }
}

async function getLookupViewConfig(
  associatedEntityName: string | undefined,
  primaryIdAttribute: string | undefined,
  primaryNameAttribute: string | undefined,
  lookupViewValue: string | undefined
) {
  if (
    typeof associatedEntityName === "undefined" ||
    typeof primaryIdAttribute === "undefined" ||
    typeof primaryNameAttribute === "undefined"
  ) {
    throw new Error("Invalid arguments");
  }

  const lookupViewVal = lookupViewValue?.trim() ?? "";
  let lookupViewConfig: LookupView | undefined;

  // check if lookupViewValue is a fetchXml
  if (lookupViewVal.startsWith("<fetch")) {
    lookupViewConfig = {
      sourceType: "FetchXml",
      source: "FetchXml",
      fetchXml: lookupViewVal,
      columns: [],
      isSystemLookupView: false,
    };
  }
  // if lookupViewValue is an ODataUrl
  else if (lookupViewVal.startsWith("https://") || lookupViewVal.startsWith("/")) {
    lookupViewConfig = {
      sourceType: "ODataUrl",
      source: lookupViewVal,
      fetchXml: "",
      columns: [],
      isSystemLookupView: false,
    };

    const url =
      lookupViewVal.startsWith("https://") || lookupViewVal.startsWith("/api")
        ? lookupViewVal
        : `/api/data/v${apiVersion}/${lookupViewVal}`;

    try {
      const { data } = await axios.get<{ value?: unknown }>(url);

      if (typeof data.value === "string") {
        lookupViewConfig.fetchXml = data.value;
      }
    } catch (error) {
      if (axios.isAxiosError(error)) {
        if (error.response?.status === 404) {
          throw new Error("Invalid OData URL");
        }
      } else {
        throw new Error("Failed to get data from OData URL");
      }
    }
  } else {
    if (lookupViewVal) {
      // get environment variable in Dataverse
      const environmentVariableValue = await getEnvironmentVariableValue(lookupViewVal);
      if (environmentVariableValue) {
        lookupViewConfig = {
          sourceType: "EnvironmentVariable",
          source: lookupViewVal,
          fetchXml: environmentVariableValue,
          columns: [],
          isSystemLookupView: false,
        };
      }
    }

    if (!lookupViewConfig) {
      // environment variable not found
      // check if lookupViewValue is a view name
      const viewDef = await getDefaultView(associatedEntityName, lookupViewVal);
      if (viewDef) {
        lookupViewConfig = {
          sourceType: "ViewName",
          source: lookupViewVal,
          fetchXml: viewDef.fetchxml,
          columns: viewDef.layoutjson.Rows[0].Cells.map((c) => c.Name),
          isSystemLookupView: viewDef.querytype === 64,
        };
      }
    }
  }

  if (!lookupViewConfig?.fetchXml) {
    throw new Error("Invalid Lookup View Configuration");
  }

  // validate fetchXml
  let fetchXml = lookupViewConfig.fetchXml;

  if (!lookupViewConfig.isSystemLookupView) {
    // match "{{...}}" pattern in the string
    const regex = /{{[^}]*}}/g;
    fetchXml = fetchXml.replace(regex, (match) => {
      // Remove "[_]" from the matched string
      return match.replace(/\[_\]/g, "_");
    });
  }

  const doc = new DOMParser().parseFromString(fetchXml, "application/xml");
  if (doc.getElementsByTagName("fetch").length === 0) {
    throw new Error("Invalid FetchXml");
  }

  // validate entity name in fetchXml with associatedEntityName
  const entityEl = doc.getElementsByTagName("entity")[0];
  const entityName = entityEl.getAttribute("name");
  if (entityName !== associatedEntityName) {
    throw new Error("Lookup View entity does not match the associated entity");
  }

  const attributes = Array.from(doc.getElementsByTagName("attribute"));
  // extract columns from fetchXml
  if (lookupViewConfig.sourceType !== "ViewName") {
    lookupViewConfig.columns = attributes.map((attr) => attr.getAttribute("name") ?? "").filter((attr) => attr !== "");
  }

  // check if doc has attribute with name equals primaryIdAttribute and primaryNameAttribute
  const hasPrimaryIdAttribute = attributes.some((attr) => attr.getAttribute("name") === primaryIdAttribute);
  const hasPrimaryNameAttribute = attributes.some((attr) => attr.getAttribute("name") === primaryNameAttribute);

  if (!hasPrimaryIdAttribute) {
    const primaryIdAttr = doc.createElement("attribute");
    primaryIdAttr.setAttribute("name", primaryIdAttribute);
    entityEl.appendChild(primaryIdAttr);
  }

  if (!hasPrimaryNameAttribute) {
    const primaryNameAttr = doc.createElement("attribute");
    primaryNameAttr.setAttribute("name", primaryNameAttribute);
    entityEl.appendChild(primaryNameAttr);
  }

  if (!hasPrimaryIdAttribute || !hasPrimaryNameAttribute) {
    lookupViewConfig.fetchXml = new XMLSerializer().serializeToString(doc);
  }

  return lookupViewConfig;
}

export function retrieveMultiple(
  entitySetName: string | undefined,
  columns: Array<string | undefined> | undefined,
  filter?: string,
  top?: number,
  orderby?: string,
  extra?: string
) {
  return typeof entitySetName === "undefined" ||
    typeof columns === "undefined" ||
    columns.some((c) => typeof c === "undefined")
    ? Promise.reject(new Error("Invalid entity set name or columns"))
    : axios
        .get<{ value: ComponentFramework.WebApi.Entity[] }>(`/api/data/v${apiVersion}/${entitySetName}`, {
          params: {
            $select: columns.join(","),
            $filter: filter,
            $top: top,
            $orderby: orderby,
          },
        })
        .then((res) => res.data.value);
}

export async function retrieveMultipleFetch(
  entitySetName: string | undefined,
  fetchXml: string | undefined,
  page?: number,
  count?: number,
  pagingCookies?: string
) {
  if (typeof entitySetName === "undefined" || typeof fetchXml === "undefined") {
    throw new Error("Invalid entity set name or fetchXml");
  }

  try {
    const doc = new DOMParser().parseFromString(fetchXml, "application/xml");
    doc.documentElement.setAttribute("page", page?.toString() ?? "1");
    doc.documentElement.setAttribute("count", count?.toString() ?? "50");
    if (pagingCookies) {
      doc.documentElement.setAttribute("paging-cookie", pagingCookies);
    }
    const newFetchXml = new XMLSerializer().serializeToString(doc);

    const { data } = await axios.get<{ value: ComponentFramework.WebApi.Entity[] }>(
      `/api/data/v${apiVersion}/${entitySetName}`,
      {
        headers: {
          Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
        },
        params: {
          fetchXml: encodeURIComponent(newFetchXml),
        },
      }
    );

    return data.value ?? [];
  } catch (error) {
    // handle serialization error
    if (error instanceof DOMException || error instanceof TypeError || error instanceof SyntaxError) {
      throw new Error("Invalid FetchXml");
    }
    // handle axios error
    if (axios.isAxiosError(error)) {
      if (error.response?.status === 400) {
        throw new Error("Invalid FetchXml");
      }
    }
    throw new Error("Failed to retrieve data");
  }
}

export async function retrieveAssociatedRecords(
  currentRecordId: string | undefined,
  intersectEntity: string | undefined,
  intersectEntitySet: string | undefined,
  intersectPrimaryIdAttribute: string | undefined,
  currentIntersectAttribute: string | undefined,
  associatedIntersectAttribute: string | undefined,
  associatedEntity: string | undefined,
  associatedEntityPrimaryIdAttribute: string | undefined,
  associatedEntityPrimaryNameAttribute: string | undefined
) {
  if (
    typeof currentRecordId === "undefined" ||
    typeof intersectEntity === "undefined" ||
    typeof intersectEntitySet === "undefined" ||
    typeof intersectPrimaryIdAttribute === "undefined" ||
    typeof currentIntersectAttribute === "undefined" ||
    typeof associatedIntersectAttribute === "undefined" ||
    typeof associatedEntity === "undefined" ||
    typeof associatedEntityPrimaryIdAttribute === "undefined" ||
    typeof associatedEntityPrimaryNameAttribute === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  const fetchXml = `<fetch>
    <entity name="${intersectEntity}">
      <attribute name="${intersectPrimaryIdAttribute}" />  
      <filter>
        <condition attribute="${currentIntersectAttribute}" operator="eq" value="${currentRecordId}" />
      </filter>
      <link-entity name="${associatedEntity}" from="${associatedEntityPrimaryIdAttribute}" to="${associatedIntersectAttribute}" alias="aLink">
        <attribute name="${associatedEntityPrimaryIdAttribute}" />  
        <attribute name="${associatedEntityPrimaryNameAttribute}" />
      </link-entity>
    </entity>
  </fetch>`;

  const results = await retrieveMultipleFetch(intersectEntitySet, fetchXml);
  return results.map((r) => {
    const id = r[`aLink.${associatedEntityPrimaryIdAttribute}`];
    const name = r[`aLink.${associatedEntityPrimaryNameAttribute}`];
    return {
      [intersectPrimaryIdAttribute]: r[intersectPrimaryIdAttribute],
      [associatedEntityPrimaryIdAttribute]: id,
      [associatedEntityPrimaryNameAttribute]: name,
    } as ComponentFramework.WebApi.Entity;
  });
}

export function associateRecord(
  entitySetName: string | undefined,
  currentRecordId: string | undefined,
  associatedEntitySet: string | undefined,
  associateRecordId: string | undefined,
  relationshipName: string | undefined,
  clientUrl: string | undefined
) {
  if (
    typeof entitySetName === "undefined" ||
    typeof currentRecordId === "undefined" ||
    typeof associatedEntitySet === "undefined" ||
    typeof associateRecordId === "undefined" ||
    typeof relationshipName === "undefined" ||
    typeof clientUrl === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  return axios.post(
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId})/${relationshipName}/$ref`,
    {
      "@odata.id": `${clientUrl}/api/data/v${apiVersion}/${associatedEntitySet}(${associateRecordId
        .replace("{", "")
        .replace("}", "")})`,
    },
    {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "If-None-Match": null,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    }
  );
}

export function disassociateRecord(
  entitySetName: string | undefined,
  currentRecordId: string | undefined,
  relationshipName: string | undefined,
  associatedRecordId: string | undefined
) {
  if (
    typeof entitySetName === "undefined" ||
    typeof currentRecordId === "undefined" ||
    typeof relationshipName === "undefined" ||
    typeof associatedRecordId === "undefined"
  ) {
    return Promise.reject(new Error("Invalid arguments"));
  }

  return axios.delete(
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId})/${relationshipName}(${associatedRecordId})/$ref`,
    {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "If-Match": null,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    }
  );
}

export function createRecord(entitySetName: string | undefined, record: ComponentFramework.WebApi.Entity) {
  if (typeof entitySetName === "undefined") return Promise.reject(new Error("Invalid entity set name"));

  return axios.post(`/api/data/v${apiVersion}/${entitySetName}`, record, {
    headers: {
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
      "Content-Type": "application/json",
    },
  });
}

export function deleteRecord(entitySetName: string | undefined, recordId: string | undefined) {
  if (typeof entitySetName === "undefined" || typeof recordId === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  return axios.delete(`/api/data/v${apiVersion}/${entitySetName}(${recordId})`, {
    headers: {
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "If-Match": null,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
  });
}

export function getCurrentRecord(attributes?: string[]): ComponentFramework.WebApi.Entity {
  let entityAttributes = Xrm.Page.data.entity.attributes.get();

  if (attributes?.length) {
    entityAttributes = entityAttributes.filter((a) => attributes.includes(a.getName()));
  }

  return Object.fromEntries(
    entityAttributes.map((attribute) => {
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

async function getEnvironmentVariableValue(environmentVariableName: string) {
  try {
    const { data: evDef } = await axios.get<{
      value: { environmentvariabledefinitionid: string; defaultvalue?: string }[];
    }>(
      `/api/data/v${apiVersion}/environmentvariabledefinitions?$filter=schemaname eq '${environmentVariableName}' or displayname eq '${environmentVariableName}' &$select=environmentvariabledefinitionid,defaultvalue&$top=1`
    );

    // environment variable not found
    if (evDef.value.length === 0) return null;

    const defaultValue = evDef.value[0].defaultvalue ?? "";

    // get environment variable value
    const { data: evValue } = await axios.get<{ value: { value: string }[] }>(
      `/api/data/v${apiVersion}/environmentvariablevalues?$filter=_environmentvariabledefinitionid_value eq '${evDef.value[0].environmentvariabledefinitionid}'&$select=value&$top=1`
    );

    return evValue?.value.at(0)?.value ?? defaultValue;
  } catch (error) {
    return null;
  }
}
