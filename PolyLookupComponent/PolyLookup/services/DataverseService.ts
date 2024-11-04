import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import axios from "axios";
import Handlebars from "handlebars";
import { LanguagePack } from "../types/languagePack";
import {
  IEntityDefinition,
  IManyToManyRelationship,
  IMetadata,
  IOneToManyRelationship,
  IViewDefinition,
  IViewLayout,
  LookupView,
  isManyToMany,
} from "../types/metadata";
import { EntityOption, RelationshipTypeEnum } from "../types/typings";
import { getHandlebarsVariables } from "./TemplateService";

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
  "HasEmailAddresses",
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
    'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.associatednavigationproperty,Microsoft.Dynamics.CRM.lookuplogicalname,Microsoft.Dynamics.CRM.morerecords"',
};

const parser = new DOMParser();
const serializer = new XMLSerializer();

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
  formType: XrmEnum.FormType | undefined,
  firstViewColumn?: string,
  selectedItemTemplate?: string | null
) {
  return useQuery({
    queryKey: ["selectedItems", metadata, currentRecordId, formType, firstViewColumn, selectedItemTemplate],
    queryFn: () => {
      if (!metadata || !currentRecordId || formType === XrmEnum.FormType.Create) {
        return [];
      }

      return retrieveAssociatedRecords(metadata, currentRecordId, firstViewColumn, selectedItemTemplate);
    },
    enabled: !!metadata?.intersectEntity.EntitySetName && !!metadata?.associatedEntity.EntitySetName,
  });
}

export function useEntityOptions(
  metadata: IMetadata | undefined,
  lookupViewConfig: LookupView | undefined,
  searchText?: string,
  pageSize?: number
) {
  return useQuery({
    queryKey: ["entityOptions", lookupViewConfig, searchText],
    queryFn: () => getEntityOptions(metadata, lookupViewConfig, searchText, pageSize),
    enabled: !!lookupViewConfig?.fetchXml,
  });
}

export function useLanguagePack(webResourcePath: string | undefined, defaultLanguagePack: LanguagePack) {
  return useQuery({
    queryKey: ["languagePack", webResourcePath],
    queryFn: () => getLanguagePack(webResourcePath, defaultLanguagePack),
    enabled: !!webResourcePath,
  });
}

export function useFilterQuery(metadata: IMetadata | undefined, lookupViewConfig: LookupView | undefined) {
  return useMutation({
    mutationFn: ({ searchText, pageSizeParam }: { searchText?: string; pageSizeParam?: number }) =>
      getEntityOptions(metadata, lookupViewConfig, searchText, pageSizeParam),
  });
}

export function useAssociateQuery(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  relationshipType: RelationshipTypeEnum,
  clientUrl: string,
  languagePack: LanguagePack
) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return associateRecord(
          metadata?.currentEntity.EntitySetName,
          currentRecordId,
          metadata?.associatedEntity?.EntitySetName,
          id,
          metadata?.relationship1.SchemaName,
          clientUrl
        );
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return createRecord(metadata?.intersectEntity.EntitySetName, {
          [`${metadata?.currentEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.currentEntity
            .EntitySetName}(${currentRecordId.replace("{", "").replace("}", "")})`,
          [`${metadata?.associatedEntityNavigationPropertyName}@odata.bind`]: `/${metadata?.associatedEntity
            .EntitySetName}(${id.replace("{", "").replace("}", "")})`,
        });
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({
        queryKey: ["selectedItems"],
      });
    },
  });
}

export function useDisassociateQuery(
  metadata: IMetadata | undefined,
  currentRecordId: string,
  relationshipType: RelationshipTypeEnum,
  languagePack: LanguagePack
) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (id: string) => {
      if (relationshipType === RelationshipTypeEnum.ManyToMany) {
        return disassociateRecord(
          metadata?.currentEntity?.EntitySetName,
          currentRecordId,
          metadata?.relationship1.SchemaName,
          id
        );
      } else if (
        relationshipType === RelationshipTypeEnum.Custom ||
        relationshipType === RelationshipTypeEnum.Connection
      ) {
        return deleteRecord(metadata?.intersectEntity.EntitySetName, id);
      }
      return Promise.reject(languagePack.RelationshipNotSupportedMessage);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({
        queryKey: ["selectedItems"],
      });
    },
  });
}

export async function getEntityOptions(
  metadata: IMetadata | undefined,
  lookupViewConfig: LookupView | undefined,
  searchText?: string,
  pageSize?: number
) {
  if (!metadata || !lookupViewConfig) return Promise.resolve({ records: [], moreRecords: false });

  let fetchXml = lookupViewConfig.fetchXml;
  let shouldDefaultSearch = false;
  if (lookupViewConfig.isSystemLookupView) {
    shouldDefaultSearch = true;
  } else {
    if (
      !fetchXml.includes("{{PolyLookupSearch}}") &&
      !fetchXml.includes("{{ PolyLookupSearch}}") &&
      !fetchXml.includes("{{PolyLookupSearch }}") &&
      !fetchXml.includes("{{ PolyLookupSearch }}")
    ) {
      shouldDefaultSearch = true;
    }

    const fetchXmlTemplateFn = Handlebars.compile(fetchXml);

    const currentRecord = getCurrentRecord();
    fetchXml =
      fetchXmlTemplateFn({
        ...currentRecord,
        PolyLookupSearch: searchText,
      }) ?? fetchXml;
  }

  if (shouldDefaultSearch && searchText) {
    // if lookup view is not specified and using default lookup fiew,
    // add filter condition to fetchxml to support search
    const doc = parser.parseFromString(fetchXml, "application/xml");
    const entities = doc.documentElement.getElementsByTagName("entity");
    for (let i = 0; i < entities.length; i++) {
      const entity = entities[i];
      if (entity.getAttribute("name") === metadata.associatedEntity.LogicalName) {
        const filter = doc.createElement("filter");
        const condition = doc.createElement("condition");
        condition.setAttribute("attribute", metadata.associatedEntity.PrimaryNameAttribute);
        if (searchText.includes("*")) {
          const beginsWithWildCard = searchText.startsWith("*") ? true : false;
          const endsWithWildCard = searchText.endsWith("*") ? true : false;
          if (beginsWithWildCard || endsWithWildCard) {
            if (beginsWithWildCard) {
              searchText = searchText.replace("*", "");
            }
            if (endsWithWildCard) {
              searchText = searchText.substring(0, searchText.length - 1);
            }
            searchText = "%" + searchText + "%";
          } else {
            searchText = searchText + "%";
          }
          searchText = searchText.split("*").join("%");
          condition.setAttribute("operator", "like");
          condition.setAttribute("value", `${searchText}`);
        } else {
          condition.setAttribute("operator", "begins-with");
          condition.setAttribute("value", `${searchText}`);
        }
        filter.appendChild(condition);
        entity.appendChild(filter);
      }
    }
    fetchXml = serializer.serializeToString(doc);
  }

  const { records, moreRecords } = await retrieveMultipleFetch(
    metadata.associatedEntity.EntitySetName,
    fetchXml,
    1,
    pageSize
  );
  return {
    records: records.map((r) => {
      const iconSrc = metadata.associatedEntity.PrimaryImageAttribute
        ? `/api/data/v${apiVersion}/${metadata.associatedEntity.EntitySetName}(${
            r[metadata.associatedEntity.PrimaryIdAttribute]
          })/${metadata.associatedEntity.PrimaryImageAttribute}/$value`
        : "";

      return {
        id: r[metadata.associatedEntity.PrimaryIdAttribute],
        intersectId: "",
        associatedId: r[metadata.associatedEntity.PrimaryIdAttribute],
        associatedName: r[metadata.associatedEntity.PrimaryNameAttribute],
        optionText: r[metadata.associatedEntity.PrimaryNameAttribute],
        selectedOptionText: r[metadata.associatedEntity.PrimaryNameAttribute],
        iconSrc,
        entity: r,
      } as EntityOption;
    }),
    moreRecords,
  };
}

function getLanguagePack(webResourceUrl: string | undefined, defaultLanguagePack: LanguagePack): Promise<LanguagePack> {
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
            headers: DEFAULT_HEADERS,
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
            headers: DEFAULT_HEADERS,
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
            headers: DEFAULT_HEADERS,
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
          headers: DEFAULT_HEADERS,
          params: {
            $select: tableDefinitionColumns.join(","),
          },
        })
        .then(({ data }) => {
          const entityIconName = data.IconVectorName ?? data.IconMediumName ?? data.IconSmallName;
          const EntityIconUrl = entityIconName ? `/WebResources/${entityIconName}` : "";

          return {
            ...data,
            DisplayNameLocalized: data.DisplayName.UserLocalizedLabel?.Label ?? "",
            DisplayCollectionNameLocalized: data.DisplayCollectionName.UserLocalizedLabel?.Label ?? "",
            EntityIconUrl,
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
    headers: DEFAULT_HEADERS,
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
      currentIntersectAttribute: currentIntesectAttribute,
      associatedIntersectAttribute: associatedIntesectAttribute,
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
): Promise<LookupView> {
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
      firstAttribute: primaryNameAttribute,
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
      firstAttribute: primaryNameAttribute,
      isSystemLookupView: false,
    };

    const url =
      lookupViewVal.startsWith("https://") || lookupViewVal.startsWith("/api")
        ? lookupViewVal
        : `/api/data/v${apiVersion}/${lookupViewVal}`;

    try {
      const { data } = await axios.get<{ value?: unknown }>(url, {
        headers: DEFAULT_HEADERS,
      });

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
          firstAttribute: primaryNameAttribute,
          isSystemLookupView: false,
        };
      }
    }

    if (!lookupViewConfig) {
      // environment variable not found
      // check if lookupViewValue is a view name
      const viewDef = await getDefaultView(associatedEntityName, lookupViewVal);
      if (viewDef) {
        const columns = viewDef.layoutjson.Rows[0].Cells.map((c) => c.Name);
        lookupViewConfig = {
          sourceType: "ViewName",
          source: lookupViewVal,
          fetchXml: viewDef.fetchxml,
          columns,
          firstAttribute: columns[0],
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
    lookupViewConfig.columns = attributes
      .map((attr) => {
        const entityAlias = attr.parentElement?.getAttribute("alias");
        const attributeName = attr.getAttribute("name") ?? "";
        const attributeAlias = attr.getAttribute("alias");
        if (attributeAlias) {
          return attributeAlias;
        }
        if (entityAlias) {
          return `${entityAlias}.${attributeName}`;
        }
        return attributeName;
      })
      .filter((attr) => attr !== "");
    lookupViewConfig.firstAttribute = attributes[0].getAttribute("name") ?? primaryNameAttribute;
  }

  // check if doc has attribute with name equals primaryIdAttribute and primaryNameAttribute
  const hasPrimaryIdAttribute = attributes.some(
    (attr) => attr.getAttribute("name") === primaryIdAttribute && !attr.hasAttribute("alias")
  );
  const hasPrimaryNameAttribute = attributes.some(
    (attr) => attr.getAttribute("name") === primaryNameAttribute && !attr.hasAttribute("alias")
  );

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

    const { data } = await axios.get<{
      value: ComponentFramework.WebApi.Entity[];
      "@Microsoft.Dynamics.CRM.morerecords": boolean;
    }>(`/api/data/v${apiVersion}/${entitySetName}`, {
      headers: DEFAULT_HEADERS,
      params: {
        fetchXml: encodeURIComponent(newFetchXml),
      },
    });

    if (data.value.length === 0) {
      return {
        records: [],
        moreRecords: false,
      };
    }

    return {
      records: data.value,
      moreRecords: data["@Microsoft.Dynamics.CRM.morerecords"],
    };
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
  {
    relationship1,
    intersectEntity,
    associatedEntity,
    currentIntersectAttribute,
    associatedIntersectAttribute,
  }: IMetadata,
  currentRecordId: string,
  firstViewColumn?: string,
  selectedItemTemplate?: string | null
) {
  const columns = getHandlebarsVariables(selectedItemTemplate ?? "");
  let fetchColumns = [
    ...columns,
    firstViewColumn ?? associatedEntity.PrimaryNameAttribute,
    associatedEntity.PrimaryIdAttribute,
    associatedEntity.PrimaryNameAttribute,
  ];
  fetchColumns = Array.from(new Set(fetchColumns));

  const fetchXml = `<fetch>
    <entity name="${intersectEntity.LogicalName}">
      <attribute name="${intersectEntity.PrimaryIdAttribute}" />  
      <filter>
        <condition attribute="${currentIntersectAttribute}" operator="eq" value="${currentRecordId}" />
      </filter>
      <link-entity name="${associatedEntity.LogicalName}" from="${
        associatedEntity.PrimaryIdAttribute
      }" to="${associatedIntersectAttribute}" alias="aLink">
        ${fetchColumns.map((c) => `<attribute name="${c}" />`).join("")}
      </link-entity>
    </entity>
  </fetch>`;

  const { records: results } = await retrieveMultipleFetch(intersectEntity.EntitySetName, fetchXml);
  return results.map((r) => {
    const intersectId = r[intersectEntity.PrimaryIdAttribute];

    const entity = {} as ComponentFramework.WebApi.Entity;

    fetchColumns.forEach((c) => {
      entity[c] = r[`aLink.${c}`];
    });

    const associatedId = entity[associatedEntity.PrimaryIdAttribute];
    const associatedName = entity[associatedEntity.PrimaryNameAttribute];

    const selectedOpionTemplateFn = selectedItemTemplate ? Handlebars.compile(selectedItemTemplate) : null;
    const selectedOptionText =
      selectedOpionTemplateFn?.(entity) ?? entity[firstViewColumn ?? associatedEntity.PrimaryNameAttribute];

    const iconSrc = associatedEntity.PrimaryImageAttribute
      ? `/api/data/v${apiVersion}/${associatedEntity.EntitySetName}(${associatedId})/${associatedEntity.PrimaryImageAttribute}/$value`
      : "";

    return {
      id: relationship1.RelationshipType === "ManyToManyRelationship" ? associatedId : intersectId,
      intersectId,
      associatedId,
      associatedName,
      optionText: associatedName,
      selectedOptionText,
      iconSrc,
      entity,
    } as EntityOption;
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
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId
      .replace("{", "")
      .replace("}", "")})/${relationshipName}/$ref`,
    {
      "@odata.id": `${clientUrl}/api/data/v${apiVersion}/${associatedEntitySet}(${associateRecordId
        .replace("{", "")
        .replace("}", "")})`,
    },
    {
      headers: DEFAULT_HEADERS,
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
    `/api/data/v${apiVersion}/${entitySetName}(${currentRecordId
      .replace("{", "")
      .replace("}", "")})/${relationshipName}(${associatedRecordId})/$ref`,
    {
      headers: DEFAULT_HEADERS,
    }
  );
}

export function createRecord(entitySetName: string | undefined, record: ComponentFramework.WebApi.Entity) {
  if (typeof entitySetName === "undefined") return Promise.reject(new Error("Invalid entity set name"));

  return axios.post(`/api/data/v${apiVersion}/${entitySetName}`, record, {
    headers: DEFAULT_HEADERS,
  });
}

export function deleteRecord(entitySetName: string | undefined, recordId: string | undefined) {
  if (typeof entitySetName === "undefined" || typeof recordId === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  return axios.delete(`/api/data/v${apiVersion}/${entitySetName}(${recordId.replace("{", "").replace("}", "")})`, {
    headers: DEFAULT_HEADERS,
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
      value: {
        defaultvalue?: string | null;
        environmentvariabledefinition_environmentvariablevalue: { value?: string | null }[];
      }[];
    }>(
      `/api/data/v${apiVersion}/environmentvariabledefinitions?$select=environmentvariabledefinitionid,defaultvalue&$filter=schemaname eq '${environmentVariableName}'&$top=1&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`,
      {
        headers: DEFAULT_HEADERS,
      }
    );

    const defaultValue = evDef?.value.at(0)?.defaultvalue ?? null;
    const currentValue =
      evDef?.value.at(0)?.environmentvariabledefinition_environmentvariablevalue?.at(0)?.value ?? null;
    return currentValue ?? defaultValue;
  } catch (error) {
    return null;
  }
}
