import axios from "axios";
import { IEntityDefinition, IManyToManyRelationship, IMetadata, IViewDefinition, IViewLayout } from "../types/metadata";
import { useQuery } from "@tanstack/react-query";

const nToNColumns = [
  "SchemaName",
  "Entity1LogicalName",
  "Entity2LogicalName",
  "Entity1IntersectAttribute",
  "Entity2IntersectAttribute",
  "RelationshipType",
  "IntersectEntityName",
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

export function useMetadata(currentTable: string, relationshipName: string, lookupView: string | undefined) {
  return useQuery({
    queryKey: ["metadata", { currentTable, relationshipName, lookupView }],
    queryFn: () => getMetadata(currentTable, relationshipName, lookupView),
    enabled: !!currentTable && !!relationshipName,
  });
}

export function useSuggestions(
  associatedTableSetName: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  fetchXmlTemplate: HandlebarsTemplateDelegate<any>,
  pageSize: number | undefined
) {
  return useQuery({
    queryKey: ["suggestionItems", { associatedTableSetName, pageSize }],
    queryFn: () => {
      const currentRecord = getCurrentRecord();
      const fetchXml = fetchXmlTemplate(currentRecord);
      return retrieveMultipleFetch(associatedTableSetName, fetchXml, 1, pageSize);
    },
    enabled: !!associatedTableSetName,
  });
}

export function useSelectedItems(
  currentTable: string,
  currentRecordId: string,
  relationshipName: string,
  metadata: IMetadata | undefined
) {
  const associatedTable = metadata?.associatedEntity.LogicalName ?? "";
  const associatedIntesectAttribute =
    metadata?.nnRelationship.Entity1LogicalName === currentTable
      ? metadata?.nnRelationship.Entity2IntersectAttribute
      : metadata?.nnRelationship.Entity1IntersectAttribute;

  const currentIntesectAttribute =
    metadata?.nnRelationship.Entity1LogicalName === currentTable
      ? metadata?.nnRelationship?.Entity1IntersectAttribute
      : metadata?.nnRelationship?.Entity2IntersectAttribute;

  return useQuery({
    queryKey: ["selectedItems", { currentTable, currentRecordId, relationshipName }],
    queryFn: () =>
      retrieveAssociatedRecords(
        currentRecordId,
        metadata?.intersectEntity.LogicalName,
        metadata?.intersectEntity.EntitySetName,
        metadata?.intersectEntity.PrimaryIdAttribute,
        currentIntesectAttribute,
        associatedIntesectAttribute,
        associatedTable,
        metadata?.associatedEntity.PrimaryIdAttribute,
        metadata?.associatedEntity.PrimaryNameAttribute
      ),
    enabled: !!metadata?.intersectEntity.EntitySetName && !!metadata?.associatedEntity.EntitySetName,
  });
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

export function getEntityDefinition(entityName: string | undefined) {
  return typeof entityName === "undefined"
    ? Promise.reject(new Error("Invalid entity name"))
    : axios
        .get<IEntityDefinition>(`/api/data/v${apiVersion}/EntityDefinitions(LogicalName='${entityName}')`, {
          params: {
            $select: tableDefinitionColumns.join(","),
          },
        })
        .then((res) => res.data);
}

export async function getViewDefinition(
  entityName: string | undefined,
  viewName: string | undefined,
  queryType?: number | undefined
) {
  if (typeof entityName === "undefined") return Promise.reject(new Error("Invalid arguments"));

  if (typeof viewName === "undefined" && typeof queryType === "undefined")
    return Promise.reject(new Error("Invalid arguments"));

  const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `returnedtypecode eq '${entityName}' ${viewName ? `and name eq '${viewName}'` : ""} ${
        queryType ? `and querytype eq ${queryType}` : ""
      } `,
      $select: viewDefinitionColumns.join(","),
    },
  });

  const layoutjson = result.data.value[0].layoutjson as unknown as string;
  const layout = JSON.parse(layoutjson) as IViewLayout;
  result.data.value[0].layoutjson = layout;
  return result.data.value[0];
}

export async function getDefaultView(entityName: string | undefined, viewName: string | undefined) {
  const viewByName = await getViewDefinition(entityName, viewName);
  if (viewByName) return viewByName;

  const defaultView = await getViewDefinition(entityName, undefined, 64);
  return defaultView;
}

export function getMetadata(
  currentTable: string | undefined,
  relationshipName: string | undefined,
  associatedViewName: string | undefined
): Promise<IMetadata> {
  return typeof currentTable === "undefined" || typeof relationshipName === "undefined"
    ? Promise.reject(new Error("Invalid arguments"))
    : getManytoManyRelationShipDefinition(currentTable, relationshipName).then((nnRelationship) => {
        const associatedEntity =
          nnRelationship.Entity1LogicalName === currentTable
            ? nnRelationship.Entity2LogicalName
            : nnRelationship.Entity1LogicalName;
        return Promise.all([
          getEntityDefinition(currentTable),
          getEntityDefinition(nnRelationship.IntersectEntityName),
          getEntityDefinition(associatedEntity),
          getDefaultView(associatedEntity, associatedViewName),
        ]).then(([currentEntity, intersectEntity, associatedEntity, associatedView]) => {
          return {
            nnRelationship,
            currentEntity,
            intersectEntity,
            associatedEntity,
            associatedView,
          };
        });
      });
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

export function retrieveMultipleFetch(
  entitySetName: string | undefined,
  fetchXml: string | undefined,
  page?: number,
  count?: number,
  pagingCookies?: string
) {
  if (typeof entitySetName === "undefined" || typeof fetchXml === "undefined")
    return Promise.reject(new Error("Invalid entity set name or fetchXml"));

  const doc = new DOMParser().parseFromString(fetchXml, "application/xml");
  doc.documentElement.setAttribute("page", page?.toString() ?? "1");
  doc.documentElement.setAttribute("count", count?.toString() ?? "50");
  if (pagingCookies) {
    doc.documentElement.setAttribute("paging-cookie", pagingCookies);
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
    .then((res) => res.data.value);
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
      <link-entity name="${associatedEntity}" from="${associatedIntersectAttribute}" to="${associatedEntityPrimaryIdAttribute}" alias="aLink">
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
