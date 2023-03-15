import axios from "axios";

type RelationshipType = "OneToManyRelationship" | "ManyToManyRelationship";

export interface IRelationshipDefinition {
  SchemaName?: string;
  RelationshipType?: RelationshipType;
}

export interface IManyToManyRelationship {
  SchemaName: string;
  Entity1LogicalName: string;
  Entity2LogicalName: string;
  Entity1IntersectAttribute: string;
  Entity2IntersectAttribute: string;
  RelationshipType: string;
  IntersectEntityName: string;
}

export interface IOneToManyRelationship extends IRelationshipDefinition {
  ReferencedAttribute?: string;
  ReferencedAttributeName?: string;
  ReferencedEntity?: string;
  ReferencedEntityName?: string;
  ReferencingAttribute?: string;
  ReferencingAttributeName?: string;
  ReferencingEntity?: string;
  ReferencingEntityName?: string;
  ReferencedEntityNavigationPropertyName?: string;
  ReferencingEntityNavigationPropertyName?: string;
}

export interface IEntityDefinition {
  LogicalName: string;
  DisplayName: {
    UserLocalizedLabel: {
      Label: string;
    };
  };
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  EntitySetName: string;
  DisplayCollectionName: {
    UserLocalizedLabel: {
      Label: string;
    };
  };
  IsQuickCreateEnabled: boolean;
}

interface IViewRow {
  Name: string;
  Id: string;
  Cells: IViewCell[];
  MultiObjectIdField: string;
  LayoutStyle: string;
}

interface IViewCell {
  Name: string;
  Width: number;
  RelatedEntityName: string;
  DisableMetaDataBinding: boolean;
  LabelId: string;
  IsHidden: boolean;
  DisableSorting: boolean;
  AddedBy: string;
  Desc: string;
  CellType: string;
  ImageProviderWebresource: string;
  ImageProviderFunctionName: string;
}

interface IViewLayout {
  Name: string;
  Object: number;
  Rows: IViewRow[];
  CustomControlDescriptions: [];
  Jump: string;
  Select: true;
  Icon: true;
  Preview: true;
  IconRenderer: "";
}

export interface IViewDefinition {
  savedqueryid: string;
  name: string;
  fetchxml: string;
  layoutjson: IViewLayout;
  querytype: number;
}

export interface IMetadata {
  nnRelationship: IManyToManyRelationship;
  currentEntity: IEntityDefinition;
  intersectEntity: IEntityDefinition;
  associatedEntity: IEntityDefinition;
  associatedView: IViewDefinition;
}

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

export async function getViewDefinition(entityName: string | undefined, viewName: string | undefined) {
  if (typeof entityName === "undefined") return Promise.reject(new Error("Invalid arguments"));

  if (viewName) {
    const result = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
      params: {
        $filter: `returnedtypecode eq '${entityName}' and name eq '${viewName}'`,
        $select: viewDefinitionColumns.join(","),
      },
    });

    if (result?.data.value?.length) {
      return result.data.value[0];
    }
  }

  const lookupResult = await axios.get<{ value: IViewDefinition[] }>(`/api/data/v${apiVersion}/savedqueries`, {
    params: {
      $filter: `returnedtypecode eq '${entityName}' and querytype eq 64`,
      $select: viewDefinitionColumns.join(","),
    },
  });

  return lookupResult.data.value[0];
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
          getViewDefinition(associatedEntity, associatedViewName),
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
