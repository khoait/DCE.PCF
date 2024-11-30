import { useQueryClient, useMutation } from "@tanstack/react-query";
import { associateRecord, createRecord } from "../../services/DataverseService";
import { LanguagePack } from "../../types/languagePack";
import { IMetadata } from "../../types/metadata";
import { RelationshipTypeEnum } from "../../types/typings";

export function useAssociate(
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
