import { useQueryClient, useMutation } from "@tanstack/react-query";
import { disassociateRecord, deleteRecord } from "../../services/DataverseService";
import { LanguagePack } from "../../types/languagePack";
import { IMetadata } from "../../types/metadata";
import { RelationshipTypeEnum } from "../../types/typings";

export function useDisassociate(
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
