import {
  Avatar,
  FluentProvider,
  InteractionTag,
  InteractionTagPrimary,
  InteractionTagSecondary,
  TagGroup,
  TagGroupProps,
  TagPicker,
  TagPickerControl,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerProps,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { useQueryClient } from "@tanstack/react-query";
import React from "react";
import {
  useAssociateQuery,
  useDisassociateQuery,
  useEntityOptions,
  useLanguagePack,
  useLookupViewConfig,
  useMetadata,
  useSelectedItems,
} from "../services/DataverseService";
import { PolyLookupProps, RelationshipTypeEnum } from "../types/typings";

const useStyle = makeStyles({
  tagGroup: {
    flexWrap: "wrap",
    gap: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
  },
});

export default function PolyLookupControlNewLook({
  currentTable,
  currentRecordId,
  relationshipName,
  relationship2Name,
  relationshipType,
  clientUrl,
  lookupView,
  itemLimit,
  pageSize,
  disabled,
  formType,
  outputSelectedItems,
  tagAction,
  defaultLanguagePack,
  languagePackPath,
  fluentDesign,
  onChange,
  onQuickCreate,
}: PolyLookupProps) {
  const queryClient = useQueryClient();
  const { tagGroup } = useStyle();

  const [searchText, setSearchText] = React.useState<string>("");

  const { data: loadedLanguagePack } = useLanguagePack(languagePackPath, defaultLanguagePack);

  const languagePack = loadedLanguagePack ?? defaultLanguagePack;

  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
    isError: isErrorMetadata,
    error: errorMetadata,
  } = useMetadata(
    currentTable,
    relationshipName,
    relationshipType === RelationshipTypeEnum.Custom || relationshipType === RelationshipTypeEnum.Connection
      ? relationship2Name
      : undefined
  );

  // get LookupView configuration
  const {
    data: lookupViewConfig,
    isSuccess: isSuccessLookupViewConfig,
    isError: isErrorLookupView,
    error: errorLookupView,
  } = useLookupViewConfig(
    metadata?.associatedEntity.LogicalName,
    metadata?.associatedEntity.PrimaryIdAttribute,
    metadata?.associatedEntity.PrimaryNameAttribute,
    lookupView
  );

  // get selected items
  const {
    data: selectedItems,
    isPending: isLoadingSelectedItems,
    isSuccess: isLoadingSelectedItemsSuccess,
    isError: isErrorSelectedItems,
    error: errorSelectedItems,
  } = useSelectedItems(metadata, currentRecordId, formType);

  const {
    data: entityOptions,
    isPending: isLoadingEntityOptions,
    isSuccess: isSuccessEntityOptions,
    isError: isErrorEntityOptions,
    error: errorEntityOptions,
  } = useEntityOptions(metadata, lookupViewConfig, searchText, pageSize);

  const { mutate: associateQuery } = useAssociateQuery(
    metadata,
    currentRecordId,
    relationshipType,
    clientUrl,
    languagePack
  );

  const { mutate: disassociateQuery } = useDisassociateQuery(metadata, currentRecordId, relationshipType, languagePack);

  const handleOnOptionSelect: TagPickerProps["onOptionSelect"] = (_e, { value }) => {
    if (itemLimit && (selectedItems?.length ?? 0) >= itemLimit) {
      return;
    }
    associateQuery(value, {
      onSuccess: () => {
        queryClient.invalidateQueries({
          queryKey: ["selectedItems"],
        });
      },
    });
  };

  const handleOnOptionDismiss: TagGroupProps["onDismiss"] = (_e, { value }) => {
    disassociateQuery(value, {
      onSuccess: () => {
        queryClient.invalidateQueries({
          queryKey: ["selectedItems"],
        });
      },
    });
  };

  return (
    <FluentProvider style={{ width: "100%" }} theme={fluentDesign?.tokenTheme}>
      <TagPicker
        appearance="filled-darker"
        size="large"
        selectedOptions={selectedItems?.map((i) => i.id) ?? []}
        onOptionSelect={handleOnOptionSelect}
      >
        <TagPickerControl>
          <TagGroup className={tagGroup} onDismiss={handleOnOptionDismiss}>
            {selectedItems?.map((i) => (
              <InteractionTag key={i.id} shape="rounded" appearance={tagAction ? "brand" : "outline"} value={i.id}>
                <InteractionTagPrimary
                  hasSecondaryAction
                  media={
                    <Avatar
                      name={i.associatedName}
                      image={{ src: i.iconSrc }}
                      color={i.iconSrc?.startsWith("/WebResource") ? "neutral" : "colorful"}
                      aria-hidden
                    />
                  }
                >
                  {i.selectedOptionText}
                </InteractionTagPrimary>
                <InteractionTagSecondary aria-label="remove" />
              </InteractionTag>
            ))}
          </TagGroup>
          <TagPickerInput value={searchText} onChange={(e) => setSearchText(e.target.value)} />
        </TagPickerControl>
        <TagPickerList>
          {entityOptions?.length
            ? entityOptions.map((option) => (
                <TagPickerOption
                  media={
                    <Avatar
                      shape="square"
                      name={option.associatedName}
                      image={{ src: option.iconSrc }}
                      color={option.iconSrc?.startsWith("/WebResource") ? "neutral" : "colorful"}
                      aria-hidden
                    />
                  }
                  value={option.id}
                  key={option.id}
                >
                  {option.optionText}
                </TagPickerOption>
              ))
            : "No options available"}
        </TagPickerList>
      </TagPicker>
    </FluentProvider>
  );
}
