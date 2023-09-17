import React from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Dropdown, DropdownMenuItemType, IStyle, IDropdownOption } from "@fluentui/react";
import { useFetchXmlData, useMetadata } from "../services/DataverseService";

const queryClient = new QueryClient();

export interface LookdownProps {
  lookupViewId: string;
  lookupEntity: string;
  selectedId: string | undefined;
  onChange?: (selectedItem: ComponentFramework.LookupValue | null) => void;
}

const DEFAULT_BORDER_STYLES: IStyle = {
  borderColor: "#666",
  borderWidth: 1,
  borderRadius: 0,
};

const Body = ({ lookupViewId, lookupEntity, selectedId, onChange }: LookdownProps) => {
  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(lookupEntity, lookupViewId);

  const {
    data: fetchData,
    isLoading: isLoadingFetchData,
    isSuccess: isLoadingFetchDataSuccess,
  } = useFetchXmlData(metadata?.lookupEntity.EntitySetName ?? "", lookupViewId, metadata?.lookupView.fetchxml ?? "");

  const dropdownOptions: IDropdownOption<ComponentFramework.WebApi.Entity>[] =
    metadata && fetchData
      ? fetchData.map((item) => {
          return {
            key: item[metadata.lookupEntity.PrimaryIdAttribute],
            text: item[metadata.lookupEntity.PrimaryNameAttribute],
            data: item,
          };
        })
      : [];

  return (
    <Dropdown
      selectedKey={selectedId ?? undefined}
      placeholder="---"
      options={dropdownOptions}
      onChange={(event, option) => {
        if (!onChange) return;
        if (!metadata) return;

        if (!option) {
          onChange(null);
          return;
        }

        const data = option.data as ComponentFramework.WebApi.Entity;
        const selectedItem = {
          entityType: metadata.lookupEntity.LogicalName,
          id: data[metadata.lookupEntity.PrimaryIdAttribute],
          name: data[metadata.lookupEntity.PrimaryNameAttribute],
        } as ComponentFramework.LookupValue;
        onChange(selectedItem);
      }}
      styles={(styleProps) => {
        return {
          root: {
            width: "100%",
          },
          title: {
            borderColor: styleProps.isOpen ? "#666" : "transparent",
            borderRadius: styleProps.isOpen ? 1 : 0,
            fontWeight: styleProps.isOpen ? 400 : 600,
            color: "#000",
            "&:active": {
              fontWeight: 400,
            },
            "&:hover": {
              fontWeight: 400,
            },
          },
          dropdown: {
            "&:active": DEFAULT_BORDER_STYLES,
            "&:focus": DEFAULT_BORDER_STYLES,
            "&:focus:after": DEFAULT_BORDER_STYLES,
            "&:hover .ms-Dropdown-caretDownWrapper": {
              display: "block",
            },
            "&:focus .ms-Dropdown-caretDownWrapper": {
              display: "block",
            },
            "&:active .ms-Dropdown-title": {
              fontWeight: 400,
            },
            "&:focus .ms-Dropdown-title": {
              fontWeight: 400,
            },
          },
          caretDownWrapper: {
            display: styleProps.isOpen ? "block" : "none",
          },
          callout: {
            maxHeight: 500,
          },
        };
      }}
    />
  );
};

export default function LookdownControl(props: LookdownProps) {
  return (
    <QueryClientProvider client={queryClient}>
      <Body {...props} />
    </QueryClientProvider>
  );
}
