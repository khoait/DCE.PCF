import React from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Dropdown, DropdownMenuItemType, IStyle, IDropdownOption } from "@fluentui/react";
import { useFetchXmlData, useMetadata } from "../services/DataverseService";

const queryClient = new QueryClient();

export interface LookdownProps {
  lookupViewId: string;
  lookupEntity: string;
  selectedId?: string | null;
  groupBy?: string | null;
  onChange?: (selectedItem: ComponentFramework.LookupValue | null) => void;
}

const DEFAULT_BORDER_STYLES: IStyle = {
  borderColor: "#666",
  borderWidth: 1,
  borderRadius: 0,
};

const Body = ({ lookupViewId, lookupEntity, selectedId, groupBy, onChange }: LookdownProps) => {
  const {
    data: metadata,
    isLoading: isLoadingMetadata,
    isSuccess: isLoadingMetadataSuccess,
  } = useMetadata(lookupEntity, lookupViewId);

  const {
    data: fetchData,
    isLoading: isLoadingFetchData,
    isSuccess: isLoadingFetchDataSuccess,
  } = useFetchXmlData(metadata?.lookupEntity.EntitySetName, lookupViewId, metadata?.lookupView.fetchxml, groupBy);

  const getDropdownOptions = (): IDropdownOption<ComponentFramework.WebApi.Entity>[] => {
    if (!metadata) return [];
    if (!fetchData) return [];

    let options: IDropdownOption<ComponentFramework.WebApi.Entity>[] = [];

    // group results by groupBy field
    if (groupBy) {
      const grouped: Record<string, ComponentFramework.WebApi.Entity[]> = {};
      fetchData.forEach((item) => {
        let key = item[groupBy]?.toString() ?? "(Blank)";

        // use formatted value if available
        const formattedValue = item[groupBy + "@OData.Community.Display.V1.FormattedValue"];
        if (formattedValue && formattedValue !== "") {
          key = formattedValue;
        }

        if (!grouped[key]) grouped[key] = [];
        grouped[key].push(item);
      });

      Object.keys(grouped).forEach((key) => {
        options.push({
          key: key,
          text: key,
          itemType: DropdownMenuItemType.Header,
        });

        grouped[key].forEach((item) => {
          options.push({
            key: item[metadata.lookupEntity.PrimaryIdAttribute],
            text: item[metadata.lookupEntity.PrimaryNameAttribute],
            data: item,
          });
        });
      });
    } else {
      options = fetchData.map((item) => {
        return {
          key: item[metadata.lookupEntity.PrimaryIdAttribute],
          text: item[metadata.lookupEntity.PrimaryNameAttribute],
          data: item,
        };
      });
    }

    return options;
  };

  return (
    <Dropdown
      selectedKey={selectedId ?? undefined}
      placeholder="---"
      options={getDropdownOptions()}
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
            "& .ms-Callout-main": {
              maxHeight: "500px !important",
            },
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
