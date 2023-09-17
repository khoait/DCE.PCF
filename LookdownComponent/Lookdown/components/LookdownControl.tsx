import React from "react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Dropdown, DropdownMenuItemType, IStyle } from "@fluentui/react";

const queryClient = new QueryClient();

export interface LookdownProps {
  lookupViewId: string;
  lookupEntity: string;
  selectedId: string | undefined;
}

const DEFAULT_BORDER_STYLES: IStyle = {
  borderColor: "#666",
  borderWidth: 1,
  borderRadius: 0,
};

const dropdownControlledExampleOptions = [
  { key: "fruitsHeader", text: "Fruits", itemType: DropdownMenuItemType.Header },
  { key: "apple", text: "Apple" },
  { key: "banana", text: "Banana" },
  { key: "orange", text: "Orange", disabled: true },
  { key: "grape", text: "Grape" },
  { key: "vegetablesHeader", text: "Vegetables", itemType: DropdownMenuItemType.Header },
  { key: "broccoli", text: "Broccoli" },
  { key: "carrot", text: "Carrot" },
  { key: "lettuce", text: "Lettuce" },
];

const Body = ({ lookupViewId, lookupEntity, selectedId }: LookdownProps) => {
  return (
    <Dropdown
      selectedKey={undefined}
      placeholder="---"
      options={dropdownControlledExampleOptions}
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
