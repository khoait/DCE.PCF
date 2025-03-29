import { DropdownMenuItemType, IDropdownOption } from "@fluentui/react";
import { LanguagePack } from "../types/languagePack";
import { EntityOption } from "../types/typings";

export function getClassicDropdownOptions(
  entityOptions: EntityOption[],
  selectedOption: EntityOption | null | undefined,
  groupBy: string | null,
  languagePack: LanguagePack
) {
  let options: IDropdownOption<EntityOption>[] = [];

  // group results by groupBy field
  if (groupBy) {
    const grouped: Record<string, EntityOption[]> = {};
    entityOptions.forEach((item) => {
      const key = item.group ?? languagePack.BlankValueLabel ?? "";

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
          key: item.id,
          text: item.optionText,
          data: item,
        });
      });
    });
  } else {
    options = entityOptions.map((item) => {
      return {
        key: item.id,
        text: item.optionText,
        data: item,
      };
    });
  }

  if (options.length === 0) {
    options.push({
      key: "no-records",
      text: languagePack?.EmptyListMessage ?? "",
      disabled: true,
    });
  }

  if (selectedOption && !options.some((o) => o.key === selectedOption.id)) {
    options.push({
      key: selectedOption.id,
      text: selectedOption.optionText,
      data: selectedOption,
      hidden: true,
    });
  }

  return options;
}
