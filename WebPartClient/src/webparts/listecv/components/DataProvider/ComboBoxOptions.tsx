import { IComboBoxOption, SelectableOptionMenuItemType } from "office-ui-fabric-react/lib/index";
export const ComboBoxOptions: IComboBoxOption[] = [
  {
    key: "Header1",
    text: "A long terme",
    itemType: SelectableOptionMenuItemType.Header
  },
  { key: "A", text: "Recrutement <2ans" },
  { key: "divider", text: "-", itemType: SelectableOptionMenuItemType.Divider },
  {
    key: "Header2",
    text: "A court terme",
    itemType: SelectableOptionMenuItemType.Header
  },
  { key: "C", text: "Maintenance" },
  { key: "D", text: "Testeur" },
  { key: "E", text: "Projet" }
];
