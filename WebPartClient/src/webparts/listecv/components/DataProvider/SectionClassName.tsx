import { mergeStyles } from "office-ui-fabric-react/lib/index";
export const SectionClassName = mergeStyles({
  display: "flex",
  selectors: {
    "& > *": { marginRight: "20px" },
    "& .ms-ComboBox": { maxWidth: "500px" }
  }
});
