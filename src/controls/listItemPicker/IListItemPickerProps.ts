import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IListItemPickerProps {
  columnInternalName: string;
  keyColumnInternalName?: string;
  context: WebPartContext |  ApplicationCustomizerContext;
  listId: string;
  itemLimit: number;
  filterList?: string;
  className?: string;
  webUrl?: string;
  defaultSelectedItems?: any[];
  disabled?: boolean;
  suggestionsHeaderText?:string;
  noResultsFoundText?:string;
  removeDuplicates?:boolean;
  required?: boolean;
  onSelectedItem: (item:any) => void;
}
