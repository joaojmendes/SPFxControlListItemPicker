export interface IListItemPickerState {
  noresultsFoundText: string;
  showError: boolean;
  errorMessage: string;
  suggestionsHeaderText:string;
  selectedItems:{ key: string; name: string }[];
  isRequired?:boolean;
  isLoading:boolean;
}
