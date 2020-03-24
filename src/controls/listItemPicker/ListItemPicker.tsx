import * as React from "react";
import * as strings from "ListItemPickerStrings";

import { BaseAutoFill, IBasePicker, ITag, Spinner, SpinnerSize } from "office-ui-fabric-react";
import { IListItemPickerProps, IListItemPickerState } from ".";
import { uniq, uniqBy } from '@microsoft/sp-lodash-subset';

import { Label } from "office-ui-fabric-react/lib/Label";
import SPservice from "../../spservices/spservices";
import { TagPicker } from "office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker";
import { escape } from "@microsoft/sp-lodash-subset";

export class ListItemPicker extends React.Component<IListItemPickerProps, IListItemPickerState> {
  private _spservice: SPservice;
  private selectedItems: any[];
  private taglist: ITag[] = [];
  private pickerRef = React.createRef<IBasePicker<ITag>>();



  constructor(props: IListItemPickerProps) {
    super(props);

    // States
    this.state = {
      noresultsFoundText: !this.props.noResultsFoundText ? strings.genericNoResultsFoundText : this.props.noResultsFoundText,
      showError: false,
      errorMessage: "",
      selectedItems: [],
      isRequired: this.props.required ? this.props.required : false,
      isLoading: false,
      suggestionsHeaderText: !this.props.suggestionsHeaderText ? strings.ListItemPickerSelectValue : this.props.suggestionsHeaderText
    };

    // Get SPService Factory
    this._spservice = new SPservice(this.props.context);

    this.selectedItems = [];
    this._onBlur = this._onBlur.bind(this);
  }

  public async componentDidUpdate(prevProps: IListItemPickerProps, prevState: IListItemPickerState): Promise<void> {

    if (this.props.listId !== prevProps.listId || this.props.filterList !== prevProps.filterList) {
      this.setState({
        selectedItems: []
      });
    }
    if ( this.props.defaultSelectedItems !== prevProps.defaultSelectedItems){
      this.setState({
        selectedItems: this.props.defaultSelectedItems
      });
    }

  }

  public componentWillMount(): void {

    if (this.props.defaultSelectedItems && this.props.defaultSelectedItems.length > 0) {
      this.setState({ selectedItems: this.props.defaultSelectedItems });
    }

  }
  /**
   * Render the field
   */
  public render(): React.ReactElement<IListItemPickerProps> {
    const { className, disabled, itemLimit } = this.props;

    return (
      <>
        <TagPicker
          onResolveSuggestions={this.onFilterChanged}
          //   getTextFromItem={(item: any) => { return item.name; }}
          getTextFromItem={this.getTextFromItem}
          selectedItems={this.state.selectedItems}
          pickerSuggestionsProps={{
            suggestionsHeaderText: this.state.suggestionsHeaderText,
            noResultsFoundText: this.state.noresultsFoundText
          }}
          onBlur={this._onBlur}
          onEmptyInputFocus={this.onEmptyInputFocus}
          defaultSelectedItems={this.props.defaultSelectedItems || []}
          onChange={this.onItemChanged}
          className={className}
          itemLimit={itemLimit}
          disabled={disabled}
        />

        <Label style={{ color: "#FF0000" }}> {this.state.errorMessage} </Label>
      </>
    );
  }


  private _onBlur = (event: React.FocusEvent<HTMLInputElement | BaseAutoFill>) => {
    event.preventDefault();

    if (this.state.isRequired && this.state.selectedItems.length === 0) {
      this.setState({ errorMessage: 'Por favor indique um valor válido.' });
    } else {
      this.setState({ errorMessage: '' });
    }
  };
  /**
   * Get text from Item
   */
  private getTextFromItem(item: any): string {
    return item.name;
  }

  /**
   * On Selected Item
   */
  private onItemChanged = (selectedItems: { key: string; name: string }[]): void => {

    this.setState({
      selectedItems: selectedItems,
      errorMessage: this.state.isRequired && selectedItems.length == 0 ? "Por favor indique um valor válido." : ''
    });
    this.props.onSelectedItem(selectedItems);
  };

  private onEmptyInputFocus = async ( tagList: ITag[]) => {
    let resolvedSugestions: { key: string; name: string }[] = await this.loadListItems('*');
    //tagList = this.taglist;
    // Filter out the already retrieved items, so that they cannot be selected again
    if (this.state.selectedItems && this.state.selectedItems.length > 0) {
      let filteredSuggestions = [];
      for (const suggestion of resolvedSugestions) {
        const exists = this.state.selectedItems.filter(sItem => sItem.key === suggestion.key);
        if (!exists || exists.length === 0) {
          filteredSuggestions.push(suggestion);
        }
      }
      resolvedSugestions = filteredSuggestions;
    }

    if (resolvedSugestions) {
      this.setState({
        errorMessage: "",
        showError: false
      });

      return resolvedSugestions;
    } else {
      return [];
    }
  };

  /**
   * Filter Change
   */
  private onFilterChanged = async (filterText: string, tagList: ITag[]) => {
    this.setState({ isLoading: true});
    let resolvedSugestions: { key: string; name: string }[] = await this.loadListItems(filterText);
    //tagList = this.taglist;
    // Filter out the already retrieved items, so that they cannot be selected again
    if (this.state.selectedItems && this.state.selectedItems.length > 0) {
      let filteredSuggestions = [];
      for (const suggestion of resolvedSugestions) {
        const exists = this.state.selectedItems.filter(sItem => sItem.key === suggestion.key);
        if (!exists || exists.length === 0) {
          filteredSuggestions.push(suggestion);
        }
      }
      resolvedSugestions = filteredSuggestions;
    }

    if (resolvedSugestions) {
      this.setState({
        errorMessage: "",
        showError: false,
        isLoading:false,
      });

      return resolvedSugestions;
    } else {
      return [];
    }
  };

  /**
   * Function to load List Items
   */
  private loadListItems = async (filterText: string): Promise<{ key: string; name: string }[]> => {
    let { listId, columnInternalName, keyColumnInternalName, webUrl, filterList } = this.props;
    let arrayItems: { key: string; name: string }[] = [];
    let keyColumn: string = keyColumnInternalName || "Id";
    let filter: string = filterList || undefined;

    try {
      let listItems = await this._spservice.getListItemsForListItemPicker(
        filterText,
        listId,
        columnInternalName,
        keyColumn,
        webUrl,
        filter
      );
      // Check if the list had items
      if (listItems.length > 0) {
        for (const item of listItems) {
          arrayItems.push({ key: item[keyColumn], name: item[columnInternalName] });
        }
      }
      if (this.props.removeDuplicates) {
        let newUniqueArray = [];
        newUniqueArray = uniqBy(arrayItems, 'name');
        return newUniqueArray;
      }

      return arrayItems;
    } catch (error) {
      console.log(`Error get Items`, error);
      this.setState({
        showError: true,
        errorMessage: error.message,
        noresultsFoundText: error.message
      });
      return null;
    }
  };
}
