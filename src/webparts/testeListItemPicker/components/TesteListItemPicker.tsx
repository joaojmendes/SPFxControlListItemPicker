import * as React from 'react';
import styles from './TesteListItemPicker.module.scss';
import { ITesteListItemPickerProps } from './ITesteListItemPickerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ListItemPicker} from '../../../controls/listItemPicker/ListItemPicker';

export default class TesteListItemPicker extends React.Component<ITesteListItemPickerProps, {}> {
  public render(): React.ReactElement<ITesteListItemPickerProps> {
    return (
      <div className={ styles.testeListItemPicker }>
             <ListItemPicker columnInternalName="Title" key="ID1" listId="c66ba3d3-dce8-4a54-a90d-3034705515e6"
             context={this.props.context} itemLimit={1} onSelectedItem={(item) =>{ alert(item[0].key)}}
             required={true} removeDuplicates={true}
             />
      </div>
    );
  }
}
