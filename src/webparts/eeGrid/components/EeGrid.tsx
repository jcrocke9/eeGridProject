import * as React from 'react';
import styles from './EeGrid.module.scss';
import { IEeGridProps } from './IEeGridProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { BaseComponent } from 'office-ui-fabric-react/lib/Utilities';
import { DefaultButton } from '@microsoft/sp-webpart-base/node_modules/office-ui-fabric-react/lib/Button';
import { IDetailsList, DetailsList, IColumn, IGroup, IGroupedListProps, CheckboxVisibility, IDetailsGroupRenderProps, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Toggle, IToggleStyles } from 'office-ui-fabric-react/lib/Toggle';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';

import { sp, PagedItemCollection } from '@pnp/sp';

const margin = '0 20px 20px 0';
const controlWrapperClass = mergeStyles({
  display: 'flex',
  alignItems: 'center',
  flexWrap: 'wrap'
});
const toggleStyles: Partial<IToggleStyles> = {
  label: { display: 'inline-block', marginLeft: '10px', marginBottom: '3px' },
  root: { display: 'flex', flexDirection: 'row-reverse', alignItems: 'center', margin: margin }
};

export interface IDetailsListGroupedExampleItem {
  Id?: string;
  Title?: string;
  TheChoiceCol?: string;
}

const _columns: IColumn[] = [
  {
    key: 'Id',
    name: 'ID',
    fieldName: 'Id',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'TheChoiceCol',
    name: 'Choice',
    fieldName: 'TheChoiceCol',
    minWidth: 100,
    maxWidth: 200
  }
];

const _items: IDetailsListGroupedExampleItem[] = [];
export interface IDetailsListGroupedExampleState {
  items?: IDetailsListGroupedExampleItem[];
  groups?: IGroup[];
}

export default class EeGrid extends React.Component<IEeGridProps, IDetailsListGroupedExampleState> {
  constructor(props: IEeGridProps) {
    super(props);

    this.state = {
      items: _items,
      groups: []
    };
  }

  public render() {
    const { items, groups } = this.state;
    return (
      <div data-is-scrollable="true">
        <DetailsList
          items={items}
          groups={groups}
          columns={_columns}
          checkboxVisibility={CheckboxVisibility.hidden}
          selectionMode={SelectionMode.none}
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          ariaLabelForSelectionColumn="Toggle selection"
          groupProps={{
            showEmptyGroups: true,
            getGroupItemLimit: this.getGroupItemLimit
          }}
          onRenderItemColumn={this._onRenderColumn}
        />
      </div>
    );
  }
  private getGroupItemLimit(group: IGroup): number {
    // return 200;
    console.log(group.count);
    if (group) {
      return group.isShowingAll ? group.count : Math.min(group.count, 30);
    } else {
      return this.state.items.length;
    }
  }
  private _onRenderColumn(item: IDetailsListGroupedExampleItem, index: number, column: IColumn) {
    const value = item && column && column.fieldName ? item[column.fieldName as keyof IDetailsListGroupedExampleItem] || '' : '';

    return <div data-is-focusable={true}>{value}</div>;
  }

  private async getItemsNonPaged(): Promise<any[]> {
    let items: IDetailsListGroupedExampleItem[] = await sp.web.lists.getByTitle("Custom").items.select("Id", "Title", "TheChoiceCol").getAll();
    return items;
  }
  private async getItems(): Promise<PagedItemCollection<any[]>> {
    let items: PagedItemCollection<IDetailsListGroupedExampleItem[]> = await sp.web.lists.getByTitle("Custom").items.select("Id", "Title", "TheChoiceCol").getPaged();
    return items;
  }
  private getAllItems(): void {
    let originalAllItems: IDetailsListGroupedExampleItem[];
    let newBuiltAllItems: IDetailsListGroupedExampleItem[] = [];
    let groupDef: IGroup[];
    let arr1: IDetailsListGroupedExampleItem[];
    let arr2: IDetailsListGroupedExampleItem[];
    let arr3: IDetailsListGroupedExampleItem[];
    let sha1: IDetailsListGroupedExampleItem[];
    let sha2: IDetailsListGroupedExampleItem[];
    let sha3: IDetailsListGroupedExampleItem[];
    this.getItemsNonPaged().then((response) => {
      originalAllItems = response;
      
      groupDef = this.state.groups.slice();
      arr1 = originalAllItems.filter(conf => conf.TheChoiceCol === 'Choice 1');
      sha1 = newBuiltAllItems.concat(arr1);
      let groupDef1 = groupDef.concat([{
        key: 'groupred0',
        name: 'Choice 1',
        startIndex: 0,
        count: arr1.length,
        isCollapsed: false
      }]);
      this.setState({
        items: sha1,
        groups: groupDef1
      });
    })
    .then(() => {
      groupDef = this.state.groups.slice();
      arr2 = originalAllItems.filter(conf => conf.TheChoiceCol === 'Choice 2');
      sha2 = sha1.concat(arr2);
      let groupDef2 = groupDef.concat([{
        key: 'groupgreen2',
        name: 'Choice 2',
        startIndex: sha1.length,
        count: arr2.length,
        isCollapsed: false
      }]);
      this.setState({
        items: sha2,
        groups: groupDef2
      });
    })
    .then(() => {
      groupDef = this.state.groups.slice();
      arr3 = originalAllItems.filter(conf => conf.TheChoiceCol === 'Choice 3');
      sha3 = sha2.concat(arr3);
      let groupDef3 = groupDef.concat([{
        key: 'groupblue2',
        name: 'Choice 3',
        startIndex: sha2.length,
        count: arr3.length,
        isCollapsed: false
      }]);
      this.setState({
        items: sha3,
        groups: groupDef3
      });
    });
  }
  private componentDidMount(): void {
    this.getAllItems();
  }
  //#region not in use: filtered by choice in the getItems method
  private async getItems01(choice: string): Promise<PagedItemCollection<any[]>> {
    // TODO: get all items instead of using a filter
    const filter: string = "TheChoiceCol eq '" + choice + "'";
    let items: PagedItemCollection<IDetailsListGroupedExampleItem[]> = await sp.web.lists.getByTitle("Custom").items.filter(filter).select("Id", "Title", "TheChoiceCol").getPaged();
    return items;
  }
  private getAllItems01(): void {
    let shallowArr: IDetailsListGroupedExampleItem[];
    let groupDef: IGroup[];
    this.getItems01('Choice 1').then((response) => {
      // TODO: instead of just getting a slice, use a filter to get the items you want
      shallowArr = this.state.items.slice();
      console.log(shallowArr);
      groupDef = this.state.groups.slice();
      let sha1 = shallowArr.concat(response.results);
      let groupDef1 = groupDef.concat([{
        key: 'groupred0',
        name: 'Choice 1',
        startIndex: shallowArr.length,
        count: response.results.length
      }]);
      this.setState({
        items: sha1,
        groups: groupDef1
      });
    });
    this.getItems01('Choice 2').then((response) => {
      shallowArr = this.state.items.slice();
      groupDef = this.state.groups.slice();
      let sha2 = shallowArr.concat(response.results);
      let groupDef2 = groupDef.concat([{
        key: 'groupgreen2',
        name: 'Choice 2',
        startIndex: shallowArr.length,
        count: response.results.length
      }]);
      this.setState({
        items: sha2,
        groups: groupDef2
      });
    });
    this.getItems01('Choice 3').then((response) => {
      shallowArr = this.state.items.slice();
      groupDef = this.state.groups.slice();
      let sha3 = shallowArr.concat(response.results);
      let groupDef3 = groupDef.concat([{
        key: 'groupblue2',
        name: 'Choice 3',
        startIndex: shallowArr.length,
        count: response.results.length
      }]);
      this.setState({
        items: sha3,
        groups: groupDef3
      });
    });
  }
  //#endregion not in use
}
