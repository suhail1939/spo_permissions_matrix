import * as React from 'react';
// import styles from './UsersPermissions.module.scss';
import type { IUsersPermissionsProps } from './IUsersPermissionsProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/controls/peoplepicker";
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from "../pnpjsConfig";
import { fileFromServerRelativePath, IFile, SPFI, spfi } from "@pnp/sp/presets/all";
import { Dropdown, IDropdownOption, IPersonaProps, PrimaryButton } from '@fluentui/react';
import styles from './UsersPermissions.module.scss';
import { IViewField, ListView } from '@pnp/spfx-controls-react';
import { IPermissionMatrix, IUserPermissionsState } from './IUserPermissionsState';

const columns: IViewField[] = [
  {
    name: 'Object',
    displayName: 'Object',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    sorting: true,
    render: (item: IPermissionMatrix) => {
      return item.Object
    }
    // render: (item?: IOCSRData) => (
    //   <span className={styles.hoverable} onClick={() => this._viewDetails(item)}>
    //     {item?.OCSSRD}
    //   </span>
    // ),
  },
  {
    name: 'Title',
    displayName: 'Title',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    sorting: true,
    render: (item: IPermissionMatrix) => {
      return item.Title
    }
    // render: (item?: IOCSRData) => (
    //   <span className={styles.hoverable} onClick={() => this._viewDetails(item)}>
    //     {item?.SrSubject}
    //   </span>
    // ),
  },
  {
    name: 'URL',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    render: (item: IPermissionMatrix) => {
      return item.URL
    }
    // render: (item: IOCSRData) => (
    //   <IconButton
    //     iconProps={{ iconName: 'More' }}
    //     title="More actions"
    //     ariaLabel="More actions"
    //     onClick={(e: React.MouseEvent<HTMLElement>) => this._showContextualMenu(e, item)}
    //   />
    // ),
  },
  {
    name: 'Type',
    displayName: 'Type',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    sorting: true,
    render: (item: IPermissionMatrix) => {
      return item.Type
    }
  },
  {
    name: 'Permissions',
    displayName: 'Permissions',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    sorting: true,
    render: (item: IPermissionMatrix) => {
      return item.Permissions
    }
  },
  {
    name: 'GrantedThrough',
    displayName: 'Granted Through',
    minWidth: 100,
    maxWidth: 350,
    isResizable: true,
    sorting: true,
    render: (item: IPermissionMatrix) => {
      return item.GrantedThrough
    }
  }
];
export default class UsersPermissions extends React.Component<IUsersPermissionsProps, IUserPermissionsState> {
  private _sp: SPFI;


  constructor(props: IUsersPermissionsProps) {
    super(props);
    this.state = {
      permissionItems: [],
      selectedUserEmail: '',
      libraryNamesDropdownOptions: [],
      selectedLibraryName: ''
    }
    this._sp = getSP();
  }

  async componentDidMount(): Promise<void> {
    // spCache.web.lists.getByTitle('Documents').items.select('ID')().then(items => {
      //   if (items.length > 0) {
        //     console.log('Items found: {0}', items.length);
        //   }
        // }).catch((error) => {
          //   console.log('Error occured: {0}', error)
          // })
    // const spCache = spfi(this._sp);
    // const url: string = "/teams/TestSuhail/Shared Documents/SPOPermissionsRpt.csv";
    // //const blob: Blob = await spCache.web.getFileByServerRelativePath(url).getBlob();
    // const file: IFile = fileFromServerRelativePath(spCache.web, url);
    // const fileContent = await file.getText();
    // const csvtojson = this.csvJSON(fileContent);
    // console.log(csvtojson)
    // console.log(blob);
  }

  private csvJSON(csvText: string) {
    let lines: any[] = [];
    const linesArray = csvText.split('\n');
    // for trimming and deleting extra space 
    linesArray.forEach((e: any) => {
      const row = e.replace(/[\s]+[,]+|[,]+[\s]+/g, ',').trim();
      lines.push(row);
    });
    // for removing empty record
    lines.splice(lines.length - 1, 1);
    const result = [];
    const headers = lines[0].split(",");

    for (let i = 1; i < lines.length; i++) {

      const obj: any = {};
      const currentline = lines[i].split(",");

      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = currentline[j];
      }
      result.push(obj);
    }
    //return result; //JavaScript object
    // return JSON.stringify(result); //JSON
    return result;
  }

  private searchUsers =async () => {
    if (this.state.selectedUserEmail) {
      const spCache = spfi(this._sp);
    // const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + '/Shared Documents/SitePermissionRptV2.csv';
    const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + '/Shared Documents/SitePermissionRptV3.csv';
    //const blob: Blob = await spCache.web.getFileByServerRelativePath(url).getBlob();
    const file: IFile = fileFromServerRelativePath(spCache.web, url);
    const fileContent = await file.getText();
    const csvJSONArr: any[] = this.csvJSON(fileContent);
    console.log(csvJSONArr)
    const permissionItems: IPermissionMatrix[] = csvJSONArr.map((v, i)=> {
      return {
        "Object": JSON.parse(v['"Object"']),
        "Title": JSON.parse(v['"Title"']),
        "URL": JSON.parse(v['"URL"']),
        "HasUniquePermissions": JSON.parse(v['"HasUniquePermissions"']),
        "Users": JSON.parse(v['"Users"']),
        "Type":JSON.parse(v['"Type"']),
        "Permissions": JSON.parse(v['"Permissions"']),
        "GrantedThrough": JSON.parse(v['"GrantedThrough"']),
        // "Object": JSON.parse(v["Object"]),
        // "Title": JSON.parse(v["Title"]),
        // "URL": JSON.parse(v["URL"]),
        // "HasUniquePermissions": JSON.parse(v["HasUniquePermissions"]),
        // "Users": JSON.parse(v["Users"]),
        // "Type":JSON.parse(v["Type"]),
        // "Permissions": JSON.parse(v["Permissions"]),
        // "GrantedThrough": JSON.parse(v["GrantedThrough"]),
      }
    });
    let filteredItems: IPermissionMatrix[] = permissionItems.filter((v, i)=> {
      // return v.Users.split(';').filter((userEmail, i) => userEmail.includes(this.state.selectedUserEmail)).length>0;
      return v.Users.split(';').filter((userEmail, i) => userEmail.includes(this.state.selectedUserEmail)).length>0 && (!this.state.selectedLibraryName || ((this.state.selectedLibraryName == 'All' && v.Object.includes('Library') && !v.URL.includes('Lists')) || (v.Object.includes('Library') && !v.URL.includes('Lists') && v.Title == this.state.selectedLibraryName)));
      // return v.Users.split(';').filter((userEmail, i) => userEmail.includes('falsettiadm@qauottawa.onmicrosoft.com')).length>0 && (!this.state.selectedLibraryName || ((this.state.selectedLibraryName == 'All' && v.Object.includes('Library') && !v.URL.includes('Lists')) || (v.Object.includes('Library') && !v.URL.includes('Lists') && v.Title == this.state.selectedLibraryName)));
    })
    this.setState({permissionItems: filteredItems});
    //library names logic
    let libraryNamesDropdownOptions: IDropdownOption[] = filteredItems.filter((v,i)=> {
      return v.Object.includes('Library') && !v.URL.includes('Lists');
    }).map((v,i)=> {
      return {
        key: v.Title, 
        text: v.Title
      }
    })
    libraryNamesDropdownOptions.unshift({key: 'All', text: 'All'})
    this.setState({libraryNamesDropdownOptions});
    }
    else{
      alert('Please select User');
    }
  }

  private onUsersPeoplePickerChange = (items: IPersonaProps[]) => {
    if(items.length > 0){
      const selectedUserEmail: string  = items[0].secondaryText!;
      this.setState({selectedUserEmail, selectedLibraryName: '', libraryNamesDropdownOptions: [],permissionItems:[]});
    }
    else{
      this.setState({ selectedUserEmail: '', selectedLibraryName: '', libraryNamesDropdownOptions: [], permissionItems: []})
    }
  }

  private onDropdownChange = (selectedOption: IDropdownOption) => {
    this.setState({selectedLibraryName: selectedOption.text, permissionItems: []})
  }

//   private csvToJson(csvString: string) {
//     const rows = csvString
//         .split("\n");

//     const headers = rows[0]
//         .split(",");

//     const jsonData = [];
//     for (let i = 1; i < rows.length; i++) {

//         const values = rows[i]
//             .split(",");

//         const obj:any = {};

//         for (let j = 0; j < headers.length; j++) {

//             const key = headers[j]
//                 .trim();
//             const value = values[j]
//                 .trim();

//             obj[key] = value;
//         }

//         jsonData.push(obj);
//     }
//     return JSON.stringify(jsonData);
// }
  public render(): React.ReactElement<IUsersPermissionsProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <>
        <h2>Check Users Permissions</h2>
        <div className={styles['fl-grid']}>
          <div>
            <div className={styles['fl-span4']}>
              <PeoplePicker
                // styles={{root:{width: 250}}}
                context={this.props.webpartContext}
                titleText="Users"
                personSelectionLimit={1}
                groupName={''} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                // required={this.state.formData.Mandatory}
                // disabled={this.props.isViewMode}
                // onChange={(items) => this._getPeoplePickerItems(items, objField.FieldName)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
              // defaultSelectedUsers={defaultSelectedUsersArr}
              // errorMessage={this.state.formErrMsg.ErrMsg}
              webAbsoluteUrl={this.props.webpartContext._pageContext._site.absoluteUrl}
              onChange={(items) => this.onUsersPeoplePickerChange(items)}
              />
            </div>
            <div className={styles['fl-span4']}>
            <Dropdown 
              options={this.state.libraryNamesDropdownOptions}
              label='Libraries'
              onChange={(ev,option) => this.onDropdownChange(option!)}
              />
            </div>
            <div className={styles['fl-span4']}>
              <PrimaryButton text='Search' onClick={this.searchUsers} />
            </div>
          </div>
          <div className={styles['fl-span12']}>
            <ListView
              items={this.state.permissionItems}
              groupByFields={[]}
              viewFields={columns}
            />
          </div>
        </div>
      </>
    );
  }
}
