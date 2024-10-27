
import * as React from 'react';
// import styles from './UsersPermissions.module.scss';
import type { IUsersPermissionsProps } from './IUsersPermissionsProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/controls/peoplepicker";
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from "../pnpjsConfig";
import { fileFromServerRelativePath, IFile, SPFI, spfi } from "@pnp/sp/presets/all";
import { Dropdown, IDropdownOption, IPersonaProps, Label, Pivot, PivotItem, PrimaryButton } from '@fluentui/react';
import styles from './UsersPermissions.module.scss';
import { GroupOrder, ListView } from '@pnp/spfx-controls-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPermissionMatrix, IUserPermissionsState } from './IUserPermissionsState';


export default class UsersPermissions extends React.Component<IUsersPermissionsProps, IUserPermissionsState> {
  private _sp: SPFI;


  constructor(props: IUsersPermissionsProps) {
    super(props);
    this.state = {
      permissionItems: [],
      permissionItemsGrid: [],
      selectedUserEmail: '',
      libraryNamesDropdownOptions: [],
      selectedLibraryName: '',
      activeTabName: 'User',
      siteUrl: this.props.webpartContext._pageContext._site.absoluteUrl,
      reportFound: false,
      csvGenerationInProgress: false,
      isSiteUrlValid: false,
      updatedReportDate: ''
    }
    this._sp = getSP();
  }

  async componentDidMount(): Promise<void> {
    await this.fetchReport();
    // await this.getPermissionMatrix();
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

  // private getPermissionMatrix = async () => {
  //   const spCache = spfi(this._sp);
  //   // const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + '/Shared Documents/SitePermissionRptV3.csv';
  //   const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + '/Shared Documents/SitePermissionRptV3(8thOct).csv';
  //   const file: IFile = fileFromServerRelativePath(spCache.web, url);
  //   const fileContent = await file.getText();
  //   const csvJSONArr: any[] = this.csvJSON(fileContent);
  //   const permissionItems: IPermissionMatrix[] = csvJSONArr.map((v, i) => {
  //     const object: string = JSON.parse(v['"Object"']);
  //     const url: string = JSON.parse(v['"URL"']);
  //     const title: string = JSON.parse(v['"Title"']);
  //     const isLibrary: boolean = (object.includes('Library') || object.includes('Folder') || object.includes('File')) && !url.includes('Lists');
  //     const libraryName: string = isLibrary ? ((object.includes('Library') && !url.includes('Lists')) ? title : (object.includes('File') ? this.getLibraryNameFromFileFolderUrl(url, true) : this.getLibraryNameFromFileFolderUrl(url, false))) : '';
  //     return {
  //       "Object": JSON.parse(v['"Object"']),
  //       "Title": JSON.parse(v['"Title"']),
  //       "URL": JSON.parse(v['"URL"']),
  //       "HasUniquePermissions": JSON.parse(v['"HasUniquePermissions"']),
  //       "Users": JSON.parse(v['"Users"']),
  //       "Type": JSON.parse(v['"Type"']),
  //       "Permissions": JSON.parse(v['"Permissions"']),
  //       "GrantedThrough": JSON.parse(v['"GrantedThrough"']),
  //       "LibraryName": libraryName
  //     }
  //   });
  //   this.setState({ permissionItems }, () => {
  //     this.setLibraryNames();
  //   });
  // }

  private isValidUrl = (value: string): boolean => {
    let siteurl = value;
    const tenantUrl: string = this.extractTenantUrl(siteurl)!;
    // let arr = tenantUrls?.filter(a => {
    //     if (siteurl.indexOf(a) >= 0)
    //         return a;
    // });
    // return !(arr?.length === 0);
    return siteurl.indexOf(tenantUrl) >= 0;
  }

  private normalizeUrl = (url: string): string => {
    const uri = new URL(url)
    return uri.origin + '/' + uri.pathname.split('/')[1] + '/' + uri.pathname.split('/')[2];
  }

  private extractTenantUrl = (siteUrl: string): string | null => {
    const regex = /^(https:\/\/([^\/]+)\.sharepoint\.com)/; // eslint-disable-line
    const match = siteUrl.match(regex);
    return match ? match[1] : null;
  }

  private isSiteExists = async (url: string): Promise<boolean> => {
    return new Promise<boolean>((resolve, reject) => {
      this.props.webpartContext.spHttpClient.get(url + '/_api/site',
        SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          if (response.status === 404) {
            alert('Entered site does not exist!');
          }
          else {
            response.json().then((responseJSON: any) => {
              // to do
              resolve(true);
            }).catch((error: Error) => {

              resolve(false);
            });
          }
        }).catch((error: Error) => {
          resolve(false);
        });
    })
  }

  private formatDate = (dateString: string): string => {
    const date = new Date(dateString);

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate() + 1).padStart(2, '0');   //added 1 day as powershell script runs every day


    return `${month}/${day}/${year}`;
  };

  private fetchReport = async () => {
    const isValidUrl: boolean = this.isValidUrl(this.state.siteUrl);
    if (!isValidUrl) {
      alert('Please enter valid Site Url');
      this.setState({
        csvGenerationInProgress: false,
        isSiteUrlValid: false
      })
    }
    else {
      const normalizedUrl: string = this.normalizeUrl(this.state.siteUrl);
      this.setState({
        siteUrl: normalizedUrl
      })
      const isSiteExist = await this.isSiteExists(normalizedUrl);
      if (isSiteExist) {
        this.setState({
          isSiteUrlValid: true
        })
        const spCache = spfi(this._sp);
        const listItems = await spCache.web.lists.getByTitle('GenerateCSV').items.filter(`SiteUrl eq '${normalizedUrl}'`).top(1)();
        console.log(listItems);

        if (listItems.length > 0) {
          const updatedReportDate: string = this.formatDate(listItems[0].Modified);
          this.setState({
            csvGenerationInProgress: listItems[0].IsCSVRequested,
            updatedReportDate: updatedReportDate
          })
          const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + `/Shared Documents/AllSitesCSV/${normalizedUrl.split('https://')[1].replaceAll('/', '_') + '.CSV'}`;
          const file: IFile = await fileFromServerRelativePath(spCache.web, url);
          await file.getText().catch((error) => {
            console.log(error);
            this.setState({
              reportFound: false
            })
            alert('Report does not exist. Please click on Generate CSV button.')
          }).then((fileContent) => {
            if (fileContent != undefined) {
              this.setState({
                reportFound: true
              })
              console.log(fileContent);
              const csvJSONArr: any[] = this.csvJSON(fileContent!);
              const permissionItems: IPermissionMatrix[] = csvJSONArr.map((v, i) => {
                const object: string = v['"Object"'] ? JSON.parse(v['"Object"']) : '';
                const url: string = v['"URL"'] ? JSON.parse(v['"URL"']) : '';
                const title: string = v['"Title"'] ? JSON.parse(v['"Title"']) : '';
                const isLibrary: boolean = (object.includes('Library') || object.includes('Folder') || object.includes('File')) && !url.includes('Lists');
                const libraryName: string = isLibrary ? ((object.includes('Library') && !url.includes('Lists')) ? title : (object.includes('File') ? this.getLibraryNameFromFileFolderUrl(url, true) : this.getLibraryNameFromFileFolderUrl(url, false))) : '';
                // const libraryName: string = isLibrary ? ((object.includes('Library') && !url.includes('Lists')) ? title : selectedLibraryName) : '';
                return {
                  "Object": object,
                  "Title": title,
                  "URL": url,
                  "HasUniquePermissions": v['"HasUniquePermissions"'] ? JSON.parse(v['"HasUniquePermissions"']) : '',
                  "Users": v['"Users"'] ? JSON.parse(v['"Users"']) : '',
                  "Type": v['"Type"'] ? JSON.parse(v['"Type"']) : '',
                  "Permissions": v['"Permissions"'] ? JSON.parse(v['"Permissions"']) : '',
                  "GrantedThrough": v['"GrantedThrough"'] ? JSON.parse(v['"GrantedThrough"']) : '',
                  "LibraryName": libraryName
                }
              });
              this.setState({ permissionItems }, () => {
                this.setLibraryNames();
                //alert('Report Fetched successfully');
              });
            }
          });

        }
        else {
          this.setState({
            reportFound: false
          })
          alert('Report does not exist. Please click on Generate CSV button.')
        }
      }
      else {
        this.setState({
          csvGenerationInProgress: false,
          reportFound: false
        })
      }
    }
  }

  private getLibraryNameFromFileFolderUrl = (fileUrl: string, isFile: boolean) => {
    // Use a regular expression to match the URL pattern
    const regex = isFile ? /https:\/\/[^/]+\/sites\/[^/]+\/([^/]+)(\/[^/]+)/ : /https:\/\/[^/]+\/sites\/[^/]+\/([^/]+)(\/[^/]+)?\/?$/;
    const match = fileUrl.match(regex);

    if (match && match[1]) {
      // Decode the library name from URL encoding
      const libraryName: string = decodeURIComponent(match[1]);
      return libraryName == 'Shared Documents' ? 'Documents' : libraryName;
    }

    return ''; // Return empty if not found
  }

  private setLibraryNames = () => {
    //library names logic
    const permissionItems: IPermissionMatrix[] = this.state.permissionItems;
    let libraryNamesDropdownOptions: IDropdownOption[] = permissionItems.filter((v, i, self) => {
      return v.Object.includes('Library') && !v.URL.includes('Lists') && self.map(x => x.Title).indexOf(v.Title) == i;
    }).map((v, i) => {
      return {
        key: v.Title,
        text: v.Title
      }
    })
    libraryNamesDropdownOptions.unshift({ key: 'All', text: 'All' })
    this.setState({ libraryNamesDropdownOptions });
  }

  private searchUsers = async () => {
    if (this.state.selectedUserEmail || this.state.selectedLibraryName) {
      //   const spCache = spfi(this._sp);
      // const url: string = this.props.webpartContext._pageContext._site.serverRelativeUrl + '/Shared Documents/SitePermissionRptV3.csv';
      // const file: IFile = fileFromServerRelativePath(spCache.web, url);
      // const fileContent = await file.getText();
      // const csvJSONArr: any[] = this.csvJSON(fileContent);
      // console.log(csvJSONArr)
      // const permissionItems: IPermissionMatrix[] = csvJSONArr.map((v, i)=> {
      //   return {
      //     "Object": JSON.parse(v['"Object"']),
      //     "Title": JSON.parse(v['"Title"']),
      //     "URL": JSON.parse(v['"URL"']),
      //     "HasUniquePermissions": JSON.parse(v['"HasUniquePermissions"']),
      //     "Users": JSON.parse(v['"Users"']),
      //     "Type":JSON.parse(v['"Type"']),
      //     "Permissions": JSON.parse(v['"Permissions"']),
      //     "GrantedThrough": JSON.parse(v['"GrantedThrough"']),
      //     // "Object": JSON.parse(v["Object"]),
      //     // "Title": JSON.parse(v["Title"]),
      //     // "URL": JSON.parse(v["URL"]),
      //     // "HasUniquePermissions": JSON.parse(v["HasUniquePermissions"]),
      //     // "Users": JSON.parse(v["Users"]),
      //     // "Type":JSON.parse(v["Type"]),
      //     // "Permissions": JSON.parse(v["Permissions"]),
      //     // "GrantedThrough": JSON.parse(v["GrantedThrough"]),
      //   }
      // });
      const permissionItems: IPermissionMatrix[] = this.state.permissionItems;
      // const selectedLibraryName : string = this.state.selectedLibraryName;
      // permissionItems.map((v, i) => {
      //   const isLibrary: boolean = (v.Object.includes('Library') || v.Object.includes('Folder') || v.Object.includes('File')) && !v.URL.includes('Lists');
      //   const libraryName: string = isLibrary ? ((v.Object.includes('Library') && !v.URL.includes('Lists')) ? v.Title :this.state.libraryNamesDropdownOptions.filter((value,index)=>value.text.replace(/[^a-zA-Z ]/g, "") == selectedLibraryName).length> 0 ? selectedLibraryName : '') : '';
      //   // if (selectedLibraryName!= 'All') {
      //   //   v.LibraryName = libraryName;
      //   // }
      //   v.LibraryName = libraryName
      // })
      let filteredItems: IPermissionMatrix[] = permissionItems.filter((v, i) => {
        // return (this.state.selectedUserEmail ? v.Users.split(';').filter((userEmail, i) => userEmail.includes(this.state.selectedUserEmail)).length>0: true) && (!this.state.selectedLibraryName || ((this.state.selectedLibraryName == 'All' && v.Object.includes('Library') && !v.URL.includes('Lists')) || (v.Object.includes('Library') && !v.URL.includes('Lists') && v.Title == this.state.selectedLibraryName)));
        return (this.state.selectedUserEmail ? v.Users.split(';').filter((userEmail, i) => userEmail.includes(this.state.selectedUserEmail)).length > 0 : true) && (!this.state.selectedLibraryName || ((this.state.selectedLibraryName == 'All' && (v.Object.includes('Library') || v.Object.includes('Folder') || v.Object.includes('File')) && !v.URL.includes('Lists')) || ((v.Object.includes('Library') || v.Object.includes('Folder') || v.Object.includes('File')) && !v.URL.includes('Lists') && v.URL.includes(this.state.selectedLibraryName.replace(/[^a-zA-Z ]/g, ""))))) && (v.URL.includes('Lists') ? v.Title != 'CustomConfig' && v.Title != 'CustomAssets' : true) && !v.URL.includes('AllSitesCSV');
        // return (this.state.selectedUserEmail ? v.Users.split(';').filter((userEmail, i) => userEmail.includes('falsettiadm@qauottawa.onmicrosoft.com')).length > 0 : true) && (!this.state.selectedLibraryName || ((this.state.selectedLibraryName == 'All' && (v.Object.includes('Library') || v.Object.includes('Folder') || v.Object.includes('File')) && !v.URL.includes('Lists')) || ((v.Object.includes('Library') || v.Object.includes('Folder') || v.Object.includes('File')) && !v.URL.includes('Lists') && v.URL.includes(this.state.selectedLibraryName.replace(/[^a-zA-Z ]/g, ""))))) && (v.URL.includes('Lists') ? v.Title != 'CustomConfig' && v.Title != 'CustomAssets' : true) && !v.URL.includes('AllSitesCSV');
      })
      this.setState({ permissionItemsGrid: filteredItems });
      // //library names logic
      // let libraryNamesDropdownOptions: IDropdownOption[] = filteredItems.filter((v,i)=> {
      //   return v.Object.includes('Library') && !v.URL.includes('Lists');
      // }).map((v,i)=> {
      //   return {
      //     key: v.Title, 
      //     text: v.Title
      //   }
      // })
      // libraryNamesDropdownOptions.unshift({key: 'All', text: 'All'})
      // this.setState({libraryNamesDropdownOptions});
    }
    else {
      alert(`Please select ${!this.state.selectedUserEmail && !this.state.selectedLibraryName ? 'User/Library' : this.state.selectedUserEmail ? 'Library' : 'User'}`);
    }
  }

  private onUsersPeoplePickerChange = (items: IPersonaProps[]) => {
    if (items.length > 0) {
      const selectedUserEmail: string = items[0].secondaryText!;
      this.setState({
        selectedUserEmail,
        // selectedLibraryName: '', libraryNamesDropdownOptions: [],
        permissionItemsGrid: []
      });
    }
    else {
      this.setState({
        selectedUserEmail: '',
        // selectedLibraryName: '', libraryNamesDropdownOptions: [], 
        permissionItemsGrid: []
      })
    }
  }

  private onDropdownChange = (selectedOption: IDropdownOption) => {
    this.setState({ selectedLibraryName: selectedOption.text, permissionItemsGrid: [] })
  }

  // private onTextEntered = (enteredValue: string) => {
  //   this.setState({
  //     siteUrl: enteredValue,
  //     csvGenerationInProgress: false,
  //     isSiteUrlValid: false,
  //     reportFound: false
  //   })
  // }

  private generateCSV = async () => {
    const spCache = spfi(this._sp);
    const listItems = await spCache.web.lists.getByTitle('GenerateCSV').items.filter(`SiteUrl eq '${this.state.siteUrl}'`).top(1)();
    const objListData: {} = {
      SiteUrl: this.state.siteUrl,
      IsCSVRequested: "true"
    };
    if (listItems.length == 0) {
      await spCache.web.lists.getByTitle("GenerateCSV").items.add(objListData).then((data) => {
        alert('CSV Generation is in process. You will be able to see the updated report after sometime.')
        this.setState({
          csvGenerationInProgress: true
        })
      });
    }
    else {
      await spCache.web.lists.getByTitle("GenerateCSV").items.getById(listItems[0]['ID']).update(objListData).then((data) => {
        alert('CSV Generation is in process. You will be able to see the updated report after sometime.')
        this.setState({
          csvGenerationInProgress: true
        })
      });
    }
  }

  private stringToArray = (str: string) => {
    let arr = [''];
    let j = 0;

    for (let i = 0; i < str.length; i++) {
      if (str.charAt(i) == " ") {
        j++;
        arr.push('');
      } else {
        arr[j] += str.charAt(i);
      }
    }
    return arr;
  }

  // private onPivotClick = (activeTabName: string) => {
  //   this.setState({ activeTabName });
  // }
  private onPivotClick = (pivotItem: PivotItem) => {
    this.setState({ activeTabName: pivotItem.props.headerText!, permissionItemsGrid: [], selectedLibraryName: '', selectedUserEmail: '' });
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

    const columns: any[] = [
      {
        name: 'Object',
        displayName: 'Object',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true,
        sorting: true,
        render: (item: IPermissionMatrix) => {
          return item.Object
        }
      },
      {
        name: 'Title',
        displayName: 'Title',
        minWidth: 50,
        maxWidth: 100,
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
        name: 'Type',
        displayName: 'Type',
        minWidth: 100,
        maxWidth: 350,
        isResizable: true,
        sorting: true,
        isVisible: false,
        render: (item: IPermissionMatrix) => {
          return item.Type
        }
      },
      {
        name: 'Permissions',
        displayName: 'Permissions',
        minWidth: 50,
        maxWidth: 50,
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
        maxWidth: 150,
        isResizable: true,
        sorting: true,
        render: (item: IPermissionMatrix) => {
          return item.GrantedThrough
        }
      },
      {
        name: 'Users',
        displayName: '',
        minWidth: 200,
        maxWidth: 350,
        isVisible: this.state.activeTabName == 'Library',
        render: (item: IPermissionMatrix) => {
          return <span title={item.Users}>{item.Users}</span>
        }
      },
      {
        name: 'URL',
        minWidth: 100,
        maxWidth: 350,
        isResizable: true,
        render: (item: IPermissionMatrix) => {
          return <a title={item.URL}>{item.URL}</a>
        }
      }
    ].filter(column => column.isVisible !== false); // Filter out non-visible columns

    const exportToExcel = () => {
      const header = columns.map(col => col.name).join(',');
      const listviewItems: any[] = this.state.permissionItemsGrid;
      const rows = listviewItems.map(item => columns.map(col => item[col.name]).join(',')).join('\n');
      const csvContent = `${header}\n${rows}`;

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.setAttribute('download', 'export.csv');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    };

    return (
      <>
        <h2>SPO Permissions Report</h2>
        <div className={styles['fl-grid']}>
          <>
            {/* <div className={`${styles['fl-span8']}`}>
              <TextField
                label="Site Url"
                placeholder='Site url e.g. https://qauottawa.sharepoint.com/sites/M365LP'
                onChange={(event, newValue) => this.onTextEntered(newValue!)}
              />
            </div>
            <div className={styles['fl-span2']}>
              <PrimaryButton style={{ marginTop: '27px' }} text='Fetch Report' onClick={this.fetchReport} />
            </div> */}
            <div className={styles['fl-span2']}>
              <PrimaryButton style={{ marginTop: '27px' }} text='Generate CSV' onClick={this.generateCSV}
                disabled={!this.state.isSiteUrlValid || this.state.csvGenerationInProgress}
              />
            </div>
            <Label
              className={`${styles['fl-span12']} ${this.state.csvGenerationInProgress ? '' : styles.hidden}`}
            >CSV Generation is in process. You will be able to see the updated report on {this.state.updatedReportDate}.</Label>
          </>
        </div>
        {/* <div className={styles['fl-grid']}>
          <>
            <div className={`${styles['fl-span4']}`}>
              <ComboBox
                label="Sites"
                options={[{ key: 'Governance', text: 'Governance' }, { key: 'M365LP', text: 'M365LP' }, { key: 'NewHomeSite', text: 'NewHomeSite' }]}
                allowFreeInput
                autoComplete="on"
                unselectable='on'
              />
            </div>
          </>
        </div> */}
        <Pivot onLinkClick={(item) => this.onPivotClick(item!)}>
          <PivotItem headerText='User'
          // onClick={() => this.onPivotClick('User')} 
          />
          <PivotItem headerText='Library'
          // onClick={() => this.onPivotClick('Library')}
          />
        </Pivot>
        <div className={styles['fl-grid']}>
          <>
            <div className={`${styles['fl-span4']} ${this.state.activeTabName == 'Library' ? styles.hidden : ''}`}>
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
                defaultSelectedUsers={this.stringToArray(this.state.selectedUserEmail)}
              />
            </div>
            <div className={`${styles['fl-span4']} ${this.state.activeTabName == 'User' ? styles.hidden : ''}`}>
              <Dropdown
                options={this.state.libraryNamesDropdownOptions}
                label='Libraries'
                onChange={(ev, option) => this.onDropdownChange(option!)}
                unselectable="on"
              />
            </div>
            <div className={styles['fl-span2']}>
              <PrimaryButton style={{ marginTop: '27px' }} text='Search' onClick={this.searchUsers}
                disabled={!this.state.reportFound}
              />
            </div>
          </>
          <div className={styles['fl-span6']}></div>
          {/* <div className={styles['fl-span4']}></div> */}
          <div className={styles['fl-span6']}>
            <PrimaryButton style={{ marginTop: '27px' }} text='Export to Excel' onClick={exportToExcel}
              disabled={this.state.permissionItemsGrid.length == 0}
            />
          </div>
          <div className={styles['fl-span12']}>
            <ListView
              items={this.state.permissionItemsGrid}
              // groupByFields={[{name: 'Users', order: GroupOrder.ascending }]}
              groupByFields={this.state.activeTabName == 'Library' ? [{ name: 'LibraryName', order: GroupOrder.ascending }] : []}
              viewFields={columns}
              showFilter
            />
          </div>
        </div>
      </>
    );
  }
}
