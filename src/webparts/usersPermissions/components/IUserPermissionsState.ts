import { IDropdownOption } from "@fluentui/react";

export interface IPermissionMatrix{
    Object: string;
    Title: string;
    URL: string;
    HasUniquePermissions: string;
    Users: string;
    Type: string;
    Permissions: string;
    GrantedThrough: string;
    LibraryName: string;
}

export interface IUserPermissionsState{
    permissionItemsGrid: IPermissionMatrix[],
    permissionItems: IPermissionMatrix[],
    selectedUserEmail: string;
    libraryNamesDropdownOptions: IDropdownOption[],
    selectedLibraryName: string;
    activeTabName: string;
    siteUrl: string;
    reportFound: boolean;
}