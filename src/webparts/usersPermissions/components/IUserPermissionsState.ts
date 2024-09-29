export interface IPermissionMatrix{
    Object: string;
    Title: string;
    URL: string;
    HasUniquePermissions: string;
    Users: string;
    Type: string;
    Permissions: string;
    GrantedThrough: string;
}

export interface IUserPermissionsState{
    permissionItems: IPermissionMatrix[],
    selectedUserEmail: string;
}